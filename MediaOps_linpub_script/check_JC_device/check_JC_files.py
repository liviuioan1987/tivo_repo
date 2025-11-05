import os
import re
import csv
import json
import requests
import datetime
import time
import xml.etree.ElementTree as ET
from typing import Set, Optional

# ===== CONFIGURATION =====
# Edit these variables instead of passing as parameters
hostname = "pdk01.ts1"
partner_ids_file = "staging_MSO_NetflixTakedown.txt"

# HTTP settings
HTTP_TIMEOUT = 10.0  # seconds
RETRY_DELAY = 1.0    # seconds between retries
MAX_RETRIES = 3

# service settings (match original script behavior)
MAX = 400
PAGE_SIZE = 50

# ===== Helper functions =====


def safe_request(url: str, params: dict = None) -> Optional[str]:
    """Perform HTTP GET with simple retry logic. Returns text or None on repeated failure."""
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            resp = requests.get(url, params=params, timeout=HTTP_TIMEOUT)
            resp.raise_for_status()
            return resp.text
        except Exception as e:
            print(f"Request error (attempt {attempt}/{MAX_RETRIES}) for URL: {url} -> {e}")
            if attempt < MAX_RETRIES:
                time.sleep(RETRY_DELAY)
    print(f"Failed to fetch URL after {MAX_RETRIES} attempts: {url}")
    return None


def _tag_without_ns(tag: str) -> str:
    """Strip namespace from an Element tag if present."""
    if '}' in tag:
        return tag.split('}', 1)[1]
    return tag


def find_first_text_for_tag(xml_text: str, tag_fragment_lower: str) -> Optional[str]:
    """
    Parse XML and return the first element text where the tag (without namespace)
    contains tag_fragment_lower (case-insensitive). Returns None if not found.
    """
    if not xml_text:
        return None
    try:
        root = ET.fromstring(xml_text)
    except ET.ParseError:
        # Try to wrap with a root if there are multiple top-level elements
        try:
            wrapped = "<root>" + xml_text + "</root>"
            root = ET.fromstring(wrapped)
        except ET.ParseError:
            return None

    for elem in root.iter():
        tag = _tag_without_ns(elem.tag).lower()
        if tag_fragment_lower in tag:
            if elem.text and elem.text.strip() != "":
                return elem.text.strip()
    return None


def find_all_texts_for_tag(xml_text: str, tag_fragment_lower: str) -> Set[str]:
    """Return set of all texts for elements whose tag contains the fragment (case-insensitive)."""
    results = set()
    if not xml_text:
        return results
    try:
        root = ET.fromstring(xml_text)
    except ET.ParseError:
        try:
            wrapped = "<root>" + xml_text + "</root>"
            root = ET.fromstring(wrapped)
        except ET.ParseError:
            return results

    for elem in root.iter():
        tag = _tag_without_ns(elem.tag).lower()
        if tag_fragment_lower in tag:
            if elem.text and elem.text.strip() != "":
                results.add(elem.text.strip())
    return results


def sanitize_mso_name(name: str) -> str:
    """
    Make MSO name filesystem-friendly and remove trailing patterns like:
      -YYYYMMDD_HHMMSS.crt
    Example: 'operator-20141215_140231.crt' -> 'operator'
    """
    if not name:
        return name or ""
    # Remove the trailing pattern -YYYYMMDD_HHMMSS.crt (if present)
    name = re.sub(r'-\d{8}_\d{6}\.crt$', '', name)
    # Replace any characters not safe for filenames with underscores
    safe_name = "".join(c if (c.isalnum() or c in "._-") else "_" for c in name).strip("_")
    return safe_name or name


# ===== Core logic: fetching and CSV generation =====


def ensure_out_dir(hostname: str) -> str:
    today = datetime.date.today().isoformat()
    folder_name = f"{today}_{hostname}"
    os.makedirs(folder_name, exist_ok=True)
    return folder_name


def fetch_mso_name(host: str, partner_id: str) -> str:
    """
    Fetch MSO name using mind99 endpoint. Returns a safe filename-friendly name (or partner_id fallback).
    """
    url = f"http://{host}.tivo.com:8085/mind/mind99"
    params = {
        "type": "partnerInfoSearch",
        "noLimit": "true",
        "partnerId": partner_id,
        "levelOfDetail": "high",
    }
    text = safe_request(url, params=params)
    name = find_first_text_for_tag(text, "name")
    if name:
        return sanitize_mso_name(name)
    return partner_id


def collect_headend_ids(host: str, partner_id: str) -> Set[str]:
    """
    Iterate offsets and collect unique headendId values from the mind39 endpoint.
    """
    headends = set()
    url = f"http://{host}.tivo.com:8085/mind/mind39"
    for offset in range(0, MAX, PAGE_SIZE):
        params = {
            "type": "serviceConfigurationSearch",
            "serviceType": "directTune",
            "count": str(PAGE_SIZE),
            "partnerId": partner_id,
            "offset": str(offset),
        }
        text = safe_request(url, params=params)
        if text is None:
            continue
        found = find_all_texts_for_tag(text, "headendid")
        if found:
            headends.update(found)
    return headends


def fetch_applicable_and_stb(host: str, partner_id: str, headend_id: str) -> (Optional[str], Optional[str]):
    """
    Fetch applicableDeviceType and any 'stb' containing element text for a given partner/headend.
    Returns (applicableDeviceType, stbDeviceType)
    """
    url = f"http://{host}.tivo.com:8085/mind/mind39"
    params = {
        "type": "serviceConfigurationSearch",
        "serviceType": "directTune",
        "count": str(PAGE_SIZE),
        "partnerId": partner_id,
        "headendId": headend_id,
    }
    text = safe_request(url, params=params)
    if text is None:
        return None, None

    applicable = find_first_text_for_tag(text, "applicabledevicetype")
    # find any tag that contains 'stb'
    stb = find_first_text_for_tag(text, "stb")
    return applicable, stb


def process_partners(host: str, partner_file: str, out_dir: str):
    """
    Main function that processes each partner id from the file and writes CSV per MSO name.
    """
    # Read partner IDs
    if not os.path.exists(partner_file):
        raise FileNotFoundError(f"partner ids file not found: {partner_file}")

    with open(partner_file, "r", encoding="utf-8") as f:
        partner_ids = [line.strip() for line in f if line.strip()]

    print(f"Processing {len(partner_ids)} partner ids from {partner_file} ...")

    for pid in partner_ids:
        print(f"\n--- Partner ID: {pid}")
        mso_name = fetch_mso_name(host, pid)
        csv_filename = os.path.join(out_dir, f"{host}_JC_applicableDeviceType_{mso_name}.csv")
        print(f"MSO Name: {mso_name}; CSV: {csv_filename}")

        # Write header
        with open(csv_filename, "w", newline="", encoding="utf-8") as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(["MSO Name", "partnerId", "HeadendId", "Linux"])

            headend_ids = collect_headend_ids(host, pid)
            print(f"Found {len(headend_ids)} headendIds for partner {pid}")

            for hid in headend_ids:
                applicable, stb_dev = fetch_applicable_and_stb(host, pid, hid)
                print(f"  headend {hid}: applicableDeviceType={applicable!r}, stb={stb_dev!r}")
                # original bash logic:
                # if applicableDeviceType is empty OR stbDeviceType is not empty => isSTB = "Y"
                if (not applicable) or (stb_dev and stb_dev.strip() != ""):
                    isSTB = "Y"
                else:
                    isSTB = "N"

                writer.writerow([mso_name, pid, hid, isSTB])

    print("\nCSV generation complete.")


# ===== CSV -> JSON transformation =====


def csvs_to_json(out_dir: str, output_filename: Optional[str] = None) -> str:
    """
    Iterate each CSV file in out_dir that matches the pattern {hostname}_JC_applicableDeviceType_*.csv,
    aggregate data into one JSON object where:
      key = partnerId (string)
      value = list of HeadendId values where Linux == "N"
    Save JSON to a single file in out_dir and return the JSON filepath.
    """
    aggregated = {}

    for fname in os.listdir(out_dir):
        if not fname.startswith(f"{hostname}_JC_applicableDeviceType_") or not fname.lower().endswith(".csv"):
            continue
        fpath = os.path.join(out_dir, fname)
        print(f"Reading CSV: {fpath}")
        with open(fpath, newline="", encoding="utf-8") as csvfile:
            reader = csv.reader(csvfile)
            header = next(reader, None)
            if header is None:
                continue
            # normalize indices
            header_lower = [h.strip().lower() for h in header]
            try:
                partner_idx = header_lower.index("partnerid")
                headend_idx = header_lower.index("headendid")
                linux_idx = header_lower.index("linux")
            except ValueError:
                # If header doesn't match exactly, attempt fallback positions:
                # The original script wrote: MSO Name,partnerId,HeadendId,Linux
                partner_idx = 1
                headend_idx = 2
                linux_idx = 3

            for row in reader:
                if len(row) <= max(partner_idx, headend_idx, linux_idx):
                    continue
                partner_key = row[partner_idx].strip()
                headend_val = row[headend_idx].strip()
                linux_val = row[linux_idx].strip()
                if linux_val.upper() == "N":
                    aggregated.setdefault(partner_key, set()).add(headend_val)

    # convert sets to lists
    aggregated_lists = {k: sorted(list(v)) for k, v in aggregated.items()}

    if output_filename is None:
        output_filename = os.path.join(out_dir, "all_operators_blacklist_headends.json")

    with open(output_filename, "w", encoding="utf-8") as jf:
        json.dump(aggregated_lists, jf, indent=2)

    print(f"JSON written to {output_filename}")
    return output_filename


# ===== Main runnable script =====

def main():
    out_dir = ensure_out_dir(hostname)
    process_partners(hostname, partner_ids_file, out_dir)
    csvs_to_json(out_dir)


if __name__ == "__main__":
    main()