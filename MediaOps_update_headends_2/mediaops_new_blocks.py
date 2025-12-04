#!/usr/bin/env python3
"""
add_vividtv_clean_pretty_rewrite_v2.py

Reads input XMLs and generates new XML files with the structure:

<partnerServiceConfiguration>
  <headend>
    <headendId>...</headendId>
    <partnerConfigurationGroup>...</partnerConfigurationGroup>
    <partnerId>...</partnerId>
    <configuration>...</configuration>
  </headend>
</partnerServiceConfiguration>

If output exceeds 3 MB, it is split into multiple files preserving full <headend> blocks.
"""

import os
import datetime
import xml.etree.ElementTree as ET
import shutil
import sys

# --- Groups (same sets) ---
linux_groups = {
    "DG_anetflix_takedown_c",
    "DG_tier_mediacom_ipsbhdstb",
    "DG_tier_mediacom_ipvod_sbhdstb-c",
    "DG_vp_mediacom",
    "DG_vp_mediacomvod"
}

managed_android_groups = {
    "DG_tier_mediacom_ipvod_managed_android",
    "DG_tier_mediacom_ipvod_managed_android-c"
}

unmanaged_groups = {
    "DG_tier_mediacom_ipvod-unmanaged-c",
    "DG_tier_mediacom_ipvod_unmanaged"
}

MAX_MB = 3
MAX_BYTES = MAX_MB * 1024 * 1024

# --- Build configuration block ---
def build_block_element(group, device_types):
    config = ET.Element("configuration")

    for dt in device_types:
        adt = ET.SubElement(config, "applicableDeviceType")
        adt.text = dt

    app = ET.SubElement(config, "application")
    ET.SubElement(app, "description").text = "VIVIDTV jump channel"
    ET.SubElement(app, "name").text = "VIVIDTV"
    ref = ET.SubElement(app, "reference")
    ET.SubElement(ref, "uiDestinationId").text = "tivo:ud.1008711"

    app_param = ET.SubElement(config, "applicationParameter")
    ET.SubElement(app_param, "reference").text = (
        "tivo:cg.cp./root/ADULT SUBSCRIPTION/VIVIDTV SUBSCRIPTION"
    )

    assoc = ET.SubElement(config, "association")
    ET.SubElement(assoc, "description").text = "VIVIDTV SUBSCRIPTION."
    ET.SubElement(assoc, "shortName").text = "VIVIDTV"
    ET.SubElement(assoc, "virtualChannelNumber").text = "481"

    ET.SubElement(config, "autoStart").text = "true"
    ET.SubElement(config, "delayed").text = "5000"
    ET.SubElement(config, "name").text = "VIVIDTV"
    ET.SubElement(config, "version").text = "1"

    return config

# --- Whitespace normalization ---
def normalize_whitespace(elem):
    if elem.text is not None:
        txt = elem.text.strip()
        elem.text = txt if txt else None
    elem.tail = None
    for child in list(elem):
        normalize_whitespace(child)

# --- Indent XML nicely ---
def safe_indent(elem, level=0, indent_str="  "):
    if hasattr(ET, "indent"):
        ET.indent(elem, space=indent_str)
        return

    i = "\n" + level * indent_str
    if len(elem):
        if not elem.text or not elem.text.strip():
            elem.text = i + indent_str
        for child in elem:
            safe_indent(child, level + 1, indent_str)
        if not child.tail or not child.tail.strip():
            child.tail = i
    else:
        if level and (not elem.tail or not elem.tail.strip()):
            elem.tail = i

# --- Split large output files ---
def split_if_large(filepath, output_folder, base_name):
    size = os.path.getsize(filepath)
    if size <= MAX_BYTES:
        return

    print(f"Splitting {base_name} ({size/1024/1024:.2f} MB)...")

    tree = ET.parse(filepath)
    root = tree.getroot()

    parts = []
    current_part = ET.Element("partnerServiceConfiguration")
    current_size = 0

    for headend_block in list(root):
        block_bytes = ET.tostring(headend_block, encoding="utf-8")
        if current_size + len(block_bytes) > MAX_BYTES and len(current_part):
            parts.append(current_part)
            current_part = ET.Element("partnerServiceConfiguration")
            current_size = 0
        current_part.append(headend_block)
        current_size += len(block_bytes)

    if len(current_part):
        parts.append(current_part)

    for i, part in enumerate(parts, start=1):
        outpath = os.path.join(
            output_folder, f"{os.path.splitext(base_name)[0]}_part{i}.xml"
        )
        normalize_whitespace(part)
        safe_indent(part)
        ET.ElementTree(part).write(outpath, encoding="utf-8", xml_declaration=True)
        print(f"  -> Wrote {outpath}")

# --- Main processing ---
def process_files(input_folder):
    if not os.path.isdir(input_folder):
        raise SystemExit(f"Input folder does not exist: {input_folder!r}")

    today = datetime.date.today().strftime("%Y%m%d")
    output_folder = f"/Users/liviu.gherasim/Downloads/New_folder_{today}"
    os.makedirs(output_folder, exist_ok=True)

    for fname in sorted(os.listdir(input_folder)):
        if not fname.lower().endswith(".xml"):
            continue

        inpath = os.path.join(input_folder, fname)
        try:
            tree = ET.parse(inpath)
            root = tree.getroot()
        except ET.ParseError as e:
            print(f"Skipping {fname}: parse error: {e}", file=sys.stderr)
            continue

        new_root = ET.Element("partnerServiceConfiguration")
        modified = False

        # --- Iterate all headends ---
        for headend in root.findall(".//headend"):
            headend_id_elem = headend.find("headendId")
            partner_id_elem = headend.find("partnerId")
            if headend_id_elem is None or partner_id_elem is None:
                continue

            # --- Check all partnerConfigurationGroups in headend ---
            for group_elem in headend.findall("partnerConfigurationGroup"):
                if group_elem.text is None:
                    continue
                group = group_elem.text.strip()

                if group in linux_groups:
                    dev_types = ["stb"]
                elif group in managed_android_groups:
                    dev_types = ["managedAndroidTv"]
                elif group in unmanaged_groups:
                    dev_types = ["androidTv", "androidTvSony", "appleTv", "fireTv"]
                else:
                    continue

                # --- Build new headend block ---
                headend_block = ET.Element("headend")

                ET.SubElement(headend_block, "headendId").text = headend_id_elem.text
                ET.SubElement(headend_block, "partnerConfigurationGroup").text = group
                ET.SubElement(headend_block, "partnerId").text = partner_id_elem.text

                # Add <configuration> block
                headend_block.append(build_block_element(group, dev_types))

                # Add to root
                new_root.append(headend_block)
                modified = True

        # --- Save the output if we have matches ---
        if not modified:
            print(f"No matches in {fname}, skipping output.")
            continue

        normalize_whitespace(new_root)
        safe_indent(new_root)

        outpath = os.path.join(output_folder, fname)
        ET.ElementTree(new_root).write(outpath, encoding="utf-8", xml_declaration=True)

        split_if_large(outpath, output_folder, fname)

        print(f"Processed: {fname} -> {output_folder}")

    print(f"\nDone. Files saved to: {output_folder}")

# --- Run ---
if __name__ == "__main__":
    input_folder = "/Users/liviu.gherasim/Downloads/tivo_pt.4177_prod_directune_Backup_20251110"
    process_files(input_folder)
