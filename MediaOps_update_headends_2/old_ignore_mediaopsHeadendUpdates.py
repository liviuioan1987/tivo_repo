#!/usr/bin/env python3
import os
import copy
import glob
import json
import xml.etree.ElementTree as ET
from collections import defaultdict
import pandas as pd
import xml.dom.minidom
from datetime import date


class Mediaops_headends():
    def __init__(self, operators_main_file, source_folder, save_directory):
        self.operators_main_file = operators_main_file
        self.source_folder = source_folder
        self.directory = save_directory
        self.operator_files_values = {}
        self.collected_headends = defaultdict(list)

        # ✅ Optional: specify operators that may have empty partnerConfigurationGroup
        self.operators_with_empty_partnerconfig = [
            "tivo:pt.3731",
            "tivo:pt.5221",
            "tivo:pt.5221",
            "tivo:pt.5161",
            "tivo:pt.4812",
            "tivo:pt.3640",
            "tivo:pt.4222",
            "tivo:pt.5174"
        ]

    def read_operators_rules(self):
        df = pd.read_excel(self.operators_main_file)
        df["Operator ID"] = df["Operator ID"].ffill()

        result = defaultdict(list)
        for _, row in df.iterrows():
            operator = row["Operator ID"]
            entry = {
                "OperatorName": row["Operator Name"],
                "OldpartnerConfigurationGroup": row["Old partnerConfigurationGroup"],
                "NewpartnerConfigurationGroup": row["New partnerConfigurationGroup"],
                "DeleteConfiguration": [x.strip().lower() for x in str(row["Delete Configuration"]).split(",")]
                if pd.notna(row["Delete Configuration"]) else []
            }
            result[operator].append(entry)

        self.operator_files_values = dict(result)
        return self.operator_files_values

    def collect_updated_headends(self, operator_id, operator_data_list, operator_file_path):
        tree = ET.parse(operator_file_path)
        root = tree.getroot()

        for headend in root.findall(".//headend"):
            # STEP 1: Check applicableDeviceType
            all_applicable_device_elems = headend.findall('.//applicableDeviceType')
            if not all_applicable_device_elems:
                has_stb = True
            else:
                has_stb = any(
                    'stb' in (adt.text or '').strip().lower()
                    for adt in all_applicable_device_elems
                )
            if not has_stb:
                continue

            # STEP 2: Get partnerConfigurationGroup text
            pcg = headend.find('partnerConfigurationGroup')
            pcg_text = pcg.text.strip() if pcg is not None and pcg.text and pcg.text.strip() else None

            matching_rules = []

            def is_empty_value(v):
                return pd.isna(v) or (isinstance(v, str) and v.strip() == "")

            # STEP 3: Matching logic
            if pcg_text is not None:
                for operator_data in operator_data_list:
                    old_pcg = operator_data.get("OldpartnerConfigurationGroup")
                    if not is_empty_value(old_pcg) and str(old_pcg).strip() == pcg_text:
                        matching_rules.append(operator_data)
            if not matching_rules:
                if pcg_text is None and operator_id in self.operators_with_empty_partnerconfig:
                    for operator_data in operator_data_list:
                        old_pcg = operator_data.get("OldpartnerConfigurationGroup")
                        if is_empty_value(old_pcg):
                            matching_rules.append(operator_data)

            if not matching_rules:
                continue

            # STEP 4: Apply updates
            for operator_data in matching_rules:
                delete_apps = [x.strip().lower() for x in (operator_data.get("DeleteConfiguration") or [])]
                configurations = headend.findall('configuration')
                has_deletable_app = any(
                    (assoc := cfg.find('association')) is not None and
                    (app_name := assoc.findtext('shortName')) and
                    app_name.strip().lower() in delete_apps
                    for cfg in configurations
                )
                if not has_deletable_app:
                    continue

                copied_headend = copy.deepcopy(headend)
                new_pcg = copied_headend.find('partnerConfigurationGroup')
                if new_pcg is None:
                    new_pcg = ET.Element("partnerConfigurationGroup")
                    new_pcg.text = operator_data["NewpartnerConfigurationGroup"]
                    headend_id = copied_headend.find("headendId")
                    if headend_id is not None:
                        children = list(copied_headend)
                        idx = children.index(headend_id)
                        copied_headend.insert(idx + 1, new_pcg)
                    else:
                        copied_headend.append(new_pcg)
                else:
                    new_pcg.text = operator_data["NewpartnerConfigurationGroup"]

                for configuration in list(copied_headend.findall('configuration')):
                    assoc = configuration.find('association')
                    if assoc is not None:
                        app_name = (assoc.findtext('shortName') or '').lower()
                        if app_name.strip() in delete_apps:
                            copied_headend.remove(configuration)

                self.collected_headends.setdefault(operator_id, []).append(
                    (operator_data["OperatorName"], copied_headend)
                )

    # ✅ Process and save each source file individually
    def read_xml_files(self):
        xml_files = glob.glob(os.path.join(self.source_folder, "**", "*.xml"), recursive=True)
        os.makedirs(self.directory, exist_ok=True)

        for file_path in xml_files:
            try:
                filename = os.path.basename(file_path)
                operator_id_from_file = filename.lstrip("tivo_pt.").split("_")[0]
                operator_id_from_file = "tivo:pt." + operator_id_from_file

                if operator_id_from_file not in self.operator_files_values:
                    print(f"⚠️ Skipping {filename} — no operator rules for {operator_id_from_file}")
                    continue

                operator_excel_data = self.operator_files_values[operator_id_from_file]

                self.collected_headends.clear()
                self.collect_updated_headends(operator_id_from_file, operator_excel_data, file_path)
                self.save_updated_file_for_source(file_path)

            except Exception as e:
                print(f"⚠️ Error processing file {file_path}: {e}")

    # ✅ Save with naming rules + 3 MB split logic
    def save_updated_file_for_source(self, original_file_path):
        import tempfile
        import xml.etree.ElementTree as ET

        if not self.collected_headends:
            print(f"⚠️ No headends collected for {original_file_path}")
            return

        filename = os.path.basename(original_file_path)
        name, ext = os.path.splitext(filename)

        lower_name = name.lower()

        # --- NAMING RULES ---
        if "backup" in lower_name:
            # Replace "backup" (case-insensitive) with "Upload"
            idx = lower_name.find("backup")
            name = name[:idx] + "Upload" + name[idx + len("backup"):]
        elif "upload" not in lower_name:
            # If no "backup" and no "upload", append "_Upload"
            name = name + "_Upload"

        output_name = name + ext
        base_name = os.path.splitext(output_name)[0]

        MAX_SIZE_MB = 3
        MAX_SIZE_BYTES = MAX_SIZE_MB * 1024 * 1024

        part_index = 1
        chunk_root = ET.Element("partnerServiceConfiguration")
        current_size = 0

        def save_chunk(root_elem, index):
            chunk_tree = ET.ElementTree(root_elem)
            if index > 1:
                chunk_filename = f"{base_name}_part{index}{ext}"
            else:
                chunk_filename = f"{base_name}{ext}"
            chunk_path = os.path.join(self.directory, chunk_filename)
            chunk_tree.write(chunk_path, encoding="utf-8", xml_declaration=True)
            size = os.path.getsize(chunk_path)
            print(f"📦 Saved {chunk_filename} ({size/1024/1024:.2f} MB)")
            return size

        for operator_id, headend_data in self.collected_headends.items():
            for _, headend in headend_data:
                # Measure headend size
                with tempfile.NamedTemporaryFile("wb", delete=False) as tmpf:
                    temp_root = ET.Element("partnerServiceConfiguration")
                    temp_root.append(headend)
                    ET.ElementTree(temp_root).write(tmpf, encoding="utf-8", xml_declaration=True)
                    tmpf.flush()
                    headend_size = os.path.getsize(tmpf.name)
                    os.remove(tmpf.name)

                if current_size + headend_size > MAX_SIZE_BYTES and len(chunk_root):
                    save_chunk(chunk_root, part_index)
                    part_index += 1
                    chunk_root = ET.Element("partnerServiceConfiguration")
                    current_size = 0

                chunk_root.append(headend)
                current_size += headend_size

        if len(chunk_root):
            save_chunk(chunk_root, part_index)

        print(f"✅ Completed saving {output_name}. Total parts: {part_index}")


# ✅ Helpers remain unchanged
def prettify_and_save_files(save_directory):
    import xml.dom.minidom
    import os
    import glob

    xml_files = glob.glob(os.path.join(save_directory, "*.xml"))
    if not xml_files:
        print(f"⚠️ No XML files found in {save_directory}")
        return

    for file_path in xml_files:
        try:
            if os.path.getsize(file_path) == 0:
                continue
            with open(file_path, 'r', encoding='utf-8') as file:
                xml_string = file.read().strip()
            if not xml_string:
                continue

            dom = xml.dom.minidom.parseString(xml_string)
            pretty_xml_as_string = dom.toprettyxml(indent="  ", encoding="utf-8")
            with open(file_path, 'wb') as file:
                file.write(pretty_xml_as_string)

        except Exception as e:
            print(f"⚠️ Failed to prettify {os.path.basename(file_path)}: {e}")


def count_headends_in_folder(folder_path, recursive=True):
    pattern = "**/*.xml" if recursive else "*.xml"
    xml_files = glob.glob(os.path.join(folder_path, pattern), recursive=recursive)

    for file_path in xml_files:
        try:
            tree = ET.parse(file_path)
            root = tree.getroot()
            headends = root.findall(".//headend")
            print(f"📄 {os.path.basename(file_path)}: {len(headends)} <headend> elements")
        except Exception as ex:
            print(f"⚠️ Error in {file_path}: {ex}")


# ✅ MAIN
if __name__ == "__main__":
    save_directory = "Updated_files_" + str(date.today())
    os.makedirs(save_directory, exist_ok=True)

    mediaops_file = "operators_main_file.xlsx"
    source_folder = "Bulk_operator_files"

    processor = Mediaops_headends(
        operators_main_file=mediaops_file,
        source_folder=source_folder,
        save_directory=save_directory,
    )
    processor.read_operators_rules()
    processor.read_xml_files()
    prettify_and_save_files(save_directory)
    count_headends_in_folder(save_directory)
