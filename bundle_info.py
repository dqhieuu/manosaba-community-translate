import os
import sys
from pathlib import Path
import UnityPy
import yaml
import json
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# Configuration
OUTPUT_XLSX = "bundle_info.xlsx"
ADDRESSES_PATH = os.path.join("patches", "addresses.txt")
IGNORED_BUNDLE_SUFFIXES = ['general-managedtext_assets_all.bundle']
IGNORED_CONTAINERS = ['Assets/#WitchTrials/Data/ScriptableObjects/SpecialThanksData.asset']

# Excel sheet constants
SHEET_NAME = "Bundle Info"
HEADER = ["Bundle Path Suffix", "Container", "Name", "Type", "PathID", "Original Object Selector", "Original",
          "Chinese Object Selector", "Chinese", "Translated"]

PATCH_SHEET_NAME = "Patch Addresses"
PATCH_HEADER = ["Bundle path suffix", "PathID", "Object selector", "Original", "Translated"]


def apply_header_and_column_widths(ws, headers, column_widths=None, freeze_panes_cell="A2"):
    """Apply bold + gray header styling and set column widths."""
    header_font = Font(bold=True)
    fill = PatternFill("solid", fgColor="DDDDDD")
    for col_idx, _ in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = header_font
        cell.fill = fill
        cell.alignment = Alignment(vertical="top", wrap_text=True)
    if freeze_panes_cell:
        ws.freeze_panes = freeze_panes_cell

    if column_widths:
        for idx, width in enumerate(column_widths, start=1):
            col_letter = get_column_letter(idx)
            ws.column_dimensions[col_letter].width = width


def apply_wrap_to_all_cells(ws):
    """Ensure wrap_text and top vertical alignment on all cells."""
    max_row = ws.max_row or 1
    max_col = ws.max_column or 1
    for row in ws.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col):
        for cell in row:
            horiz = getattr(cell.alignment, 'horizontal', None) if cell.alignment else None
            cell.alignment = Alignment(wrap_text=True, vertical="top", horizontal=horiz)


def load_patches_from_file() -> dict:
    """Load patch data from addresses.txt if it exists."""
    if not os.path.exists(ADDRESSES_PATH):
        return {}
    try:
        with open(ADDRESSES_PATH, 'r', encoding='utf-8') as f:
            content = f.read().strip()
            if not content:
                return {}
            data = yaml.safe_load(content)
            return data or {}
    except Exception as e:
        print(f"Warning: Failed to read {ADDRESSES_PATH}: {e}")
        return {}


def extract_localized_texts(tree, base_path="", bundle_suffix=""):
    """Recursively extract localized texts, pairing locale 0 and 2 for the same item index."""
    texts = []
    if isinstance(tree, dict):
        if "_locale" in tree and "_text" in tree:
            selector = f"{base_path}._text" if base_path else "_text"
            locale = tree["_locale"]
            text = tree["_text"]
            return [(selector, text, locale)]

        for key, value in tree.items():
            new_path = f"{base_path}.{key}" if base_path else key
            if isinstance(value, list) and value:
                if isinstance(value[0], dict) and "_locale" in value[0] and "_text" in value[0]:
                    locale_texts = {}
                    for idx, item in enumerate(value):
                        selector = f"{new_path}[{idx}]._text"
                        locale = item["_locale"]
                        text = item["_text"]
                        locale_texts[locale] = (selector, text)
                    orig_selector = locale_texts.get(0, ("", ""))[0]
                    orig_text = locale_texts.get(0, ("", ""))[1]
                    cn_selector = locale_texts.get(2, ("", ""))[0]
                    cn_text = locale_texts.get(2, ("", ""))[1]
                    if orig_text or cn_text:
                        texts.append((orig_selector, orig_text, cn_selector, cn_text))
                else:
                    for idx, item in enumerate(value):
                        item_path = f"{new_path}[{idx}]"
                        texts.extend(extract_localized_texts(item, item_path, bundle_suffix))
            else:
                texts.extend(extract_localized_texts(value, new_path, bundle_suffix))
    elif isinstance(tree, list):
        for idx, item in enumerate(tree):
            new_path = f"{base_path}[{idx}]"
            texts.extend(extract_localized_texts(item, new_path, bundle_suffix))
    return texts


def get_extracted_texts(obj, bundle_suffix=""):
    """Extract object selectors and texts for original (locale 0) and Chinese (locale 2)."""
    if obj.type.name == "TextAsset":
        data = obj.read()
        try:
            text = data.script.decode('utf-8')
        except:
            text = ""
        try:
            tree = json.loads(text)
            return extract_localized_texts(tree, bundle_suffix=bundle_suffix)
        except:
            return [("", text, "", "")]
    elif obj.type.name == "MonoBehaviour":
        try:
            tree = obj.read_typetree()
            extracted = extract_localized_texts(tree, bundle_suffix=bundle_suffix)
            if extracted:
                return extracted
            if 'm_Text' in tree:
                return [("m_Text", tree['m_Text'], "", "")]
            return [("", "", "", "")]
        except:
            return [("", "", "", "")]
    return [("", "", "", "")]


def generate_bundle_info(folder_path: str):
    """Generate an Excel file with bundle asset information, grouping by container."""
    folder = Path(folder_path)
    bundle_paths = [p for p in folder.rglob("*.bundle") if
                    not any(str(p).endswith(suf) for suf in IGNORED_BUNDLE_SUFFIXES)]

    if not bundle_paths:
        print(f"No .bundle files found in {folder_path}")
        return

    # Load patches
    patches = load_patches_from_file()

    # Create Excel workbook
    wb = Workbook()
    ws = wb.active
    ws.title = SHEET_NAME
    ws.append(HEADER)
    apply_header_and_column_widths(ws, HEADER, [50, 60, 40, 20, 16, 60, 60, 60, 60, 60])

    # Collect all asset data grouped by bundle
    bundle_data = {}
    for bundle_path in bundle_paths:
        bundle_suffix = str(bundle_path.relative_to(folder))
        bundle_data[bundle_suffix] = []

        try:
            bundle = UnityPy.load(str(bundle_path))

            for obj in bundle.objects:
                if obj.type.name not in ["AssetBundle"]:
                    continue
                ab = obj.read_typetree()
                break

            for obj in bundle.objects:
                if obj.type.name not in ["TextAsset", "MonoBehaviour", "Texture2D"]:
                    continue

                name = obj.name or f"Unnamed_{obj.type.name}_{obj.path_id}"
                container = next((x[0] for x in ab['m_Container'] if x[1]['asset']['m_PathID'] == obj.path_id),
                                 None) if 'ab' in locals() else ""

                if container in IGNORED_CONTAINERS:
                    continue

                extracted = get_extracted_texts(obj, bundle_suffix)
                for orig_selector, original, cn_selector, chinese in extracted:
                    bundle_data[bundle_suffix].append({
                        "container": container,
                        "name": name,
                        "type": obj.type.name,
                        "path_id": str(obj.path_id),
                        "original_selector": orig_selector,
                        "original": original,
                        "chinese_selector": cn_selector,
                        "chinese": chinese,
                        "translated": "",
                        "patch_entries": []
                    })

                if obj.type.name == "MonoBehaviour" and bundle_suffix in patches:
                    id_map = patches.get(bundle_suffix, {})
                    pid_str = str(obj.path_id)
                    if pid_str in id_map:
                        for entry in id_map[pid_str]:
                            selector = entry.get('object_selector', '')
                            if selector:
                                patched_entry = {
                                    "container": container,
                                    "name": name,
                                    "type": obj.type.name,
                                    "path_id": pid_str,
                                    "original_selector": selector,
                                    "original": entry.get('patched_value', ''),
                                    "chinese_selector": selector,
                                    "chinese": entry.get('patched_value', ''),
                                    "translated": "",
                                    "patch_entries": []
                                }
                                if container not in IGNORED_CONTAINERS:
                                    bundle_data[bundle_suffix].append(patched_entry)

            print(f"Processed {bundle_path}")

        except Exception as e:
            print(f"Error processing {bundle_path}: {e}")

    # Write to Excel, including all fields for every row
    all_rows_for_patch = []
    for bundle_suffix in sorted(bundle_data.keys()):
        assets = bundle_data[bundle_suffix]
        if not assets:
            continue

        sorted_assets = sorted(assets, key=lambda x: (x["container"] or "", x["name"], x["type"], x["path_id"]))

        for asset in sorted_assets:
            ws.append([
                bundle_suffix,
                asset["container"],
                asset["name"],
                asset["type"],
                asset["path_id"],
                asset["original_selector"],
                asset["original"],
                asset["chinese_selector"],
                asset["chinese"],
                asset["translated"]
            ])

            all_rows_for_patch.append([
                bundle_suffix,
                asset["path_id"],
                asset["chinese_selector"],
                asset["chinese"],
                asset["translated"]
            ])

    apply_wrap_to_all_cells(ws)

    # Tạo sheet Patch Addresses
    ws_patch = wb.create_sheet(PATCH_SHEET_NAME)
    ws_patch.append(PATCH_HEADER)
    apply_header_and_column_widths(ws_patch, PATCH_HEADER, [50, 16, 60, 60, 60])
    for row in all_rows_for_patch:
        ws_patch.append(row)
    apply_wrap_to_all_cells(ws_patch)

    wb.save(OUTPUT_XLSX)
    print(f"Saved bundle information to {OUTPUT_XLSX}")


# ====================== PACK FUNCTION IMPORT ======================
from translate_tool import (
    rebuild_translated_files,
    write_patches_from_sheet,
    pack_translated_files
)

def main():
    import sys  # thêm dòng này

    command_usage = "python bundle_info.py [info <folder>|pack <folder>]"
    if len(sys.argv) < 3:
        print(f"Usage: {command_usage}")
        sys.exit(1)

    cmd = sys.argv[1].lower()
    folder_path = sys.argv[2]

    if not os.path.isdir(folder_path):
        print(f"Error: {folder_path} is not a valid directory")
        sys.exit(1)

    if cmd == 'info':
        generate_bundle_info(folder_path)
    elif cmd == 'pack':
        rebuild_translated_files()
        write_patches_from_sheet()
        pack_translated_files(folder_path)
    else:
        print(f"Unknown command. Use {command_usage}.")
        sys.exit(1)

if __name__ == "__main__":
    main()
