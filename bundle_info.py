import os
from pathlib import Path
import UnityPy
import yaml
import json
import re
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import argparse

# Configuration
OUTPUT_XLSX = "bundle_info.xlsx"
ADDRESSES_PATH = os.path.join("patches", "addresses.txt")
IGNORED_BUNDLE_SUFFIXES = ['general-managedtext_assets_all.bundle']

# Excel sheet constants
SHEET_NAME = "Bundle Info"
HEADER = ["Bundle Path Suffix", "Container", "Name", "Type", "PathID", "Object Selector", "Original", "Chinese", "Translated"]

def apply_header_and_column_widths(ws, headers, freeze_panes_cell="A2"):
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

    # Set column widths
    column_widths = [50, 60, 40, 20, 16, 60, 60, 60, 60]
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

def extract_localized_texts(tree, base_path=""):
    """Recursively extract localized texts from the tree structure, specifically handling _items and _taggedText."""
    texts = []
    if isinstance(tree, dict):
        for key, value in tree.items():
            new_path = f"{base_path}.{key}" if base_path else key
            if key == "_items" and isinstance(value, list):
                for idx, item in enumerate(value):
                    item_path = f"{new_path}[{idx}]"
                    if isinstance(item, dict) and "_taggedText" in item:
                        tagged_text = item["_taggedText"]
                        if isinstance(tagged_text, list):
                            for t_idx, text_entry in enumerate(tagged_text):
                                if isinstance(text_entry, dict) and "_locale" in text_entry and "_text" in text_entry:
                                    locales = {text_entry["_locale"]: text_entry["_text"]}
                                    original = locales.get(0, "")
                                    chinese = locales.get(2, "")  # Assuming 2 is Simplified Chinese
                                    selector = f"{item_path}._taggedText[{t_idx}]._text"
                                    texts.append((selector, original, chinese))
                    texts.extend(extract_localized_texts(item, item_path))
            else:
                texts.extend(extract_localized_texts(value, new_path))
    elif isinstance(tree, list):
        for idx, item in enumerate(tree):
            new_path = f"{base_path}[{idx}]"
            texts.extend(extract_localized_texts(item, new_path))
    return texts

def get_extracted_texts(obj):
    """Extract object selectors, original (locale 0), and Chinese (locale 2) texts."""
    if obj.type.name == "TextAsset":
        data = obj.read()
        try:
            text = data.script.decode('utf-8')
        except:
            text = ""
        try:
            tree = json.loads(text)
            return extract_localized_texts(tree)
        except:
            return [("", text, "")]
    elif obj.type.name == "MonoBehaviour":
        try:
            tree = obj.read_typetree()
            extracted = extract_localized_texts(tree)
            if extracted:
                return extracted
            if 'm_Text' in tree:
                return [("m_Text", tree['m_Text'], "")]
            return [("", "", "")]
        except:
            return [("", "", "")]
    return [("", "", "")]

def generate_bundle_info(folder_path: str):
    """Generate an Excel file with bundle asset information, listing bundle path suffix, container, name, type, and pathID once."""
    from itertools import groupby
    from operator import itemgetter

    folder = Path(folder_path)
    bundle_paths = [p for p in folder.rglob("*.bundle") if
                    not any(p.name.endswith(suf) for suf in IGNORED_BUNDLE_SUFFIXES)]

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
    apply_header_and_column_widths(ws, HEADER)

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
                container = next((x[0] for x in ab['m_Container'] if x[1]['asset']['m_PathID'] == obj.path_id), None) if 'ab' in locals() else ""

                extracted = get_extracted_texts(obj)
                for object_selector, original, chinese in extracted:
                    bundle_data[bundle_suffix].append({
                        "container": container,
                        "name": name,
                        "type": obj.type.name,
                        "path_id": str(obj.path_id),
                        "object_selector": object_selector,
                        "original": original,
                        "chinese": chinese,
                        "translated": "",
                        "patch_entries": []
                    })

                # Check for patches for MonoBehaviour
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
                                    "object_selector": selector,
                                    "original": entry.get('patched_value', ''),
                                    "chinese": "",
                                    "translated": "",
                                    "patch_entries": []
                                }
                                bundle_data[bundle_suffix].append(patched_entry)

            print(f"Processed {bundle_path}")

        except Exception as e:
            print(f"Error processing {bundle_path}: {e}")

    # Write to Excel, listing bundle path suffix and container, name, type, pathID once
    for bundle_suffix in sorted(bundle_data.keys()):
        assets = bundle_data[bundle_suffix]
        if not assets:
            continue

        # Write bundle path suffix once
        ws.append([bundle_suffix, "", "", "", "", "", "", "", ""])

        # Sort assets by container, name, type, path_id for consistent grouping
        sorted_assets = sorted(assets, key=lambda x: (x["container"] or "", x["name"], x["type"], x["path_id"]))

        # Group by container, name, type, and path_id
        for key, group in groupby(sorted_assets, key=lambda x: (x["container"], x["name"], x["type"], x["path_id"])):
            group_list = list(group)
            for idx, asset in enumerate(group_list):
                # Write container, name, type, and path_id only for the first asset in the group
                container_value = key[0] if idx == 0 else ""
                name_value = key[1] if idx == 0 else ""
                type_value = key[2] if idx == 0 else ""
                path_id_value = key[3] if idx == 0 else ""
                ws.append([
                    "",  # Empty bundle suffix for asset rows
                    container_value,
                    name_value,
                    type_value,
                    path_id_value,
                    asset["object_selector"],
                    asset["original"],
                    asset["chinese"],
                    asset["translated"]
                ])

    apply_wrap_to_all_cells(ws)
    wb.save(OUTPUT_XLSX)
    print(f"Saved bundle information to {OUTPUT_XLSX}")

def _set_by_selector(tree, selector: str, value: str) -> bool:
    """Set value in tree by selector path."""
    # Split on . not inside brackets
    keys = re.split(r'\.(?![^\[\]]*])', selector)
    current = tree
    for key in keys[:-1]:
        if '[' in key:
            base, idx_str = key.split('[', 1)
            idx_str = idx_str.rstrip(']')
            if base in current:
                arr = current[base]
                if isinstance(arr, list) and len(arr) > int(idx_str):
                    current = arr[int(idx_str)]
                else:
                    return False
            else:
                return False
        else:
            if key in current:
                current = current[key]
            else:
                return False
    last_key = keys[-1]
    if '[' in last_key:
        base, idx_str = last_key.split('[', 1)
        idx_str = idx_str.rstrip(']')
        if base in current:
            arr = current[base]
            if isinstance(arr, list) and len(arr) > int(idx_str):
                if isinstance(arr[int(idx_str)], dict) and "_text" in arr[int(idx_str)]:
                    arr[int(idx_str)]["_text"] = value
                    return True
                return False
            else:
                return False
        else:
            return False
    else:
        if last_key in current:
            current[last_key] = value
            return True
        return False

def build_patches_from_excel(excel_path: str):
    """Generate addresses.txt from the Chinese column in the Excel file."""
    if not os.path.exists(excel_path):
        print(f"Error: Excel file {excel_path} not found")
        return

    try:
        wb = load_workbook(excel_path)
        ws = wb[SHEET_NAME]
    except Exception as e:
        print(f"Error: Failed to load Excel file {excel_path}: {e}")
        return

    patches = {}
    current_bundle = None

    for row in ws.iter_rows(min_row=2, values_only=True):
        bundle_suffix, container, name, type_, path_id, object_selector, _, chinese, _ = row

        # Update current bundle if bundle_suffix is non-empty
        if bundle_suffix:
            current_bundle = bundle_suffix

        # Skip rows without a Chinese text or without a bundle
        if not chinese or not current_bundle:
            continue

        # Initialize patch dictionary for the bundle if not exists
        if current_bundle not in patches:
            patches[current_bundle] = {}

        # Initialize patch entry for the path_id if not exists
        if path_id and path_id not in patches[current_bundle]:
            patches[current_bundle][path_id] = []

        # Add patch entry if there is an object selector or for TextAsset
        if path_id and (object_selector or type_ == "TextAsset"):
            patches[current_bundle][path_id].append({
                "object_selector": object_selector,
                "patched_value": chinese
            })

    # Save patches to addresses.txt
    os.makedirs(os.path.dirname(ADDRESSES_PATH), exist_ok=True)
    try:
        with open(ADDRESSES_PATH, 'w', encoding='utf-8') as f:
            yaml.safe_dump(patches, f, allow_unicode=True, sort_keys=True)
        print(f"Saved patches to {ADDRESSES_PATH}")
    except Exception as e:
        print(f"Error: Failed to save patches to {ADDRESSES_PATH}: {e}")

def apply_patches_to_bundle(bundle_path: str, patches: dict):
    """Apply patches from addresses.txt to the specified bundle file using Chinese text."""
    bundle_suffix = str(Path(bundle_path).relative_to(Path(bundle_path).parent.parent))
    if bundle_suffix not in patches:
        print(f"No patches found for {bundle_suffix}")
        return

    try:
        env = UnityPy.load(str(bundle_path))
        modified = False

        for obj in env.objects:
            if obj.type.name not in ["TextAsset", "MonoBehaviour"]:
                continue

            path_id = str(obj.path_id)
            if path_id not in patches[bundle_suffix]:
                continue

            todo_entries = patches[bundle_suffix][path_id]

            if obj.type.name == "TextAsset":
                data = obj.read()
                # Assume it's JSON if selector present
                has_selector = any(ent.get('object_selector') for ent in todo_entries)
                if has_selector:
                    try:
                        json_str = data.script.decode('utf-8')
                        tree = json.loads(json_str)
                        any_patched = False
                        for ent in todo_entries:
                            selector = ent.get('object_selector')
                            value = ent.get('patched_value')
                            if selector is None:
                                continue
                            ok = _set_by_selector(tree, selector, value)
                            if ok:
                                any_patched = True
                        if any_patched:
                            new_json = json.dumps(tree, ensure_ascii=False, indent=None)
                            data.script = new_json.encode('utf-8')
                            modified = True
                    except json.JSONDecodeError as e:
                        print(f"Failed to parse JSON for TextAsset in {bundle_suffix}: {e}")
                else:
                    # No selector: replace entire content
                    for ent in todo_entries:
                        if ent.get('object_selector') is None:
                            data.script = ent.get('patched_value', '').encode('utf-8')
                            modified = True
                            break

            elif obj.type.name == "MonoBehaviour":
                try:
                    tree = obj.read_typetree()
                except Exception as e:
                    print(f"Failed to parse MonoBehaviour in {bundle_suffix}: {e}")
                    continue
                any_patched = False
                for ent in todo_entries:
                    selector = ent.get('object_selector')
                    value = ent.get('patched_value')
                    if selector is None:
                        continue
                    ok = _set_by_selector(tree, selector, value)
                    if ok:
                        any_patched = True
                if any_patched:
                    obj.save_typetree(tree)
                    modified = True

        if modified:
            output_path = str(Path(bundle_path).parent / f"{Path(bundle_path).stem}_patched.bundle")
            with open(output_path, 'wb') as f:
                f.write(env.save(pack='lz4'))
            print(f"Saved patched bundle to {output_path}")

    except Exception as e:
        print(f"Error processing bundle {bundle_suffix}: {e}")

def build_and_patch(folder_path: str):
    """Generate patches from Excel and apply them to bundle files."""
    # First, build the patches from the Excel file
    build_patches_from_excel(OUTPUT_XLSX)

    # Load the generated patches
    patches = load_patches_from_file()
    if not patches:
        print("No patches to apply")
        return

    # Apply patches to all bundle files
    folder = Path(folder_path)
    bundle_paths = [p for p in folder.rglob("*.bundle") if
                    not any(p.name.endswith(suf) for suf in IGNORED_BUNDLE_SUFFIXES)]

    for bundle_path in bundle_paths:
        apply_patches_to_bundle(str(bundle_path), patches)

def main():
    parser = argparse.ArgumentParser(description="Process Unity bundle files for translation.")
    parser.add_argument("command", choices=["parse", "build", "build+patch"], help="Operation to perform: parse, build, or build+patch")
    parser.add_argument("folder_path", help="Path to the folder containing .bundle files")
    args = parser.parse_args()

    folder_path = args.folder_path
    if not os.path.isdir(folder_path):
        print(f"Error: {folder_path} is not a valid directory")
        return

    if args.command == "parse":
        generate_bundle_info(folder_path)
    elif args.command == "build":
        build_patches_from_excel(OUTPUT_XLSX)
    elif args.command == "build+patch":
        build_and_patch(folder_path)

if __name__ == "__main__":
    main()