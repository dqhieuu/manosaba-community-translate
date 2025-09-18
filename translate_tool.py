import os
import re
import shutil
import sys
from typing import List, Tuple, Optional
from pathlib import Path

from openai import OpenAI
from openai.types.shared_params import Reasoning
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

import UnityPy

ROOT = os.path.dirname(os.path.abspath(__file__))
ORIGINAL_DIR = os.path.join(ROOT, "original")
TRANSLATED_DIR = os.path.join(ROOT, "translated")
XLSX_PATH = os.path.join(ROOT, "translate.xlsx")

HEADER = ["ID", "Original value", "Translated value", "My translated value"]
META_HEADER = ["Sheet name", "Mapped file name", "File type"]

# Knowledge base
KNOWLEDGE_SHEET = "Knowledge base"
KNOWLEDGE_HEADER = ["Knowledge"]
INITIAL_PROJECT_HEADER = [
    "Project type: Visual Novel translation.",
    "Goal: Produce high‑quality, natural translations suitable for a visual novel UI/dialogue.",
    "Guidelines: Preserve placeholders, variables, control codes, and line breaks. Maintain speaker tone, honorifics where appropriate, and context.",
    "Text between <ruby> should be converted to Romaji"
]

SHEETNAME_MAXLEN = 31


def sanitize_sheet_name(name: str) -> str:
    # Excel sheet name restrictions
    invalid = set('[]:*?/\\')
    cleaned = ''.join(c if c not in invalid else '_' for c in name)
    if len(cleaned) > SHEETNAME_MAXLEN:
        cleaned = cleaned[:SHEETNAME_MAXLEN]
    # Avoid empty or duplicate names handling will be done by caller when adding
    return cleaned or "Sheet"


def style_and_freeze(ws):
    # Header styling
    header_font = Font(bold=True)
    fill = PatternFill("solid", fgColor="DDDDDD")
    for col_idx, _ in enumerate(HEADER, start=1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = header_font
        cell.fill = fill
        cell.alignment = Alignment(vertical="top", wrap_text=True)
    ws.freeze_panes = "A2"
    # Column widths (rough defaults)
    ws.column_dimensions['A'].width = 32
    ws.column_dimensions['B'].width = 60
    ws.column_dimensions['C'].width = 60
    ws.column_dimensions['D'].width = 60


def apply_wrap_to_all_cells(ws):
    """Ensure wrap_text and top vertical alignment on all cells in the worksheet."""
    max_row = ws.max_row or 1
    max_col = ws.max_column or 1
    for row in ws.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col):
        for cell in row:
            # Preserve existing horizontal alignment if set
            horiz = getattr(cell.alignment, 'horizontal', None) if cell.alignment else None
            cell.alignment = Alignment(wrap_text=True, vertical="top", horizontal=horiz)


def is_alnum_start(s: str) -> bool:
    s = s.lstrip()
    return bool(s) and bool(re.match(r"^[A-Za-z0-9]", s))


def trim_blank_lines(text: str) -> str:
    # Normalize newlines, trim leading/trailing blank lines
    lines = text.replace('\r\n', '\n').replace('\r', '\n').split('\n')
    # Strip trailing spaces on each line but preserve internal blank lines
    lines = [ln.rstrip() for ln in lines]
    # Remove leading blank lines
    while lines and lines[0].strip() == "":
        lines.pop(0)
    # Remove trailing blank lines
    while lines and lines[-1].strip() == "":
        lines.pop()
    return "\n".join(lines)


def parse_type2(lines: List[str]) -> List[Tuple[str, str, str]]:
    """Return list of (ID, original, localized)"""
    results = []
    last_comment_block: List[str] = []

    for raw in lines:
        line = raw.rstrip('\n')
        if not line.strip():
            # blank resets nothing but separates comment blocks visually
            continue
        if line.lstrip().startswith(';'):
            # Accumulate contiguous comment lines; strip first ';' and one optional space
            comment = line.lstrip()[1:]
            if comment.startswith(' '):
                comment = comment[1:]
            last_comment_block.append(comment)
            continue
        # Non-comment: expect ID: value
        m = re.match(r"^\s*([^:]+):\s*(.*)$", line)
        if m:
            _id = m.group(1).strip()
            localized = m.group(2)
            original = "\n".join(last_comment_block).strip()
            localized = trim_blank_lines(localized)
            results.append((_id, original, localized))
            last_comment_block = []
        else:
            # Line doesn't match expected pattern; treat whole line as ID with empty value
            _id = line.strip()
            results.append((_id, "\n".join(last_comment_block).strip(), ""))
            last_comment_block = []
    return results


def parse_type1(lines: List[str]) -> List[Tuple[str, str, str]]:
    """Parse blocks of:
    # ID
    ; Original (one or more comment lines)
    Localized (one or more non-comment lines until next '#')
    Return list of (ID, original, localized).
    """
    results: List[Tuple[str, str, str]] = []
    i = 0
    n = len(lines)
    while i < n:
        s = lines[i].rstrip('\n')
        if not s.strip():
            i += 1
            continue
        if s.lstrip().startswith('#'):
            # ID line
            id_part = s.lstrip()[1:]
            if id_part.startswith(' '):
                id_part = id_part[1:]
            _id = id_part.strip()
            i += 1
            # Collect comment lines as original
            orig_lines: List[str] = []
            while i < n:
                t = lines[i].rstrip('\n')
                if t.lstrip().startswith(';'):
                    c = t.lstrip()[1:]
                    if c.startswith(' '):
                        c = c[1:]
                    orig_lines.append(c)
                    i += 1
                else:
                    break
            # Collect localized lines until next '#' or EOF
            loc_lines: List[str] = []
            while i < n:
                t = lines[i].rstrip('\n')
                if t.lstrip().startswith('#'):
                    break
                # Only treat non-comment as localized lines, but keep blank lines
                if not t.lstrip().startswith(';'):
                    loc_lines.append(t)
                i += 1
            original = trim_blank_lines("\n".join(orig_lines))
            localized = trim_blank_lines("\n".join(loc_lines))
            results.append((_id, original, localized))
            # Do not increment i here; loop will handle based on while conditions
            continue
        else:
            # Unexpected line before first '#': skip
            i += 1
            continue
    return results


def detect_file_type(lines: List[str]) -> Optional[int]:
    for raw in lines:
        s = raw.rstrip('\n')
        if not s.strip():
            continue
        if s.lstrip().startswith(';'):
            continue
        if s.lstrip().startswith('#'):
            return 1
        if is_alnum_start(s):
            return 2
        # Otherwise keep scanning
    return None


def parse_original_files() -> None:
    if os.path.exists(XLSX_PATH):
        print(f"translate.xlsx already exists at {XLSX_PATH}. Skipping parse.")
        return

    if not os.path.isdir(ORIGINAL_DIR):
        print(f"Original directory not found: {ORIGINAL_DIR}")
        sys.exit(1)

    wb = Workbook()
    # Remove the default sheet; we'll add our own
    wb.remove(wb.active)

    metadata_rows = []

    for fname in sorted(os.listdir(ORIGINAL_DIR)):
        if not fname.lower().endswith('.txt'):
            continue
        path = os.path.join(ORIGINAL_DIR, fname)
        try:
            with open(path, 'r', encoding='utf-8') as f:
                lines = f.readlines()
        except UnicodeDecodeError:
            # Fallback to cp932 (Shift-JIS) commonly used in JP assets
            with open(path, 'r', encoding='cp932', errors='replace') as f:
                lines = f.readlines()

        ftype = detect_file_type(lines)
        if ftype is None:
            print(f"Warning: Could not detect file type for {fname}. Skipping.")
            continue

        data = parse_type1(lines) if ftype == 1 else parse_type2(lines)

        # Create worksheet with sanitized, unique name
        base_sheet_name = os.path.splitext(fname)[0]
        sheet_name = sanitize_sheet_name(base_sheet_name)
        # Ensure uniqueness
        existing = set(wb.sheetnames)
        if sheet_name in existing:
            suffix = 1
            while f"{sheet_name}_{suffix}" in existing:
                suffix += 1
            sheet_name = sanitize_sheet_name(f"{sheet_name}_{suffix}")
        ws = wb.create_sheet(title=sheet_name)

        # Write header
        ws.append(HEADER)
        style_and_freeze(ws)

        for _id, original, localized in data:
            ws.append([_id, original, trim_blank_lines(localized), ""])  # My translated value blank

        # Ensure wrapping for all cells in this sheet
        apply_wrap_to_all_cells(ws)

        metadata_rows.append([sheet_name, fname, ftype])

    # Metadata sheet
    meta_ws = wb.create_sheet(title="Metadata")
    meta_ws.append(META_HEADER)
    # style
    header_font = Font(bold=True)
    fill = PatternFill("solid", fgColor="DDDDDD")
    for col_idx, _ in enumerate(META_HEADER, start=1):
        cell = meta_ws.cell(row=1, column=col_idx)
        cell.font = header_font
        cell.fill = fill
    meta_ws.freeze_panes = "A2"
    meta_ws.column_dimensions['A'].width = 32
    meta_ws.column_dimensions['B'].width = 60
    meta_ws.column_dimensions['C'].width = 12

    for row in metadata_rows:
        meta_ws.append(row)

    # Ensure wrapping for all cells in Metadata sheet
    apply_wrap_to_all_cells(meta_ws)

    # Knowledge base sheet (single column for user-provided context)
    kb_ws = wb.create_sheet(title=KNOWLEDGE_SHEET)
    kb_ws.append(KNOWLEDGE_HEADER)
    # style knowledge header
    kb_ws.cell(row=1, column=1).font = header_font
    kb_ws.cell(row=1, column=1).fill = fill
    kb_ws.freeze_panes = "A2"
    kb_ws.column_dimensions['A'].width = 100
    # Place initial project header as the first lines to guide the translator/AI
    for line in INITIAL_PROJECT_HEADER:
        kb_ws.append([line])
    apply_wrap_to_all_cells(kb_ws)

    wb.save(XLSX_PATH)
    print(f"translate.xlsx created at {XLSX_PATH}")


def get_knowledge_text(wb) -> str:
    if KNOWLEDGE_SHEET in wb.sheetnames:
        ws = wb[KNOWLEDGE_SHEET]
        # Concatenate non-empty cells in first column excluding header
        parts = []
        for r in ws.iter_rows(min_row=2, max_col=1, values_only=True):
            if not r:
                continue
            v = (r[0] or "").strip()
            if v:
                parts.append(v)
        return "\n\n".join(parts).strip()
    return "\n\n".join(INITIAL_PROJECT_HEADER)


def translate_ai(num_lines: int) -> None:
    if not os.path.exists(XLSX_PATH):
        print(f"translate.xlsx not found at {XLSX_PATH}. Run parse first.")
        sys.exit(1)

    api_key = os.environ.get("OPENAI_API_KEY")
    if not api_key:
        print("Environment variable OPENAI_API_KEY is not set.")
        sys.exit(1)

    client = OpenAI(api_key=api_key)
    model = os.environ.get("OPENAI_MODEL", "gpt-5-mini")

    wb = load_workbook(XLSX_PATH)
    knowledge_text = get_knowledge_text(wb)

    processed = 0

    def make_messages(sheet_name: str, row_id: str, original: str, translated: str):
        sys_prompt = "\n\nKnowledge base (user-provided notes):\n" + (knowledge_text or "<empty>")
        user_prompt = (
            f"Sheet: {sheet_name}\n"
            f"ID: {row_id}\n\n"
            f"Original value (source 1):\n{original or '<empty>'}\n\n"
            f"Translated value (source 2):\n{translated or '<empty>'}\n\n"
            "Task: Produce the final localized line(s) for 'My translated value' in Tiếng Việt (Vietnamese). Should use both of Original value and Translated value for the best result\n"
            "Rules: Keep placeholders, tags, and control codes intact. Translate the text inside XML tag pairs, not the tag itself. Preserve line breaks. Return ONLY the final text, no explanations."
        )
        return (sys_prompt, user_prompt)

    # Iterate sheets in workbook order, skipping special sheets
    special = {"Metadata", KNOWLEDGE_SHEET}
    for sheet_name in wb.sheetnames:
        if sheet_name in special:
            continue
        ws = wb[sheet_name]
        # Identify columns by header row 1
        headers = [ (ws.cell(row=1, column=c).value or "").strip() for c in range(1, 10) ]
        try:
            col_id = headers.index("ID") + 1
            col_orig = headers.index("Original value") + 1
            col_trans = headers.index("Translated value") + 1
            col_my = headers.index("My translated value") + 1
        except ValueError:
            # Unexpected sheet format
            continue
        # Iterate rows starting at 2
        for r in range(2, ws.max_row + 1):
            if processed >= num_lines:
                break
            row_id = (ws.cell(row=r, column=col_id).value or "").strip()
            if not row_id:
                continue
            my_val = (ws.cell(row=r, column=col_my).value or "").strip()
            if my_val != "":
                continue
            original = ws.cell(row=r, column=col_orig).value or ""
            translated = ws.cell(row=r, column=col_trans).value or ""

            # Skip if there is nothing to translate
            if not original and not translated:
                continue

            (sys_prompt, user_prompt) = make_messages(sheet_name, row_id, original, translated)
            try:
                resp = client.responses.create(
                    model=model,
                    reasoning=Reasoning(effort="minimal"),
                    instructions=sys_prompt,
                    input=user_prompt
                )
                ai_text = (resp.output_text or "").strip()
            except Exception as e:
                print(f"OpenAI API error on {sheet_name} row {r}: {e}")
                ai_text = ""

            if ai_text:
                ws.cell(row=r, column=col_my).value = ai_text
                processed += 1
                print(f"Translated: {sheet_name} | ID {row_id}. Result: {ai_text}")

        if processed >= num_lines:
            break

    if processed > 0:
        wb.save(XLSX_PATH)
        print(f"Saved {processed} AI translations to {XLSX_PATH}")
    else:
        print("No rows required translation or already filled.")


def rebuild_translated_files() -> None:
    if not os.path.exists(XLSX_PATH):
        print(f"translate.xlsx not found at {XLSX_PATH}. Run parse first.")
        sys.exit(1)

    wb = load_workbook(XLSX_PATH)
    if "Metadata" not in wb.sheetnames:
        print("Metadata sheet not found in translate.xlsx")
        sys.exit(1)

    meta_ws = wb["Metadata"]
    # Read metadata rows, skip header
    mappings: List[Tuple[str, str, int]] = []
    for r in meta_ws.iter_rows(min_row=2, values_only=True):
        if not r or all(v is None for v in r):
            continue
        sheet_name, mapped_file, ftype = r[:3]
        if not sheet_name or not mapped_file or not ftype:
            continue
        mappings.append((str(sheet_name), str(mapped_file), int(ftype)))

    if os.path.exists(TRANSLATED_DIR):
        shutil.rmtree(TRANSLATED_DIR)
    os.makedirs(TRANSLATED_DIR, exist_ok=True)

    for sheet_name, mapped_file, ftype in mappings:
        if sheet_name not in wb.sheetnames:
            print(f"Warning: Sheet '{sheet_name}' not found. Skipping {mapped_file}.")
            continue
        ws = wb[sheet_name]
        # Build a dictionary ID -> (original, translated, my_translated)
        id_rows = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row:
                continue
            _id = (row[0] or "").strip()
            if not _id:
                continue
            original = (row[1] or "")
            translated = trim_blank_lines(row[2] or "")
            my_translated = trim_blank_lines(row[3] or "")
            id_rows.append((_id, original, translated, my_translated))

        out_path = os.path.join(TRANSLATED_DIR, mapped_file)
        os.makedirs(os.path.dirname(out_path), exist_ok=True)

        lines_out: List[str] = []

        def add_comment_block(original_text: str, localized_text: str):
            # Original comment
            for ln in (original_text or "").replace('\r\n', '\n').replace('\r', '\n').split('\n'):
                lines_out.append("; " + ln if ln.strip() != "" else ";")
            # Localized header
            lines_out.append("; **Localized**")
            # Localized comment
            for ln in (localized_text or "").replace('\r\n', '\n').replace('\r', '\n').split('\n'):
                lines_out.append("; " + ln if ln.strip() != "" else ";")

        if ftype == 2:
            # Each row becomes: comments + "ID: value"
            for _id, original, translated, my_translated in id_rows:
                used_value = my_translated.strip() if my_translated.strip() != "" else translated
                used_value = trim_blank_lines(used_value)
                add_comment_block(original, translated)
                # Add the actual line
                lines_out.append(f"{_id}: {used_value}")
                # Blank line between entries for readability
                lines_out.append("")
        elif ftype == 1:
            # Blocks: "# ID" then comments then localized lines, separated until next block
            for _id, original, translated, my_translated in id_rows:
                used_value = my_translated.strip() if my_translated.strip() != "" else translated
                used_value = trim_blank_lines(used_value)
                lines_out.append(f"# {_id}")
                add_comment_block(original, translated)
                # Localized content lines
                if used_value == "":
                    # Keep at least a blank line to separate blocks
                    lines_out.append("")
                else:
                    for ln in used_value.split('\n'):
                        lines_out.append(ln)
                # Extra blank line between blocks
                lines_out.append("")
        else:
            print(f"Warning: Unknown file type {ftype} for {mapped_file}. Skipping.")
            continue

        # Write out with UTF-8 BOM to be safe for some VN engines
        with open(out_path, 'w', encoding='utf-8-sig', newline='') as f:
            f.write("\n".join(lines_out).rstrip() + "\n")
        print(f"Wrote {out_path}")

def pack_translated_files(folder_path: str) -> None:
    folder = Path(folder_path)
    bundle_paths = list(folder.rglob("*.bundle"))

    if not bundle_paths:
        print(f"No .bundle files found in {folder_path}")
        return

    # Prepare backup folder (only back up modified bundles)
    backup_folder = Path(folder_path + "_backup")
    backup_folder.mkdir(parents=True, exist_ok=True)
    print(f"Using backup folder: {backup_folder}")

    print(f"Found {len(bundle_paths)} .bundle files:")

    translated_file_dict = {}
    for file in Path(TRANSLATED_DIR).rglob("*.txt"):
        translated_file_dict[file.name] = file

    for bundle_path in bundle_paths:
        try:
            bundle_path_str = str(bundle_path)
            bundle = UnityPy.load(bundle_path_str)
            bundle_modified = False
            for obj in bundle.objects:
                if obj.type.name == "TextAsset":
                    name = obj.container.split('/')[-1]
                    if name not in translated_file_dict: continue

                    translated_file = translated_file_dict[name]
                    with open(translated_file, 'r', encoding='utf-8') as f:
                        translated_text = f.read()

                    data = obj.read()
                    data.m_Script = translated_text
                    data.save()
                    print(f"Replaced {name} in {bundle_path_str}")
                    bundle_modified = True
            if bundle_modified:
                # Back up the original bundle into backup folder, preserving relative path
                try:
                    rel_path = bundle_path.relative_to(folder)
                except ValueError:
                    rel_path = Path(bundle_path.name)
                backup_path = backup_folder / rel_path
                backup_path.parent.mkdir(parents=True, exist_ok=True)
                if not backup_path.exists():
                    shutil.copy2(bundle_path, backup_path)
                    print(f"Backed up original: {backup_path}")

                # Save modified bundle back to its original location
                bundle.save(pack="lz4", out_path=os.path.dirname(bundle_path))
                print(f"Saved {bundle_path_str}")


        except Exception as e:
            print(f"Error processing {bundle_path}: {e}")



def main():
    if len(sys.argv) < 2:
        print("Usage: python translate_tool.py [parse|build|pack <folder>|build+pack <folder>|translate <num>]")
        sys.exit(1)
    cmd = sys.argv[1].lower()
    if cmd == 'parse':
        parse_original_files()
    elif cmd == 'build':
        rebuild_translated_files()
    elif cmd == 'pack' and len(sys.argv) >= 3:
        pack_translated_files(sys.argv[2])
    elif cmd == 'build+pack' and len(sys.argv) >= 3:
        rebuild_translated_files()
        pack_translated_files(sys.argv[2])
    elif cmd == 'translate' and len(sys.argv) >= 3:
        try:
            n = int(sys.argv[2])
        except ValueError:
            print("For 'translate', provide a number of lines to process. Example: translate 10")
            sys.exit(1)
        if n <= 0:
            print("Number of lines must be positive.")
            sys.exit(1)
        translate_ai(n)
    else:
        print("Unknown command. Use 'parse' or 'build' or 'pack <folder>' or 'build+pack <folder>' or 'translate <num>'.")
        sys.exit(1)


if __name__ == '__main__':
    main()
