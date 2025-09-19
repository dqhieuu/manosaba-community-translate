import os
import re
import shutil
import sys
from typing import List, Tuple, Optional
from pathlib import Path
from collections import defaultdict

from openai import OpenAI
from openai.types.shared_params import Reasoning
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

import UnityPy

ROOT = os.path.dirname(os.path.abspath(__file__))
ORIGINAL_DIR = os.path.join(ROOT, "original")
TRANSLATED_DIR = os.path.join(ROOT, "translated")
XLSX_PATH = os.path.join(ROOT, "translate.xlsx")

HEADER = ["ID", "Original", "Chinese", "MTL", "Edited"]
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
    ws.column_dimensions['E'].width = 60

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
            continue
        if line.lstrip().startswith(';'):
            comment = line.lstrip()[1:]
            if comment.startswith(' '):
                comment = comment[1:]
            last_comment_block.append(comment)
            continue
        m = re.match(r"^\s*([^:]+):\s*(.*)$", line)
        if m:
            _id = m.group(1).strip()
            localized = m.group(2)
            original = "\n".join(last_comment_block).strip()
            localized = trim_blank_lines(localized)
            results.append((_id, original, localized))
            last_comment_block = []
        else:
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
            id_part = s.lstrip()[1:]
            if id_part.startswith(' '):
                id_part = id_part[1:]
            _id = id_part.strip()
            i += 1
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
            loc_lines: List[str] = []
            while i < n:
                t = lines[i].rstrip('\n')
                if t.lstrip().startswith('#'):
                    break
                if not t.lstrip().startswith(';'):
                    loc_lines.append(t)
                i += 1
            original = trim_blank_lines("\n".join(orig_lines))
            localized = trim_blank_lines("\n".join(loc_lines))
            results.append((_id, original, localized))
            continue
        else:
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
    return None

def get_content_sheets(wb):
    special_sheets = {"Overview", "Metadata", KNOWLEDGE_SHEET, "Summaries"}
    return [s for s in wb.sheetnames if s not in special_sheets]

def update_overview(wb):
    if "Overview" not in wb.sheetnames:
        ov_ws = wb.create_sheet(title="Overview", index=0)
        ov_ws.append(["Act", "Chapter", "File", "Total Lines", "MTL %", "Edited %"])
        header_font = Font(bold=True)
        fill = PatternFill("solid", fgColor="DDDDDD")
        for col_idx in range(1, 7):
            cell = ov_ws.cell(row=1, column=col_idx)
            cell.font = header_font
            cell.fill = fill
        ov_ws.freeze_panes = "A2"
        ov_ws.column_dimensions['A'].width = 20
        ov_ws.column_dimensions['B'].width = 20
        ov_ws.column_dimensions['C'].width = 40
        ov_ws.column_dimensions['D'].width = 20
        ov_ws.column_dimensions['E'].width = 20
        ov_ws.column_dimensions['F'].width = 20
        apply_wrap_to_all_cells(ov_ws)
    else:
        ov_ws = wb["Overview"]
        for row in range(ov_ws.max_row, 1, -1):
            ov_ws.delete_rows(row)

    content_sheets = get_content_sheets(wb)
    structure = defaultdict(lambda: defaultdict(list))
    for sheet_name in content_sheets:
        if sheet_name.lower().startswith('common'):
            structure['Common']['Common'].append(sheet_name)
        else:
            match = re.match(r"^(Act\d+)_Chapter(\d+)_(.+)$", sheet_name)
            if not match:
                continue
            act, chapter, file_type = match.groups()
            structure[act][chapter].append(sheet_name)

    total_all_lines = 0
    total_all_mtl = 0
    total_all_edited = 0

    last_act = None
    last_chapter = None
    for act in sorted(structure.keys()):
        act_total_lines = 0
        act_mtl_completed = 0
        act_edited_completed = 0
        for chapter in sorted(structure[act].keys()):
            chapter_total_lines = 0
            chapter_mtl_completed = 0
            chapter_edited_completed = 0
            adv_total = trial_total = bad_total = common_total = 0
            adv_mtl = trial_mtl = bad_mtl = common_mtl = 0
            adv_edited = trial_edited = bad_edited = common_edited = 0

            for sheet_name in sorted(structure[act][chapter]):
                ws = wb[sheet_name]
                headers = [(ws.cell(row=1, column=c).value or "").strip() for c in range(1, ws.max_column + 1)]
                try:
                    col_id = headers.index("ID") + 1
                    col_mtl = headers.index("MTL") + 1
                    col_edited = headers.index("Edited") + 1
                except ValueError:
                    continue
                total_lines = 0
                mtl_completed = 0
                edited_completed = 0
                for r in range(2, ws.max_row + 1):
                    row_id = (ws.cell(row=r, column=col_id).value or "").strip()
                    if not row_id:
                        continue
                    total_lines += 1
                    mtl = (ws.cell(row=r, column=col_mtl).value or "").strip()
                    edited = (ws.cell(row=r, column=col_edited).value or "").strip()
                    if mtl:
                        mtl_completed += 1
                    if edited:
                        edited_completed += 1
                chapter_total_lines += total_lines
                chapter_mtl_completed += mtl_completed
                chapter_edited_completed += edited_completed
                act_total_lines += total_lines
                act_mtl_completed += mtl_completed
                act_edited_completed += edited_completed
                total_all_lines += total_lines
                total_all_mtl += mtl_completed
                total_all_edited += edited_completed

                if 'adv' in sheet_name.lower():
                    adv_total += total_lines
                    adv_mtl += mtl_completed
                    adv_edited += edited_completed
                elif 'trial' in sheet_name.lower():
                    trial_total += total_lines
                    trial_mtl += mtl_completed
                    trial_edited += edited_completed
                elif 'bad' in sheet_name.lower():
                    bad_total += total_lines
                    bad_mtl += mtl_completed
                    bad_edited += edited_completed
                elif 'common' in sheet_name.lower():
                    common_total += total_lines
                    common_mtl += mtl_completed
                    common_edited += edited_completed

                act_display = act if act != last_act else ""
                chapter_display = f"Chapter{chapter}" if chapter != last_chapter or act != last_act else ""
                mtl_perc = (mtl_completed / total_lines * 100) if total_lines > 0 else 0
                edited_perc = (edited_completed / total_lines * 100) if total_lines > 0 else 0
                ov_ws.append([
                    act_display,
                    chapter_display,
                    sheet_name,
                    total_lines,
                    f"{mtl_perc:.2f}%",
                    f"{edited_perc:.2f}%"
                ])
                last_act = act
                last_chapter = chapter

            # Add per-file-type totals for the chapter
            if adv_total > 0:
                adv_mtl_perc = (adv_mtl / adv_total * 100) if adv_total > 0 else 0
                adv_edited_perc = (adv_edited / adv_total * 100) if adv_total > 0 else 0
                ov_ws.append([
                    "" if act == last_act else act,
                    "" if chapter == last_chapter and act == last_act else f"Chapter{chapter}",
                    "Adv Total",
                    adv_total,
                    f"{adv_mtl_perc:.2f}%",
                    f"{adv_edited_perc:.2f}%"
                ])
            if trial_total > 0:
                trial_mtl_perc = (trial_mtl / trial_total * 100) if trial_total > 0 else 0
                trial_edited_perc = (trial_edited / trial_total * 100) if trial_total > 0 else 0
                ov_ws.append([
                    "" if act == last_act else act,
                    "" if chapter == last_chapter and act == last_act else f"Chapter{chapter}",
                    "Trial Total",
                    trial_total,
                    f"{trial_mtl_perc:.2f}%",
                    f"{trial_edited_perc:.2f}%"
                ])
            if bad_total > 0:
                bad_mtl_perc = (bad_mtl / bad_total * 100) if bad_total > 0 else 0
                bad_edited_perc = (bad_edited / bad_total * 100) if bad_total > 0 else 0
                ov_ws.append([
                    "" if act == last_act else act,
                    "" if chapter == last_chapter and act == last_act else f"Chapter{chapter}",
                    "Bad Total",
                    bad_total,
                    f"{bad_mtl_perc:.2f}%",
                    f"{bad_edited_perc:.2f}%"
                ])
            if common_total > 0:
                common_mtl_perc = (common_mtl / common_total * 100) if common_total > 0 else 0
                common_edited_perc = (common_edited / common_total * 100) if common_total > 0 else 0
                ov_ws.append([
                    "" if act == last_act else act,
                    "" if chapter == last_chapter and act == last_act else f"Chapter{chapter}",
                    "Common Total",
                    common_total,
                    f"{common_mtl_perc:.2f}%",
                    f"{common_edited_perc:.2f}%"
                ])

            # Add chapter total
            if chapter_total_lines > 0:
                chapter_mtl_perc = (chapter_mtl_completed / chapter_total_lines * 100) if chapter_total_lines > 0 else 0
                chapter_edited_perc = (chapter_edited_completed / chapter_total_lines * 100) if chapter_total_lines > 0 else 0
                ov_ws.append([
                    "" if act == last_act else act,
                    f"Chapter{chapter} Total",
                    "",
                    chapter_total_lines,
                    f"{chapter_mtl_perc:.2f}%",
                    f"{chapter_edited_perc:.2f}%"
                ])

        # Add act total
        if act_total_lines > 0:
            act_mtl_perc = (act_mtl_completed / act_total_lines * 100) if act_total_lines > 0 else 0
            act_edited_perc = (act_edited_completed / act_total_lines * 100) if act_total_lines > 0 else 0
            ov_ws.append([
                f"{act} Total",
                "",
                "",
                act_total_lines,
                f"{act_mtl_perc:.2f}%",
                f"{act_edited_perc:.2f}%"
            ])

    # Add grand total including Common
    if total_all_lines > 0:
        total_mtl_perc = (total_all_mtl / total_all_lines * 100) if total_all_lines > 0 else 0
        total_edited_perc = (total_all_edited / total_all_lines * 100) if total_all_lines > 0 else 0
        ov_ws.append([
            "Grand Total",
            "",
            "",
            total_all_lines,
            f"{total_mtl_perc:.2f}%",
            f"{total_edited_perc:.2f}%"
        ])

def parse_original_files() -> None:
    if os.path.exists(XLSX_PATH):
        print(f"translate.xlsx already exists at {XLSX_PATH}. Skipping parse.")
        return

    if not os.path.isdir(ORIGINAL_DIR):
        print(f"Original directory not found: {ORIGINAL_DIR}")
        sys.exit(1)

    wb = Workbook()
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
            with open(path, 'r', encoding='cp932', errors='replace') as f:
                lines = f.readlines()

        ftype = detect_file_type(lines)
        if ftype is None:
            print(f"Warning: Could not detect file type for {fname}. Skipping.")
            continue

        data = parse_type1(lines) if ftype == 1 else parse_type2(lines)

        base_sheet_name = os.path.splitext(fname)[0]
        sheet_name = sanitize_sheet_name(base_sheet_name)
        existing = set(wb.sheetnames)
        if sheet_name in existing:
            suffix = 1
            while f"{sheet_name}_{suffix}" in existing:
                suffix += 1
            sheet_name = sanitize_sheet_name(f"{sheet_name}_{suffix}")
        ws = wb.create_sheet(title=sheet_name)

        ws.append(HEADER)
        style_and_freeze(ws)

        for _id, original, localized in data:
            ws.append([_id, original, trim_blank_lines(localized), "", ""])

        apply_wrap_to_all_cells(ws)

        metadata_rows.append([sheet_name, fname, ftype])

    meta_ws = wb.create_sheet(title="Metadata")
    meta_ws.append(META_HEADER)
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

    apply_wrap_to_all_cells(meta_ws)

    kb_ws = wb.create_sheet(title=KNOWLEDGE_SHEET)
    kb_ws.append(KNOWLEDGE_HEADER)
    kb_ws.cell(row=1, column=1).font = header_font
    kb_ws.cell(row=1, column=1).fill = fill
    kb_ws.freeze_panes = "A2"
    kb_ws.column_dimensions['A'].width = 100
    for line in INITIAL_PROJECT_HEADER:
        kb_ws.append([line])
    apply_wrap_to_all_cells(kb_ws)

    sum_ws = wb.create_sheet(title="Summaries")
    sum_ws.append(["Sheet name", "Summary"])
    for col_idx in range(1, 3):
        cell = sum_ws.cell(row=1, column=col_idx)
        cell.font = header_font
        cell.fill = fill
    sum_ws.column_dimensions['A'].width = 32
    sum_ws.column_dimensions['B'].width = 100
    apply_wrap_to_all_cells(sum_ws)

    update_overview(wb)

    wb.save(XLSX_PATH)
    print(f"translate.xlsx created at {XLSX_PATH}")

def get_knowledge_text(wb) -> str:
    if KNOWLEDGE_SHEET in wb.sheetnames:
        ws = wb[KNOWLEDGE_SHEET]
        parts = []
        for r in ws.iter_rows(min_row=2, max_col=1, values_only=True):
            if not r:
                continue
            v = (r[0] or "").strip()
            if v:
                parts.append(v)
        return "\n\n".join(parts).strip()
    return "\n\n".join(INITIAL_PROJECT_HEADER)

def generate_file_summary(client, sheet_name: str, rows: List[Tuple[str, str, str]]) -> str:
    context = "\n\n".join(
        f"ID: {row[0]}\nOriginal: {row[1] or '<empty>'}\nChinese: {row[2] or '<empty>'}"
        for row in rows
    )
    sys_prompt = (
        "You are a translator for a visual novel. Summarize the content of the following file in 2-3 concise lines, "
        "specifying the main context, key characters, and primary events. The summary must guide the tone and style of the translation "
        "(e.g., somber, emotional). Do not translate individual lines, only provide the summary in Vietnamese.\n\n"
        "Knowledge base (user-provided notes):\n" + get_knowledge_text(load_workbook(XLSX_PATH))
    )
    user_prompt = f"Sheet: {sheet_name}\n\nContent:\n{context}\n\nSummarize in 2-3 lines in Vietnamese."
    try:
        resp = client.responses.create(
            model=os.environ.get("OPENAI_MODEL", "gpt-5-mini"),
            reasoning=Reasoning(effort="medium"),
            instructions=sys_prompt,
            input=user_prompt
        )
        # print(f"Summary Input: ")
        # print(sys_prompt + "\n" + user_prompt)
        return (resp.output_text or "").strip()
    except Exception as e:
        print(f"Error generating summary for {sheet_name}: {e}")
        return ""

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

    sum_ws = wb["Summaries"] if "Summaries" in wb.sheetnames else None

    for sheet_name in wb.sheetnames:
        if sheet_name in {"Overview", "Metadata", KNOWLEDGE_SHEET, "Summaries"}:
            continue
        ws = wb[sheet_name]
        headers = [(ws.cell(row=1, column=c).value or "").strip() for c in range(1, 10)]
        try:
            col_id = headers.index("ID") + 1
            col_orig = headers.index("Original") + 1
            col_chinese = headers.index("Chinese") + 1
            col_mtl = headers.index("MTL") + 1
            col_edited = headers.index("Edited") + 1
        except ValueError:
            print(f"Warning: Sheet {sheet_name} has invalid headers. Skipping.")
            continue

        # Thu thập các dòng cần dịch
        rows_to_translate = []
        row_indices = []
        for r in range(2, ws.max_row + 1):
            if processed >= num_lines:
                break
            row_id = (ws.cell(row=r, column=col_id).value or "").strip()
            if not row_id:
                continue
            mtl_val = (ws.cell(row=r, column=col_mtl).value or "").strip()
            if mtl_val:
                continue
            original = ws.cell(row=r, column=col_orig).value or ""
            chinese = ws.cell(row=r, column=col_chinese).value or ""
            if not original and not chinese:
                continue
            rows_to_translate.append((row_id, original, chinese))
            row_indices.append(r)
            processed += 1

        if not rows_to_translate:
            continue

        # Tạo tóm tắt nếu cần
        summary = ""
        if sum_ws:
            for row in sum_ws.iter_rows(min_row=2, values_only=True):
                if row and row[0] == sheet_name:
                    summary = row[1] or ""
                    break
        if not summary:
            summary = generate_file_summary(client, sheet_name, rows_to_translate)
            if summary and sum_ws:
                sum_ws.append([sheet_name, summary])
                print(f"Summary for {sheet_name}: {summary}")

        # Tạo prompt duy nhất cho toàn bộ sheet
        prompt_lines = []
        for idx, (row_id, original, chinese) in enumerate(rows_to_translate, 1):
            prompt_lines.append(
                f"Line {idx}:\n"
                f"ID: {row_id}\n"
                f"Original value (source 1): {original or '<empty>'}\n"
                f"Chinese value (source 2): {chinese or '<empty>'}\n"
            )
        content = "\n".join(prompt_lines)
        sys_prompt = (
            f"Knowledge base (user-provided notes):\n{knowledge_text or '<empty>'}\n\n"
            f"File summary:\n{summary or '<no summary>'}\n\n"
            "You are a translator for a visual novel. Translate the following lines into Vietnamese. "
            "Return the translations in a numbered list corresponding to each line's index. "
            "Preserve placeholders, variables, control codes, line breaks, speaker tone, honorifics where appropriate, and context. "
            "Do not provide explanations, only the translations in the format:\n"
            "1. <translation>\n2. <translation>\n..."
        )
        user_prompt = f"Sheet: {sheet_name}\n\nContent:\n{content}\n\nTranslate into Vietnamese as a numbered list."

        try:
            resp = client.responses.create(
                model=model,
                reasoning=Reasoning(effort="medium"),
                instructions=sys_prompt,
                input=user_prompt
            )
            # print(f"Translate Input: ")
            # print(sys_prompt + "\n" + user_prompt)
            ai_text = (resp.output_text or "").strip()
            if not ai_text:
                print(f"Warning: Empty response for {sheet_name}")
                continue

            # print(ai_text)

            # Phân tích phản hồi thành danh sách các bản dịch
            translations = []
            current_translation = []
            current_num = None
            for line in ai_text.split('\n'):
                line = line.strip()
                if not line:
                    continue
                match = re.match(r'^(\d+)\.\s*(.*)$', line)
                if match:
                    if current_translation and current_num is not None:
                        translations.append((current_num, '\n'.join(current_translation).strip()))
                    current_num = int(match.group(1))
                    current_translation = [match.group(2).strip()]
                else:
                    if current_num is not None:
                        current_translation.append(line)
            if current_translation and current_num is not None:
                translations.append((current_num, '\n'.join(current_translation).strip()))

            # Gán bản dịch vào các ô tương ứng
            for num, translation in translations:
                if num > len(rows_to_translate):
                    print(f"Warning: Translation index {num} exceeds number of rows in {sheet_name}")
                    continue
                row_idx = row_indices[num - 1]
                if translation.lower().startswith("tóm tắt") or "summary" in translation.lower():
                    print(f"Warning: AI output for {sheet_name} | Line {num} contains summary: {translation}")
                    continue
                if translation:
                    ws.cell(row=row_idx, column=col_mtl).value = translation
                    print(f"Translated: {sheet_name} | ID {rows_to_translate[num - 1][0]}. Result: {translation}")

        except Exception as e:
            print(f"OpenAI API error on {sheet_name}: {e}")

        if processed >= num_lines:
            break

    if processed > 0 or (sum_ws and sum_ws.max_row > 1):
        update_overview(wb)
        wb.save(XLSX_PATH)
        print(f"Saved {processed} AI translations and summaries to {XLSX_PATH}")
    else:
        print("No rows required translation or already filled.")

def unpack_bundle(folder_path: str) -> None:
    folder = Path(folder_path)
    bundle_paths = list(folder.rglob("*.bundle"))

    if not bundle_paths:
        print(f"No .bundle files found in {folder_path}")
        return

    os.makedirs(ORIGINAL_DIR, exist_ok=True)

    print(f"Found {len(bundle_paths)} .bundle files to unpack:")

    for bundle_path in bundle_paths:
        try:
            bundle = UnityPy.load(str(bundle_path))
            for obj in bundle.objects:
                if obj.type.name == "TextAsset":
                    data = obj.read()
                    file_name = obj.container.split('/')[-1]
                    if not file_name.endswith(".txt"):
                        file_name += ".txt"
                    out_path = os.path.join(ORIGINAL_DIR, file_name)
                    os.makedirs(os.path.dirname(out_path), exist_ok=True)
                    with open(out_path, "w", encoding="utf-8-sig", newline="") as f:
                        f.write(data.m_Script)
                    print(f"Extracted {file_name} from {bundle_path}")
        except Exception as e:
            print(f"Error unpacking {bundle_path}: {e}")

def rebuild_translated_files() -> None:
    if not os.path.exists(XLSX_PATH):
        print(f"translate.xlsx not found at {XLSX_PATH}. Run parse first.")
        sys.exit(1)

    wb = load_workbook(XLSX_PATH)
    if "Metadata" not in wb.sheetnames:
        print("Metadata sheet not found in translate.xlsx")
        sys.exit(1)

    meta_ws = wb["Metadata"]
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
        id_rows = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row:
                continue
            _id = (row[0] or "").strip()
            if not _id:
                continue
            original = (row[1] or "")
            chinese = trim_blank_lines(row[2] or "")
            mtl = trim_blank_lines(row[3] or "")
            edited = trim_blank_lines(row[4] or "")
            id_rows.append((_id, original, chinese, mtl, edited))

        out_path = os.path.join(TRANSLATED_DIR, mapped_file)
        os.makedirs(os.path.dirname(out_path), exist_ok=True)

        lines_out: List[str] = []

        def add_comment_block(original_text: str, chinese_text: str, mtl_text: str):
            for ln in (original_text or "").replace('\r\n', '\n').replace('\r', '\n').split('\n'):
                lines_out.append("; " + ln if ln.strip() != "" else ";")
            lines_out.append("; **Chinese**")
            for ln in (chinese_text or "").replace('\r\n', '\n').replace('\r', '\n').split('\n'):
                lines_out.append("; " + ln if ln.strip() != "" else ";")
            lines_out.append("; **MTL**")
            for ln in (mtl_text or "").replace('\r\n', '\n').replace('\r', '\n').split('\n'):
                lines_out.append("; " + ln if ln.strip() != "" else ";")

        if ftype == 2:
            for _id, original, chinese, mtl, edited in id_rows:
                used_value = edited.strip() if edited.strip() != "" else mtl.strip() if mtl.strip() != "" else chinese
                used_value = trim_blank_lines(used_value)
                add_comment_block(original, chinese, mtl)
                lines_out.append(f"{_id}: {used_value}")
                lines_out.append("")
        elif ftype == 1:
            for _id, original, chinese, mtl, edited in id_rows:
                used_value = edited.strip() if edited.strip() != "" else mtl.strip() if mtl.strip() != "" else chinese
                used_value = trim_blank_lines(used_value)
                lines_out.append(f"# {_id}")
                add_comment_block(original, chinese, mtl)
                if used_value == "":
                    lines_out.append("")
                else:
                    for ln in used_value.split('\n'):
                        lines_out.append(ln)
                lines_out.append("")
        else:
            print(f"Warning: Unknown file type {ftype} for {mapped_file}. Skipping.")
            continue

        with open(out_path, 'w', encoding='utf-8-sig', newline='') as f:
            f.write("\n".join(lines_out).rstrip() + "\n")
        print(f"Wrote {out_path}")

def pack_translated_files(folder_path: str) -> None:
    folder = Path(folder_path)
    bundle_paths = list(folder.rglob("*.bundle"))

    if not bundle_paths:
        print(f"No .bundle files found in {folder_path}")
        return

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
                try:
                    rel_path = bundle_path.relative_to(folder)
                except ValueError:
                    rel_path = Path(bundle_path.name)
                backup_path = backup_folder / rel_path
                backup_path.parent.mkdir(parents=True, exist_ok=True)
                if not backup_path.exists():
                    shutil.copy2(bundle_path, backup_path)
                    print(f"Backed up original: {backup_path}")

                bundle.save(pack="lz4", out_path=os.path.dirname(bundle_path))
                print(f"Saved {bundle_path_str}")

        except Exception as e:
            print(f"Error processing {bundle_path}: {e}")

def refresh():
    if not os.path.exists(XLSX_PATH):
        print(f"translate.xlsx not found at {XLSX_PATH}.")
        sys.exit(1)

    wb = load_workbook(XLSX_PATH)

    # --- Kiểm tra file .txt mới trong ORIGINAL_DIR ---
    existing_sheets = set(wb.sheetnames)
    existing_meta = set()
    if "Metadata" in wb.sheetnames:
        for r in wb["Metadata"].iter_rows(min_row=2, values_only=True):
            if r and r[0]:
                existing_meta.add(str(r[0]))

    new_metadata_rows = []
    for fname in sorted(os.listdir(ORIGINAL_DIR)):
        if not fname.lower().endswith(".txt"):
            continue
        base_sheet_name = os.path.splitext(fname)[0]
        sheet_name = sanitize_sheet_name(base_sheet_name)

        if sheet_name in existing_sheets or sheet_name in existing_meta:
            continue  # đã có

        path = os.path.join(ORIGINAL_DIR, fname)
        try:
            with open(path, "r", encoding="utf-8") as f:
                lines = f.readlines()
        except UnicodeDecodeError:
            with open(path, "r", encoding="cp932", errors="replace") as f:
                lines = f.readlines()

        ftype = detect_file_type(lines)
        if ftype is None:
            print(f"Warning: Could not detect file type for {fname}. Skipping.")
            continue

        data = parse_type1(lines) if ftype == 1 else parse_type2(lines)

        # tạo sheet mới
        ws = wb.create_sheet(title=sheet_name)
        ws.append(HEADER)
        style_and_freeze(ws)
        for _id, original, localized in data:
            ws.append([_id, original, trim_blank_lines(localized), "", ""])
        apply_wrap_to_all_cells(ws)

        new_metadata_rows.append([sheet_name, fname, ftype])
        print(f"Added new sheet for {fname} -> {sheet_name}")

    # --- Cập nhật Metadata ---
    if new_metadata_rows:
        if "Metadata" not in wb.sheetnames:
            meta_ws = wb.create_sheet(title="Metadata")
            meta_ws.append(META_HEADER)
        else:
            meta_ws = wb["Metadata"]
        for row in new_metadata_rows:
            meta_ws.append(row)

    # --- Update lại Overview ---
    update_overview(wb)
    wb.save(XLSX_PATH)
    print(f"Refreshed {XLSX_PATH}. Added {len(new_metadata_rows)} new sheet(s).")

def main():
    if len(sys.argv) < 2:
        print("Usage: python translate_tool.py [parse|build|pack <folder>|build+pack <folder>|translate <num>|refresh]")
        sys.exit(1)
    cmd = sys.argv[1].lower()
    if cmd == 'parse':
        parse_original_files()
    elif cmd == 'unpack' and len(sys.argv) >= 3:
        unpack_bundle(sys.argv[2])
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
    elif cmd == 'refresh':
        refresh()
    else:
        print("Unknown command. Use 'parse', 'build', 'pack <folder>', 'build+pack <folder>', 'translate <num>', or 'refresh'.")
        sys.exit(1)

if __name__ == '__main__':
    main()