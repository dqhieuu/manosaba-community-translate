import re


def is_alnum_start(s: str) -> bool:
    s = s.lstrip()
    return bool(re.match(r"^\w", s))


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
