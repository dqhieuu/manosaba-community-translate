import re


def _parse_selector(selector: str):
    # Support consecutive indices, e.g., a[1][2].b[3]
    parts = selector.split('.') if selector else []
    tokens = []  # list of (name: str, indices: List[int])
    for part in parts:
        m = re.match(r"^(\w+)((\[\d+])*)$", part)
        if not m:
            tokens.append((part, []))
        else:
            name = m.group(1)
            idxs_str = m.group(2) or ""
            idxs = [int(mm.group(1)) for mm in re.finditer(r"\[(\d+)]", idxs_str)]
            tokens.append((name, idxs))
    return tokens


def set_by_selector(root, selector: str, value):
    tokens = _parse_selector(selector)
    if not tokens:
        return False
    cur = root
    parent = None
    parent_is_list = False
    key_or_index = None
    for (name, idxs) in tokens:
        # Access dict field by name
        if not isinstance(cur, dict) or name not in cur:
            return False
        parent = cur
        parent_is_list = False
        key_or_index = name
        cur = cur[name]
        # Apply consecutive indices, if any
        if idxs:
            for idx in idxs:
                if not isinstance(cur, list):
                    return False
                if idx < 0 or idx >= len(cur):
                    return False
                parent = cur
                parent_is_list = True
                key_or_index = idx
                cur = cur[idx]
    # Set value at the last resolved location
    if parent is None:
        return False
    if parent_is_list:
        if not isinstance(parent, list):
            return False
        idx = key_or_index
        if idx < 0 or idx >= len(parent):
            return False
        parent[idx] = value
        return True
    else:
        if not isinstance(parent, dict):
            return False
        parent[key_or_index] = value
        return True
