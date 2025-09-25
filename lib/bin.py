import sys


def _normalize_hex_to_bytes(hex_str: str) -> bytes:
    """Convert a hex string to bytes, ignoring spaces and underscores."""
    cleaned = hex_str.replace(" ", "").replace("_", "").strip()
    return bytes.fromhex(cleaned)


def validate_bin_patch_map(mapping: dict) -> dict:
    """Validate that mapping has equal-length pairs and return a dict[bytes, bytes]."""
    byte_map = {}
    for k, v in mapping.items():
        try:
            kb = _normalize_hex_to_bytes(str(k))
            vb = _normalize_hex_to_bytes(str(v))
        except ValueError:
            print(f"Invalid hex in mapping: {k} -> {v}")
            sys.exit(1)
        if len(kb) != len(vb):
            print(
                f"Source and destination hex must be the same length: {k} ({len(kb)} bytes) vs {v} ({len(vb)} bytes)"
            )
            sys.exit(1)
        if kb in byte_map:
            print(f"Duplicate source pattern detected in mapping: {k}")
            sys.exit(1)
        byte_map[kb] = vb
    return byte_map
