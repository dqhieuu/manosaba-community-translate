import os
import sys


def is_file_editable(path: str) -> bool:
    """https://stackoverflow.com/a/37256114"""
    if not os.path.exists(path): return False
    try:
        os.rename(path, path)
        return True
    except OSError:
        return False


def is_running_in_exe() -> bool:
    return getattr(sys, 'frozen', False)