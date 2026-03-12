"""
filename_utils.py — shared filename sanitisation helpers.

Used by generate_all_reports, send_email, validate_reports, and app/main.py.
"""

import pandas as pd


def safe_filename(name) -> str:
    """Return a filesystem-safe version of *name* (spaces→underscores)."""
    if pd.isna(name) or name == "":
        return "Unknown"
    return "".join(
        c if c.isalnum() or c in (" ", "-") else "_" for c in str(name)
    ).replace(" ", "_")


def safe_display_name(name) -> str:
    """Return a display-safe version of *name* (strips illegal path chars)."""
    if pd.isna(name) or name == "":
        return "Unknown"
    s = str(name).strip()
    for old, new in (
        ("/", "-"),
        ("\\", "-"),
        (":", "-"),
        ("*", ""),
        ("?", ""),
        ('"', "'"),
        ("<", "("),
        (">", ")"),
        ("|", "-"),
    ):
        s = s.replace(old, new)
    return s
