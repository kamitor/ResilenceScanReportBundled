"""
filename_utils.py — shared filename sanitisation helpers.

Used by generate_all_reports, send_email, validate_reports, and app/main.py.
"""

from utils.constants import UNKNOWN_NAME_PLACEHOLDER


def _is_missing(name) -> bool:
    """Return True if name is None, NaN, pd.NA, or blank string."""
    if name is None:
        return True
    # float NaN: the only value not equal to itself (IEEE 754)
    if isinstance(name, float) and name != name:
        return True
    # pd.NA raises TypeError on bool(); treat as missing
    try:
        return not name or str(name).strip() == ""
    except TypeError:
        return True


def safe_filename(name) -> str:
    """Return a filesystem-safe version of *name* (spaces→underscores)."""
    if _is_missing(name):
        return UNKNOWN_NAME_PLACEHOLDER
    return "".join(
        c if c.isalnum() or c in (" ", "-") else "_" for c in str(name)
    ).replace(" ", "_")


def safe_display_name(name) -> str:
    """Return a display-safe version of *name* (strips illegal path chars)."""
    if _is_missing(name):
        return UNKNOWN_NAME_PLACEHOLDER
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
