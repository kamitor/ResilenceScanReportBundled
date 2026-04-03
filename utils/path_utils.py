"""
path_utils.py — shared user-data base directory detection.

Used by convert_data, clean_data, email_tracker, send_email, validate_reports.
Centralises the frozen-vs-dev path logic so it only lives in one place.
"""

import os
import sys
from pathlib import Path


def get_user_base_dir() -> Path:
    """Return the writable user-data base directory.

    Frozen app  (Windows) : %APPDATA%\\ResilienceScan
    Frozen app  (macOS)   : ~/Library/Application Support/ResilienceScan
    Frozen app  (Linux)   : ~/.local/share/resiliencescan
    Dev mode              : repo root (same directory as this file's parent)
    """
    if getattr(sys, "frozen", False):
        if sys.platform == "win32":
            return Path(os.environ.get("APPDATA", Path.home())) / "ResilienceScan"
        if sys.platform == "darwin":
            return Path.home() / "Library" / "Application Support" / "ResilienceScan"
        return Path.home() / ".local" / "share" / "resiliencescan"
    # dev: two levels up from utils/ → repo root
    return Path(__file__).resolve().parent.parent
