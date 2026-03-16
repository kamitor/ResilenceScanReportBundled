"""
update_checker.py — checks GitHub releases for a newer version of the app.

Public API:
    check_for_update() -> dict | None
        Returns {"version": "x.y.z", "url": "https://..."} if a newer
        release exists on GitHub, None otherwise (including on any error).

    start_background_check(callback)
        Calls check_for_update() in a daemon thread and invokes callback(result)
        on the main thread via the Tkinter event loop.  callback receives the
        result dict or None.

Usage in the GUI:
    from update_checker import start_background_check

    def _on_update(info):
        if info:
            # show status bar message linking to info["url"]

    start_background_check(_on_update)
"""

import re
import sys
import threading
from pathlib import Path

# GitHub repository (owner/repo)
_GITHUB_REPO = "Windesheim-A-I-Support/ResilenceScanReportBuilder"
_API_URL = f"https://api.github.com/repos/{_GITHUB_REPO}/releases/latest"
_RELEASES_URL = f"https://github.com/{_GITHUB_REPO}/releases"
_TIMEOUT = 5  # seconds


def _current_version() -> str:
    """Return the running app's version from pyproject.toml (or _version.py).

    Falls back to "0.0.0" so the update check is always shown when the
    version cannot be determined (useful for dev testing).
    """
    # When frozen by PyInstaller, _version.py is baked in by CI.
    try:
        from app._version import __version__  # type: ignore[import]

        return __version__
    except ImportError:
        pass

    # Dev: parse pyproject.toml at the repo root
    try:
        if getattr(sys, "frozen", False):
            root = Path(sys._MEIPASS)
        else:
            root = Path(__file__).resolve().parent
        toml = (root / "pyproject.toml").read_text(encoding="utf-8")
        m = re.search(r'^version\s*=\s*["\']([^"\']+)', toml, re.M)
        if m:
            return m.group(1)
    except Exception:
        pass

    return "0.0.0"


def _parse_version(v: str) -> tuple[int, ...]:
    """Parse "1.2.3" → (1, 2, 3).  Non-numeric parts become 0."""
    parts = re.sub(r"[^0-9.]", "", v).split(".")
    return tuple(int(p) if p.isdigit() else 0 for p in parts)


def check_for_update() -> dict | None:
    """Query the GitHub releases API for a newer version.

    Returns:
        {"version": "x.y.z", "url": "<release URL>"} if newer, else None.
        Always returns None on any network / parse error.
    """
    try:
        import urllib.request
        import json as _json

        req = urllib.request.Request(
            _API_URL,
            headers={"User-Agent": "ResilienceScan-UpdateChecker/1.0"},
        )
        with urllib.request.urlopen(req, timeout=_TIMEOUT) as resp:
            data = _json.loads(resp.read().decode("utf-8"))

        tag = data.get("tag_name", "")
        latest_ver = tag.lstrip("v")
        release_url = data.get("html_url", _RELEASES_URL)

        if not latest_ver:
            return None

        current = _current_version()
        if _parse_version(latest_ver) > _parse_version(current):
            return {"version": latest_ver, "url": release_url}
    except Exception as e:
        if not getattr(sys, "frozen", False):
            print(
                f"[DEBUG] update check failed: {type(e).__name__}: {e}",
                file=sys.stderr,
            )

    return None


def start_background_check(callback, tk_root=None) -> None:
    """Run check_for_update() in a daemon thread.

    When the check completes, callback(result) is scheduled on the Tkinter
    main thread via tk_root.after(0, ...) if tk_root is provided, otherwise
    it is called directly from the background thread.

    Args:
        callback:  callable(dict | None)
        tk_root:   the Tk root window (optional but recommended for thread safety)
    """

    def _worker():
        result = check_for_update()
        if tk_root is not None:
            try:
                tk_root.after(0, callback, result)
            except Exception:
                pass
        else:
            try:
                callback(result)
            except Exception:
                pass

    t = threading.Thread(target=_worker, daemon=True, name="UpdateChecker")
    t.start()
