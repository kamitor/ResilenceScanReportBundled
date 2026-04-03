"""
test_app_paths.py — tests for app.app_paths._check_r_packages_ready().
"""

import subprocess
import sys
from pathlib import Path
from unittest.mock import MagicMock, patch

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))


# ---------------------------------------------------------------------------
# _check_r_packages_ready
# ---------------------------------------------------------------------------


def test_check_r_packages_ready_rscript_not_found(monkeypatch):
    """Returns an error string when R_BIN is None (no R in bundle or PATH)."""
    import app.app_paths as ap

    monkeypatch.setattr(ap, "R_BIN", None)
    result = ap._check_r_packages_ready()
    assert result is not None
    assert "not found" in result.lower() or "rscript" in result.lower()


def test_check_r_packages_ready_ok(monkeypatch):
    """Returns None when Rscript outputs 'OK'."""
    import app.app_paths as ap

    monkeypatch.setattr(ap, "R_BIN", "/usr/bin/Rscript")

    mock_result = MagicMock()
    mock_result.stdout = "OK"
    mock_result.stderr = ""
    with patch("subprocess.run", return_value=mock_result):
        result = ap._check_r_packages_ready()

    assert result is None


def test_check_r_packages_ready_missing_packages(monkeypatch):
    """Returns error string when Rscript reports missing packages."""
    import app.app_paths as ap

    monkeypatch.setattr(ap, "R_BIN", "/usr/bin/Rscript")

    mock_result = MagicMock()
    mock_result.stdout = "MISSING: fmsb, ggrepel"
    mock_result.stderr = ""
    with patch("subprocess.run", return_value=mock_result):
        result = ap._check_r_packages_ready()

    assert result is not None
    assert "MISSING" in result or "fmsb" in result


def test_check_r_packages_ready_subprocess_timeout(monkeypatch):
    """Returns error string when subprocess.run raises TimeoutExpired."""
    import app.app_paths as ap

    monkeypatch.setattr(ap, "R_BIN", "/usr/bin/Rscript")

    with patch(
        "subprocess.run", side_effect=subprocess.TimeoutExpired(cmd="R", timeout=30)
    ):
        result = ap._check_r_packages_ready()

    assert result is not None
    assert "error" in result.lower() or "timeout" in result.lower()


def test_check_r_packages_ready_subprocess_exception(monkeypatch):
    """Returns error string when subprocess.run raises any exception."""
    import app.app_paths as ap

    monkeypatch.setattr(ap, "R_BIN", "/usr/bin/Rscript")

    with patch("subprocess.run", side_effect=OSError("No such file")):
        result = ap._check_r_packages_ready()

    assert result is not None
    assert "error" in result.lower() or "no such file" in result.lower()
