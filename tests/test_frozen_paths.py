"""
test_frozen_paths.py — tests for frozen-app vs dev path resolution.

Covers utils.path_utils.get_user_base_dir() under all four conditions:
  - dev mode (default in test runner)
  - frozen + win32
  - frozen + linux
  - frozen + win32 with no APPDATA env var
"""

import importlib
import pathlib
import sys

import pytest

ROOT = pathlib.Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------


def _reimport(monkeypatch_or_none=None):
    """Force-reimport utils.path_utils so module-level state is fresh."""
    import utils.path_utils as m

    importlib.reload(m)
    return m


# ---------------------------------------------------------------------------
# dev mode
# ---------------------------------------------------------------------------


def test_dev_mode_returns_repo_root():
    """In dev mode (not frozen) get_user_base_dir() returns the repo root."""
    from utils.path_utils import get_user_base_dir

    result = get_user_base_dir()
    # In dev mode the result is two levels up from utils/ == repo root
    assert result == ROOT
    assert result.is_dir()


# ---------------------------------------------------------------------------
# frozen + win32
# ---------------------------------------------------------------------------


def test_frozen_win32_uses_appdata(monkeypatch, tmp_path):
    """Frozen+win32: result is %APPDATA%\\ResilienceScan."""
    fake_appdata = str(tmp_path / "AppData" / "Roaming")
    monkeypatch.setenv("APPDATA", fake_appdata)
    monkeypatch.setattr(sys, "frozen", True, raising=False)
    monkeypatch.setattr(sys, "platform", "win32")

    import utils.path_utils as m

    importlib.reload(m)
    result = m.get_user_base_dir()
    assert result == pathlib.Path(fake_appdata) / "ResilienceScan"


def test_frozen_win32_no_appdata_falls_back_to_home(monkeypatch):
    """Frozen+win32 with no APPDATA env: falls back to Path.home()."""
    monkeypatch.delenv("APPDATA", raising=False)
    monkeypatch.setattr(sys, "frozen", True, raising=False)
    monkeypatch.setattr(sys, "platform", "win32")

    import utils.path_utils as m

    importlib.reload(m)
    result = m.get_user_base_dir()
    # Path.home() is used as fallback
    assert result == pathlib.Path.home() / "ResilienceScan"


# ---------------------------------------------------------------------------
# frozen + linux
# ---------------------------------------------------------------------------


def test_frozen_linux_uses_xdg_path(monkeypatch):
    """Frozen+linux: result is ~/.local/share/resiliencescan."""
    monkeypatch.setattr(sys, "frozen", True, raising=False)
    monkeypatch.setattr(sys, "platform", "linux")

    import utils.path_utils as m

    importlib.reload(m)
    result = m.get_user_base_dir()
    assert result == pathlib.Path.home() / ".local" / "share" / "resiliencescan"


def test_frozen_linux_not_appdata(monkeypatch, tmp_path):
    """Frozen+linux should NOT use APPDATA even if set."""
    monkeypatch.setenv("APPDATA", str(tmp_path))
    monkeypatch.setattr(sys, "frozen", True, raising=False)
    monkeypatch.setattr(sys, "platform", "linux")

    import utils.path_utils as m

    importlib.reload(m)
    result = m.get_user_base_dir()
    assert "AppData" not in str(result)
    assert "resiliencescan" in str(result).lower()


# ---------------------------------------------------------------------------
# return type
# ---------------------------------------------------------------------------


def test_returns_path_object():
    """get_user_base_dir() always returns a pathlib.Path."""
    from utils.path_utils import get_user_base_dir

    result = get_user_base_dir()
    assert isinstance(result, pathlib.Path)


# ---------------------------------------------------------------------------
# cleanup: restore sys.frozen after monkeypatched tests
# ---------------------------------------------------------------------------
# pytest's monkeypatch fixture handles teardown automatically, but we must
# reload the module after each test that patched sys.frozen so subsequent
# tests see the original dev-mode behaviour.


@pytest.fixture(autouse=True)
def _reload_path_utils_after():
    yield
    import utils.path_utils as m

    importlib.reload(m)
