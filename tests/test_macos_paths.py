"""
test_macos_paths.py — tests for macOS frozen-app path resolution.

Unlike test_frozen_paths.py (which reloads modules), these tests call the
path-resolution functions directly after patching sys.frozen / sys.platform.
This avoids triggering the module-level code in app.app_paths (ROOT_DIR =
_asset_root() etc.) which would fail without a real sys._MEIPASS.
"""

import pathlib
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import app.app_paths as ap  # noqa: E402
import utils.path_utils as pu  # noqa: E402


# ---------------------------------------------------------------------------
# utils.path_utils.get_user_base_dir() — darwin branch
# ---------------------------------------------------------------------------


def test_get_user_base_dir_frozen_darwin(monkeypatch):
    """Frozen+darwin: get_user_base_dir() → ~/Library/Application Support/ResilienceScan."""
    monkeypatch.setattr(sys, "frozen", True, raising=False)
    monkeypatch.setattr(sys, "platform", "darwin")
    result = pu.get_user_base_dir()
    assert (
        result
        == pathlib.Path.home() / "Library" / "Application Support" / "ResilienceScan"
    )


def test_get_user_base_dir_frozen_darwin_returns_path(monkeypatch):
    """get_user_base_dir() always returns pathlib.Path, including on macOS."""
    monkeypatch.setattr(sys, "frozen", True, raising=False)
    monkeypatch.setattr(sys, "platform", "darwin")
    assert isinstance(pu.get_user_base_dir(), pathlib.Path)


def test_get_user_base_dir_frozen_darwin_not_appdata(monkeypatch, tmp_path):
    """Frozen+darwin must NOT use APPDATA even if the env var is set."""
    monkeypatch.setenv("APPDATA", str(tmp_path))
    monkeypatch.setattr(sys, "frozen", True, raising=False)
    monkeypatch.setattr(sys, "platform", "darwin")
    result = pu.get_user_base_dir()
    assert "AppData" not in str(result)
    assert "Library" in str(result)


def test_get_user_base_dir_frozen_darwin_not_xdg(monkeypatch):
    """Frozen+darwin must NOT return .local/share (that is the Linux path)."""
    monkeypatch.setattr(sys, "frozen", True, raising=False)
    monkeypatch.setattr(sys, "platform", "darwin")
    assert ".local" not in str(pu.get_user_base_dir())


def test_get_user_base_dir_dev_mode_darwin(monkeypatch):
    """Dev mode on darwin still returns the repo root, not macOS user dirs."""
    monkeypatch.setattr(sys, "frozen", False, raising=False)
    monkeypatch.setattr(sys, "platform", "darwin")
    assert pu.get_user_base_dir() == ROOT


# ---------------------------------------------------------------------------
# app.app_paths._data_root() — darwin branch (called directly, no module reload)
# ---------------------------------------------------------------------------


def test_data_root_frozen_darwin(monkeypatch):
    """app_paths._data_root() returns Library/Application Support/ResilienceScan on macOS."""
    monkeypatch.setattr(sys, "frozen", True, raising=False)
    monkeypatch.setattr(sys, "platform", "darwin")
    expected = (
        pathlib.Path.home() / "Library" / "Application Support" / "ResilienceScan"
    )
    assert ap._data_root() == expected


def test_data_root_frozen_darwin_returns_path(monkeypatch):
    """_data_root() always returns pathlib.Path."""
    monkeypatch.setattr(sys, "frozen", True, raising=False)
    monkeypatch.setattr(sys, "platform", "darwin")
    assert isinstance(ap._data_root(), pathlib.Path)


def test_data_root_frozen_darwin_not_appdata(monkeypatch, tmp_path):
    """Frozen+darwin must NOT use APPDATA."""
    monkeypatch.setenv("APPDATA", str(tmp_path))
    monkeypatch.setattr(sys, "frozen", True, raising=False)
    monkeypatch.setattr(sys, "platform", "darwin")
    result = ap._data_root()
    assert "AppData" not in str(result)
    assert "Library" in str(result)


def test_data_root_frozen_darwin_not_xdg(monkeypatch):
    """Frozen+darwin must NOT use .local/share (that is the Linux path)."""
    monkeypatch.setattr(sys, "frozen", True, raising=False)
    monkeypatch.setattr(sys, "platform", "darwin")
    assert ".local" not in str(ap._data_root())


def test_data_root_frozen_darwin_and_linux_differ(monkeypatch):
    """darwin and linux frozen paths must be distinct."""
    monkeypatch.setattr(sys, "frozen", True, raising=False)

    monkeypatch.setattr(sys, "platform", "darwin")
    darwin_path = ap._data_root()

    monkeypatch.setattr(sys, "platform", "linux")
    linux_path = ap._data_root()

    assert darwin_path != linux_path


def test_data_root_three_platforms_all_different(monkeypatch):
    """win32, linux, darwin each produce a distinct _data_root() when frozen."""
    monkeypatch.setattr(sys, "frozen", True, raising=False)
    monkeypatch.setenv("APPDATA", "/fake/appdata")

    paths = {}
    for platform in ("win32", "linux", "darwin"):
        monkeypatch.setattr(sys, "platform", platform)
        paths[platform] = str(ap._data_root())

    assert paths["win32"] != paths["linux"]
    assert paths["win32"] != paths["darwin"]
    assert paths["linux"] != paths["darwin"]


# ---------------------------------------------------------------------------
# app.app_paths._default_output_dir() — darwin branch
# ---------------------------------------------------------------------------


def test_default_output_dir_frozen_darwin_suffix(monkeypatch):
    """Frozen+darwin: default output dir ends with ResilienceScanReports."""
    monkeypatch.setattr(sys, "frozen", True, raising=False)
    monkeypatch.setattr(sys, "platform", "darwin")
    result = ap._default_output_dir()
    assert result.name == "ResilienceScanReports"


def test_default_output_dir_frozen_darwin_returns_path(monkeypatch):
    """_default_output_dir() returns pathlib.Path on macOS."""
    monkeypatch.setattr(sys, "frozen", True, raising=False)
    monkeypatch.setattr(sys, "platform", "darwin")
    assert isinstance(ap._default_output_dir(), pathlib.Path)


# ---------------------------------------------------------------------------
# app.app_paths._config_path() — darwin branch
# ---------------------------------------------------------------------------


def test_config_path_frozen_darwin(monkeypatch):
    """CONFIG_FILE lives inside the macOS data root."""
    monkeypatch.setattr(sys, "frozen", True, raising=False)
    monkeypatch.setattr(sys, "platform", "darwin")
    cfg = ap._config_path()
    expected_parent = (
        pathlib.Path.home() / "Library" / "Application Support" / "ResilienceScan"
    )
    assert cfg.parent == expected_parent
    assert cfg.name == "config.yml"


# ---------------------------------------------------------------------------
# app.app_paths._r_library_path() — frozen darwin
# ---------------------------------------------------------------------------


def test_r_library_path_frozen_darwin(monkeypatch, tmp_path):
    """_r_library_path() uses sys.executable.parent/r-library on macOS."""
    monkeypatch.setattr(sys, "frozen", True, raising=False)
    monkeypatch.setattr(sys, "platform", "darwin")
    fake_exe = tmp_path / "ResilenceScanReportBuilder"
    monkeypatch.setattr(sys, "executable", str(fake_exe))
    result = ap._r_library_path()
    assert result == tmp_path / "r-library"


# ---------------------------------------------------------------------------
# darwin path contains 'Library'
# ---------------------------------------------------------------------------


def test_darwin_path_contains_library(monkeypatch):
    """macOS data root always contains 'Library'."""
    monkeypatch.setattr(sys, "frozen", True, raising=False)
    monkeypatch.setattr(sys, "platform", "darwin")
    assert "Library" in str(ap._data_root())
