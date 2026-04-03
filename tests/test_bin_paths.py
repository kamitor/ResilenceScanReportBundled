"""
test_bin_paths.py — tests for utils.bin_paths binary resolution.
"""

import sys
from pathlib import Path
from unittest.mock import patch

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from utils.bin_paths import (  # noqa: E402
    _bundle_root,
    _search_roots,
    _vendor_root,
    build_r_env,
    find_quarto_bin,
    find_r_bin,
    find_r_library,
    find_tinytex_bin,
)


# ---------------------------------------------------------------------------
# Root helpers
# ---------------------------------------------------------------------------


def test_bundle_root_dev_is_repo_root():
    """In dev (non-frozen) mode _bundle_root() should return the repo root."""
    root = _bundle_root()
    assert root.is_dir()
    # The repo root contains pyproject.toml
    assert (root / "pyproject.toml").exists()


def test_vendor_root_is_vendor_subdir():
    """_vendor_root() should be <repo_root>/vendor."""
    vr = _vendor_root()
    assert vr.parent == _bundle_root()
    assert vr.name == "vendor"


def test_search_roots_dev_includes_vendor_first():
    """In dev mode the vendor/ directory should be checked before the bundle root."""
    roots = _search_roots()
    assert len(roots) >= 1
    # vendor/ is first in dev mode
    assert roots[0] == _vendor_root()


# ---------------------------------------------------------------------------
# find_r_bin
# ---------------------------------------------------------------------------


def test_find_r_bin_returns_str_or_none():
    """find_r_bin() must return a str path or None — never raise."""
    result = find_r_bin()
    assert result is None or isinstance(result, str)


def test_find_r_bin_finds_fake_bundle(tmp_path):
    """find_r_bin() returns the bundle binary when it exists."""
    exe = "Rscript.exe" if sys.platform == "win32" else "Rscript"
    fake_bin = tmp_path / "r" / "bin" / exe
    fake_bin.parent.mkdir(parents=True)
    fake_bin.touch()

    with patch("utils.bin_paths._search_roots", return_value=[tmp_path]):
        result = find_r_bin()

    assert result == str(fake_bin)


def test_find_r_bin_falls_back_to_system_path(tmp_path):
    """find_r_bin() falls back to shutil.which when no bundle binary exists."""
    with patch("utils.bin_paths._search_roots", return_value=[tmp_path]):
        with patch("shutil.which", side_effect=lambda name: f"/usr/bin/{name}"):
            result = find_r_bin()

    assert result == "/usr/bin/Rscript"


def test_find_r_bin_returns_none_when_nowhere(tmp_path):
    """find_r_bin() returns None when there is no bundle binary and R is not on PATH."""
    with patch("utils.bin_paths._search_roots", return_value=[tmp_path]):
        with patch("shutil.which", return_value=None):
            result = find_r_bin()

    assert result is None


# ---------------------------------------------------------------------------
# find_quarto_bin
# ---------------------------------------------------------------------------


def test_find_quarto_bin_returns_str_or_none():
    result = find_quarto_bin()
    assert result is None or isinstance(result, str)


def test_find_quarto_bin_finds_fake_bundle(tmp_path):
    exe = "quarto.exe" if sys.platform == "win32" else "quarto"
    fake_bin = tmp_path / "quarto" / "bin" / exe
    fake_bin.parent.mkdir(parents=True)
    fake_bin.touch()

    with patch("utils.bin_paths._search_roots", return_value=[tmp_path]):
        result = find_quarto_bin()

    assert result == str(fake_bin)


def test_find_quarto_bin_falls_back_to_system_path(tmp_path):
    with patch("utils.bin_paths._search_roots", return_value=[tmp_path]):
        with patch("shutil.which", return_value="/usr/local/bin/quarto"):
            result = find_quarto_bin()

    assert result == "/usr/local/bin/quarto"


# ---------------------------------------------------------------------------
# find_tinytex_bin
# ---------------------------------------------------------------------------


def test_find_tinytex_bin_returns_str_or_none():
    result = find_tinytex_bin()
    assert result is None or isinstance(result, str)


def test_find_tinytex_bin_finds_fake_bundle(tmp_path):
    import platform as _platform

    if sys.platform == "win32":
        arch, tlmgr = "windows", "tlmgr.bat"
    elif sys.platform == "darwin":
        arch, tlmgr = "universal-darwin", "tlmgr"
    else:
        arch = "aarch64-linux" if _platform.machine() == "aarch64" else "x86_64-linux"
        tlmgr = "tlmgr"

    fake_bin = tmp_path / "tinytex" / "bin" / arch / tlmgr
    fake_bin.parent.mkdir(parents=True)
    fake_bin.touch()

    with patch("utils.bin_paths._search_roots", return_value=[tmp_path]):
        result = find_tinytex_bin()

    assert result == str(fake_bin)


# ---------------------------------------------------------------------------
# find_r_library
# ---------------------------------------------------------------------------


def test_find_r_library_returns_path_or_none():
    result = find_r_library()
    assert result is None or isinstance(result, Path)


def test_find_r_library_finds_fake_bundle(tmp_path):
    fake_lib = tmp_path / "r-library"
    fake_lib.mkdir()

    with patch("utils.bin_paths._search_roots", return_value=[tmp_path]):
        result = find_r_library()

    assert result == fake_lib


def test_find_r_library_returns_none_when_absent(tmp_path):
    with patch("utils.bin_paths._search_roots", return_value=[tmp_path]):
        result = find_r_library()

    assert result is None


# ---------------------------------------------------------------------------
# build_r_env
# ---------------------------------------------------------------------------


def test_build_r_env_returns_dict():
    env = build_r_env()
    assert isinstance(env, dict)


def test_build_r_env_is_copy_not_same_object():
    """build_r_env() must not return the same dict as os.environ."""
    import os

    env = build_r_env(base_env=dict(os.environ))
    assert env is not os.environ


def test_build_r_env_sets_r_libs_with_bundled_library(tmp_path):
    """R_LIBS should be set to the bundled library when it exists."""
    fake_lib = tmp_path / "r-library"
    fake_lib.mkdir()

    with patch("utils.bin_paths._search_roots", return_value=[tmp_path]):
        env = build_r_env(base_env={})

    assert "R_LIBS" in env
    assert str(fake_lib) in env["R_LIBS"]


def test_build_r_env_sets_r_home_with_bundled_r(tmp_path):
    """R_HOME should be set when a bundled R installation exists."""
    fake_r = tmp_path / "r"
    fake_r.mkdir()

    with patch("utils.bin_paths._search_roots", return_value=[tmp_path]):
        env = build_r_env(base_env={})

    assert env.get("R_HOME") == str(fake_r)


def test_build_r_env_no_bundle_leaves_base_unchanged(tmp_path):
    """If no bundle dirs exist, base env should pass through unchanged."""
    base = {"PATH": "/usr/bin", "CUSTOM": "value"}

    with patch("utils.bin_paths._search_roots", return_value=[tmp_path]):
        env = build_r_env(base_env=base)

    assert env["CUSTOM"] == "value"
    assert "R_HOME" not in env
    assert "R_LIBS" not in env
