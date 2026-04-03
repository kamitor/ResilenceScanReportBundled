"""
utils/bin_paths.py — bundle-first binary path resolution.

Finds R, Quarto, and TinyTeX binaries by checking the bundled vendor
directory first (vendor/ in dev, _internal/ in the frozen app), then
falling back to the system PATH.

This module is intentionally dependency-free with respect to the rest of
the app package so it can be imported freely from both gui_system_check
and app.app_paths without creating circular imports.
"""

import platform
import shutil
import sys
from pathlib import Path


# ---------------------------------------------------------------------------
# Root helpers
# ---------------------------------------------------------------------------


def _bundle_root() -> Path:
    """Root of the bundled assets directory.

    Frozen (installed): sys._MEIPASS — the _internal/ directory where
        PyInstaller extracts --add-data files.
    Dev:                repo root (one level up from utils/).
    """
    if getattr(sys, "frozen", False):
        return Path(sys._MEIPASS)
    return Path(__file__).resolve().parents[1]


def _vendor_root() -> Path:
    """Dev-mode vendor/ directory at the repo root.

    In the frozen app the bundle root already contains the vendored
    binaries so this is only consulted in dev mode.
    """
    return Path(__file__).resolve().parents[1] / "vendor"


def _search_roots() -> list[Path]:
    """Ordered list of directories to search for bundled binaries.

    Frozen: [_bundle_root()]          — only the bundle, no vendor/
    Dev:    [_vendor_root(), _bundle_root()]  — vendor/ first, then repo root
    """
    if getattr(sys, "frozen", False):
        return [_bundle_root()]
    return [_vendor_root(), _bundle_root()]


# ---------------------------------------------------------------------------
# Binary finders
# ---------------------------------------------------------------------------


def find_r_bin() -> str | None:
    """Find Rscript — bundled first, then system PATH.

    Expected bundle layout:
      Windows:  r/bin/Rscript.exe
      Unix:     r/bin/Rscript
    """
    exe = "Rscript.exe" if sys.platform == "win32" else "Rscript"
    for root in _search_roots():
        candidate = root / "r" / "bin" / exe
        if candidate.exists():
            return str(candidate)
    return shutil.which("Rscript") or shutil.which("R")


def find_quarto_bin() -> str | None:
    """Find quarto — bundled first, then system PATH.

    Expected bundle layout:
      Windows:  quarto/bin/quarto.exe
      Unix:     quarto/bin/quarto
    """
    exe = "quarto.exe" if sys.platform == "win32" else "quarto"
    for root in _search_roots():
        candidate = root / "quarto" / "bin" / exe
        if candidate.exists():
            return str(candidate)
    return shutil.which("quarto")


def find_tinytex_bin() -> str | None:
    """Find tlmgr — bundled first, then system PATH.

    TinyTeX binary layout under tinytex/bin/<arch>/:
      Windows:       windows/tlmgr.bat
      macOS:         universal-darwin/tlmgr
      Linux x86_64:  x86_64-linux/tlmgr
      Linux aarch64: aarch64-linux/tlmgr
    """
    if sys.platform == "win32":
        arch, tlmgr = "windows", "tlmgr.bat"
    elif sys.platform == "darwin":
        arch, tlmgr = "universal-darwin", "tlmgr"
    else:
        machine = platform.machine()
        arch = "aarch64-linux" if machine == "aarch64" else "x86_64-linux"
        tlmgr = "tlmgr"

    for root in _search_roots():
        candidate = root / "tinytex" / "bin" / arch / tlmgr
        if candidate.exists():
            return str(candidate)
    return shutil.which("tlmgr") or shutil.which("tlmgr.bat")


def find_r_library() -> Path | None:
    """Return the bundled R package library directory, or None if absent.

    Expected bundle layout: r-library/ at the root of the bundle.
    """
    for root in _search_roots():
        candidate = root / "r-library"
        if candidate.exists():
            return candidate
    return None


# ---------------------------------------------------------------------------
# Subprocess environment builder
# ---------------------------------------------------------------------------


def build_r_env(base_env: "dict | None" = None) -> dict:
    """Return a copy of the subprocess environment with R paths configured.

    Sets:
    - ``R_HOME`` to the bundled R root directory (if a bundled R is present).
    - ``R_LIBS`` prepended with the bundled R package library path.
    - ``LD_LIBRARY_PATH`` / ``DYLD_LIBRARY_PATH`` extended with the bundled
      R shared library directory (Linux / macOS only).

    If no bundled components are present the returned env is a plain copy of
    ``base_env`` (or ``os.environ``), so callers can always use this function
    unconditionally.
    """
    import os

    env = dict(base_env if base_env is not None else os.environ)

    # R_HOME — point at the bundled R root so R finds its modules and
    # libraries even when executed from a relocated installation path.
    for root in _search_roots():
        r_root = root / "r"
        if r_root.exists():
            env["R_HOME"] = str(r_root)
            # Add bundled R shared-library dir to dynamic-linker path (Unix).
            r_lib_dir = r_root / "lib"
            if r_lib_dir.exists() and sys.platform != "win32":
                ld_key = (
                    "DYLD_LIBRARY_PATH"
                    if sys.platform == "darwin"
                    else "LD_LIBRARY_PATH"
                )
                existing = env.get(ld_key, "")
                sep = ":"
                env[ld_key] = (
                    f"{r_lib_dir}{sep}{existing}" if existing else str(r_lib_dir)
                )
            break

    # R_LIBS — prepend bundled package library so R finds our packages first.
    r_pkgs = find_r_library()
    if r_pkgs is not None:
        existing = env.get("R_LIBS", "")
        sep = ";" if sys.platform == "win32" else ":"
        env["R_LIBS"] = f"{r_pkgs}{sep}{existing}" if existing else str(r_pkgs)

    return env
