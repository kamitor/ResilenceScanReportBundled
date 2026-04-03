"""
test_macos_installer.py — sanity checks for packaging/setup_macos.sh.

Mirrors the tests in test_installer_sanity.py and test_version_consistency.py
for the macOS installer script.  Runs on every push (no macOS required).
"""

import re
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
MACOS_SH = ROOT / "packaging" / "setup_macos.sh"
LINUX_SH = ROOT / "packaging" / "setup_linux.sh"
E2E = ROOT / ".github" / "workflows" / "e2e.yml"


# ---------------------------------------------------------------------------
# File presence and basic properties
# ---------------------------------------------------------------------------


def test_setup_macos_sh_exists():
    """packaging/setup_macos.sh must exist."""
    assert MACOS_SH.exists(), "packaging/setup_macos.sh not found"


def test_setup_macos_sh_is_shell_script():
    """setup_macos.sh must start with a bash shebang."""
    first_line = MACOS_SH.read_text(encoding="utf-8").splitlines()[0]
    assert first_line.startswith("#!/bin/bash"), (
        f"Expected bash shebang, got: {first_line!r}"
    )


def test_setup_macos_sh_ascii_only():
    """setup_macos.sh must contain only ASCII characters.

    Shell scripts with non-ASCII bytes can fail silently on locales that
    do not use UTF-8 (e.g. the POSIX/C locale used in some CI environments).
    """
    text = MACOS_SH.read_text(encoding="utf-8")
    non_ascii = [(i, ch) for i, ch in enumerate(text) if ord(ch) > 127]
    if non_ascii:
        samples = non_ascii[:5]
        detail = ", ".join(f"offset {i} U+{ord(ch):04X}" for i, ch in samples)
        pytest.fail(f"setup_macos.sh has {len(non_ascii)} non-ASCII char(s): {detail}")


# ---------------------------------------------------------------------------
# Key structural requirements
# ---------------------------------------------------------------------------


def test_setup_macos_sh_references_quarto_version():
    """setup_macos.sh must declare a QUARTO_VERSION variable."""
    text = MACOS_SH.read_text(encoding="utf-8")
    assert re.search(r'QUARTO_VERSION="[^"]+"', text), (
        "setup_macos.sh does not define QUARTO_VERSION"
    )


def test_setup_macos_sh_references_r_version():
    """setup_macos.sh must declare an R_VERSION variable."""
    text = MACOS_SH.read_text(encoding="utf-8")
    assert re.search(r'R_VERSION="[^"]+"', text), (
        "setup_macos.sh does not define R_VERSION"
    )


def test_setup_macos_sh_references_tinytex():
    """setup_macos.sh must reference TinyTeX installation."""
    text = MACOS_SH.read_text(encoding="utf-8")
    assert "tinytex" in text.lower(), "setup_macos.sh does not mention tinytex"


def test_setup_macos_sh_references_r_packages():
    """setup_macos.sh must define an R_PKGS variable."""
    text = MACOS_SH.read_text(encoding="utf-8")
    assert re.search(r"R_PKGS=", text), "setup_macos.sh does not define R_PKGS"


def test_setup_macos_sh_handles_arm64_and_x86():
    """setup_macos.sh must handle both arm64 and x86_64 R download URLs."""
    text = MACOS_SH.read_text(encoding="utf-8")
    assert "arm64" in text, "setup_macos.sh does not reference arm64 architecture"
    assert "x86_64" in text, "setup_macos.sh does not reference x86_64 architecture"


def test_setup_macos_sh_uses_installer_for_pkg():
    """setup_macos.sh must use the 'installer -pkg' command for .pkg files."""
    text = MACOS_SH.read_text(encoding="utf-8")
    assert "installer -pkg" in text, (
        "setup_macos.sh does not use 'installer -pkg' to install .pkg files"
    )


def test_setup_macos_sh_has_setup_result_flag():
    """setup_macos.sh must write a PASS/FAIL completion flag."""
    text = MACOS_SH.read_text(encoding="utf-8")
    assert "SETUP_RESULT" in text, "setup_macos.sh does not set SETUP_RESULT"
    assert "PASS" in text, "setup_macos.sh never sets result to PASS"
    assert "FAIL" in text, "setup_macos.sh never sets result to FAIL"


def test_setup_macos_sh_has_exit_trap():
    """setup_macos.sh must use a trap to write the completion flag on any exit."""
    text = MACOS_SH.read_text(encoding="utf-8")
    assert re.search(r"trap\s+\S+\s+EXIT", text), (
        "setup_macos.sh does not have a 'trap ... EXIT' handler"
    )


# ---------------------------------------------------------------------------
# Version sync: macOS Quarto version must match Linux and e2e.yml
# ---------------------------------------------------------------------------


def _quarto_version_macos_sh() -> str:
    text = MACOS_SH.read_text(encoding="utf-8")
    m = re.search(r'QUARTO_VERSION="([^"]+)"', text)
    return m.group(1) if m else ""


def _quarto_version_linux_sh() -> str:
    text = LINUX_SH.read_text(encoding="utf-8")
    m = re.search(r'QUARTO_VERSION="([^"]+)"', text)
    return m.group(1) if m else ""


def _quarto_version_e2e() -> str:
    text = E2E.read_text(encoding="utf-8")
    m = re.search(r"quarto-cli/releases/download/v([0-9]+\.[0-9]+\.[0-9]+)/", text)
    return m.group(1) if m else ""


def test_quarto_version_matches_linux_sh():
    """setup_macos.sh Quarto version must match setup_linux.sh."""
    macos_ver = _quarto_version_macos_sh()
    linux_ver = _quarto_version_linux_sh()
    assert macos_ver, "Could not extract Quarto version from setup_macos.sh"
    assert linux_ver, "Could not extract Quarto version from setup_linux.sh"
    assert macos_ver == linux_ver, (
        f"Quarto version mismatch: setup_macos.sh={macos_ver!r}, "
        f"setup_linux.sh={linux_ver!r}"
    )


def test_quarto_version_matches_e2e_yml():
    """setup_macos.sh Quarto version must match e2e.yml."""
    macos_ver = _quarto_version_macos_sh()
    e2e_ver = _quarto_version_e2e()
    assert macos_ver, "Could not extract Quarto version from setup_macos.sh"
    assert e2e_ver, "Could not extract Quarto version from e2e.yml"
    assert macos_ver == e2e_ver, (
        f"Quarto version mismatch: setup_macos.sh={macos_ver!r}, e2e.yml={e2e_ver!r}"
    )


# ---------------------------------------------------------------------------
# R package sync: macOS must include the same packages as Linux
# ---------------------------------------------------------------------------


def _extract_r_packages(text: str) -> set[str]:
    """Extract R package names from an R_PKGS='...' variable definition."""
    m = re.search(r"R_PKGS='([^']+)'", text) or re.search(r'R_PKGS="([^"]+)"', text)
    if not m:
        return set()
    return set(re.findall(r"'([A-Za-z][A-Za-z0-9.]+)'", m.group(1)))


def test_macos_r_packages_match_linux():
    """R package list in setup_macos.sh must match setup_linux.sh."""
    macos_pkgs = _extract_r_packages(MACOS_SH.read_text(encoding="utf-8"))
    linux_pkgs = _extract_r_packages(LINUX_SH.read_text(encoding="utf-8"))
    assert len(macos_pkgs) >= 15, (
        f"setup_macos.sh R package parser returned only {len(macos_pkgs)} packages"
    )
    assert len(linux_pkgs) >= 15, (
        f"setup_linux.sh R package parser returned only {len(linux_pkgs)} packages"
    )
    missing_from_macos = linux_pkgs - macos_pkgs
    assert not missing_from_macos, (
        f"R packages in setup_linux.sh but missing from setup_macos.sh: "
        f"{sorted(missing_from_macos)}"
    )
