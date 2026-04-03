"""
Installer sanity tests — verify that all files and scripts needed for
the Windows/Linux installer to work are present and internally consistent.

These run on every push (no R/Quarto needed) and catch packaging regressions
before they reach the release build.
"""

import pathlib
import re
import sys
import tomllib

import pytest

ROOT = pathlib.Path(__file__).resolve().parent.parent


# ---------------------------------------------------------------------------
# pyproject.toml consistency
# ---------------------------------------------------------------------------


def _read_version() -> str:
    with open(ROOT / "pyproject.toml", "rb") as f:
        data = tomllib.load(f)
    return data["project"]["version"]


def test_version_is_semver():
    """pyproject.toml version must be a valid semver (X.Y.Z)."""
    version = _read_version()
    assert re.fullmatch(r"\d+\.\d+\.\d+", version), (
        f"Version {version!r} is not in X.Y.Z semver format"
    )


def test_version_readable_by_update_checker():
    """update_checker._current_version() must return the pyproject.toml version."""
    if str(ROOT) not in sys.path:
        sys.path.insert(0, str(ROOT))
    import update_checker

    vc = update_checker._current_version()
    pyproject_ver = _read_version()
    assert vc == pyproject_ver, (
        f"update_checker returns {vc!r}, pyproject.toml says {pyproject_ver!r}"
    )


# ---------------------------------------------------------------------------
# Required bundle assets (must exist so PyInstaller --add-data works)
# ---------------------------------------------------------------------------

REQUIRED_ASSETS = [
    "ResilienceReport.qmd",
    "references.bib",
    "QTDublinIrish.otf",
    "packaging/setup_dependencies.ps1",
    "packaging/launch_setup.ps1",
    "packaging/setup_macos.sh",
    "img",  # directory
    "tex",  # directory
    "_extensions",  # directory
]


@pytest.mark.parametrize("asset", REQUIRED_ASSETS)
def test_required_asset_exists(asset):
    """Every asset referenced in the PyInstaller --add-data flags must exist."""
    path = ROOT / asset
    assert path.exists(), (
        f"Required asset missing: {asset}  "
        f"(referenced by PyInstaller --add-data in ci.yml)"
    )


# ---------------------------------------------------------------------------
# setup_dependencies.ps1 — key marker checks (ASCII-only was verified by CI)
# ---------------------------------------------------------------------------


def test_ps1_setup_references_r_version():
    """setup_dependencies.ps1 must declare an R_VERSION variable."""
    ps1 = (ROOT / "packaging" / "setup_dependencies.ps1").read_text(encoding="utf-8")
    assert re.search(r"\$R_VERSION\s*=", ps1), (
        "setup_dependencies.ps1 does not define $R_VERSION"
    )


def test_ps1_setup_references_quarto():
    """setup_dependencies.ps1 must reference Quarto installation."""
    ps1 = (ROOT / "packaging" / "setup_dependencies.ps1").read_text(encoding="utf-8")
    assert "quarto" in ps1.lower(), "setup_dependencies.ps1 does not mention quarto"


def test_ps1_setup_references_tinytex():
    """setup_dependencies.ps1 must reference TinyTeX installation."""
    ps1 = (ROOT / "packaging" / "setup_dependencies.ps1").read_text(encoding="utf-8")
    assert "tinytex" in ps1.lower(), "setup_dependencies.ps1 does not mention tinytex"


def test_ps1_ascii_only():
    """setup_dependencies.ps1 must contain only ASCII characters.

    PowerShell 5.1 reads .ps1 files as Windows-1252. Non-ASCII bytes (e.g.
    the UTF-8 encoding of em dash E2 80 94) map to different characters in
    Windows-1252, silently corrupting string literals and causing parse errors.
    """
    ps1 = (ROOT / "packaging" / "setup_dependencies.ps1").read_text(encoding="utf-8")
    non_ascii = [(i, ch) for i, ch in enumerate(ps1) if ord(ch) > 127]
    if non_ascii:
        samples = non_ascii[:5]
        detail = ", ".join(f"offset {i} U+{ord(ch):04X} ({ch!r})" for i, ch in samples)
        pytest.fail(
            f"setup_dependencies.ps1 contains {len(non_ascii)} non-ASCII char(s): {detail}"
        )


# ---------------------------------------------------------------------------
# setup_dependencies.ps1 — function ordering
# ---------------------------------------------------------------------------


def test_ps1_write_log_defined_before_first_use():
    """Write-Log must be defined before any code that calls it.

    PS5.1 executes scripts top-to-bottom; calling a function before its
    'function' block is defined raises 'The term Write-Log is not recognized'.
    This was the root cause of the v0.21.18 installer crash on line 58.
    """
    ps1 = (ROOT / "packaging" / "setup_dependencies.ps1").read_text(encoding="utf-8")
    lines = ps1.splitlines()

    def_line = None
    first_call_line = None

    for i, line in enumerate(lines, start=1):
        stripped = line.strip()
        if def_line is None and re.match(r"function\s+Write-Log\b", stripped):
            def_line = i
        if first_call_line is None and re.search(r"\bWrite-Log\b", stripped):
            # Skip the function definition line itself and comment lines
            if not re.match(
                r"function\s+Write-Log\b", stripped
            ) and not stripped.startswith("#"):
                first_call_line = i

    assert def_line is not None, (
        "Write-Log function definition not found in setup_dependencies.ps1"
    )
    assert first_call_line is not None, (
        "No Write-Log call found in setup_dependencies.ps1"
    )
    assert def_line < first_call_line, (
        f"Write-Log is called on line {first_call_line} "
        f"but not defined until line {def_line} — "
        f"PS5.1 will crash with 'not recognized' error"
    )


# ---------------------------------------------------------------------------
# launch_setup.ps1
# ---------------------------------------------------------------------------


def test_launch_setup_ps1_exists_and_ascii():
    """launch_setup.ps1 must exist and be ASCII-only."""
    ps1 = (ROOT / "packaging" / "launch_setup.ps1").read_text(encoding="utf-8")
    non_ascii = [(i, ch) for i, ch in enumerate(ps1) if ord(ch) > 127]
    assert not non_ascii, f"launch_setup.ps1 has {len(non_ascii)} non-ASCII char(s)"


# ---------------------------------------------------------------------------
# PDF filename format (generated by generate_all_reports.py)
# ---------------------------------------------------------------------------


def test_pdf_filename_format():
    """generate_all_reports.safe_display_name must not introduce path separators."""
    if str(ROOT) not in sys.path:
        sys.path.insert(0, str(ROOT))
    import generate_all_reports as gar

    nasty_names = [
        "Acme/Corp",
        "Company\\Name",
        'Firm: "Premium"',
        "Org <special> | pipe",
    ]
    for name in nasty_names:
        result = gar.safe_display_name(name)
        # Must not contain characters illegal in Windows filenames
        illegal = set('/\\:*?"<>|')
        bad = [c for c in result if c in illegal]
        assert not bad, (
            f"safe_display_name({name!r}) = {result!r} still contains illegal char(s): {bad}"
        )
