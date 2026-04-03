"""
gui_system_check.py — verifies R, Quarto, and TinyTeX are available at runtime.

Called by the GUI at startup and via the System Check button.  Returns a
structured result so the GUI can display per-component pass/fail details.
"""

import os
import subprocess
import sys
from pathlib import Path

from utils.bin_paths import (
    find_quarto_bin,
    find_r_bin,
    find_r_library,
    find_tinytex_bin,
)
from utils.constants import R_SUBPROCESS_TIMEOUT

# All R packages required by ResilienceReport.qmd
_R_PACKAGES = [
    "readr",
    "dplyr",
    "stringr",
    "tidyr",
    "ggplot2",
    "knitr",
    "fmsb",
    "scales",
    "viridis",
    "patchwork",
    "RColorBrewer",
    "gridExtra",
    "png",
    "lubridate",
    "kableExtra",
    "rmarkdown",
    "jsonlite",
    "ggrepel",
    "cowplot",
]


# ---------------------------------------------------------------------------
# PATH helpers — the frozen app inherits PATH from the Windows Explorer process
# that was running at login, *before* the setup script updated the machine PATH.
# ---------------------------------------------------------------------------


def _refresh_windows_path() -> None:
    """Re-read machine + user PATH from the Windows registry and patch os.environ.

    The installer's setup script runs as SYSTEM *after* the user session has
    already started, so R and TinyTeX bin dirs added to the machine PATH are
    invisible to the running process.  Reading the registry directly picks them
    up without requiring a reboot or re-login.
    """
    if sys.platform != "win32":
        return
    try:
        import winreg  # noqa: PLC0415

        with winreg.OpenKey(
            winreg.HKEY_LOCAL_MACHINE,
            r"SYSTEM\CurrentControlSet\Control\Session Manager\Environment",
        ) as k:
            machine_path, _ = winreg.QueryValueEx(k, "PATH")
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, "Environment") as k:
                user_path, _ = winreg.QueryValueEx(k, "PATH")
        except OSError:
            user_path = ""
        # Registry values may contain unexpanded %vars%
        machine_path = os.path.expandvars(machine_path)
        user_path = os.path.expandvars(user_path)
        os.environ["PATH"] = machine_path + ";" + user_path
    except Exception:
        pass  # non-Windows or registry unavailable — leave PATH unchanged


def _find_rscript() -> str | None:
    """Find Rscript — bundle-first via utils.bin_paths, then system PATH."""
    return find_r_bin()


def _find_quarto() -> str | None:
    """Find quarto — bundle-first via utils.bin_paths, then system PATH."""
    return find_quarto_bin()


def _find_tlmgr() -> str | None:
    """Find tlmgr — bundle-first via utils.bin_paths, then system PATH."""
    return find_tinytex_bin()


def _setup_flag_dir() -> Path:
    """Return the directory where setup_running.flag / setup_complete.flag are written."""
    if sys.platform == "win32":
        return Path(os.environ.get("PROGRAMDATA", "C:/ProgramData")) / "ResilienceScan"
    return Path("/opt/ResilenceScanReportBuilder")


def setup_status() -> str:
    """Return the current background-setup state.

    Returns one of:
      'complete_pass' -- setup finished successfully
      'complete_fail' -- setup finished with errors
      'running'       -- setup started but has not finished yet
      'unknown'       -- no flag files found (dev mode, or flags not yet written)
    """
    flag_dir = _setup_flag_dir()
    complete = flag_dir / "setup_complete.flag"
    running = flag_dir / "setup_running.flag"
    try:
        if complete.exists():
            content = complete.read_text(encoding="utf-8").strip().upper()
            return "complete_pass" if "PASS" in content else "complete_fail"
        if running.exists():
            return "running"
    except OSError:
        pass
    return "unknown"


def _r_lib_path() -> Path | None:
    """Return the bundled R package library path via utils.bin_paths."""
    return find_r_library()


# ---------------------------------------------------------------------------
# Internal runner
# ---------------------------------------------------------------------------


def _run(cmd: list, env: dict | None = None) -> tuple[int, str]:
    """Run a command and return (returncode, combined stdout+stderr)."""
    try:
        r = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=R_SUBPROCESS_TIMEOUT,
            env=env,
        )
        return r.returncode, (r.stdout + r.stderr).strip()
    except Exception as e:
        return -1, str(e)


class SystemChecker:
    """Verifies runtime dependencies (R, Quarto, TinyTeX, R packages).

    Usage::

        checker = SystemChecker()
        result = checker.check_all()   # dict: {component: {"ok": bool, "version": str|None}}
        # checker.checks, checker.errors, checker.warnings populated as side-effects
    """

    def __init__(self, root_dir=None) -> None:
        # root_dir accepted for GUI compatibility but not required
        self.checks: list = []  # [{"item": str, "status": str, "description": str}]
        self.errors: list = []  # critical failures
        self.warnings: list = []  # non-critical issues

    # ------------------------------------------------------------------ public

    def check_all(self) -> dict:
        """Run all checks.

        Returns a dict compatible with the smoke test::

            {"python": {"ok": bool, "version": str|None}, "R": ..., ...}

        Also populates ``self.checks``, ``self.errors``, ``self.warnings``
        for the GUI system-check report.
        """
        # Refresh PATH from the Windows registry so that tools installed by
        # the setup script (which runs after user login) are discoverable.
        _refresh_windows_path()

        self.checks = []
        self.errors = []
        self.warnings = []

        result = {}
        result["python"] = self._check_python()
        result["R"] = self._check_r()
        result["quarto"] = self._check_quarto()
        result["tinytex"] = self._check_tinytex()
        result["r_packages"] = self._check_r_packages()
        return result

    # ----------------------------------------------------------------- private

    def _record(
        self,
        item: str,
        ok: bool,
        status: str,
        description: str = "",
        warning_only: bool = False,
    ) -> dict:
        """Append to self.checks and self.errors/warnings; return component dict."""
        self.checks.append({"item": item, "status": status, "description": description})
        if not ok:
            msg = f"{' '.join(item.split()[1:])}: {description or status}"
            if warning_only:
                self.warnings.append(msg)
            else:
                self.errors.append(msg)
        version = status if ok else None
        return {"ok": ok, "version": version}

    def _check_python(self) -> dict:
        ver = (
            f"{sys.version_info.major}.{sys.version_info.minor}"
            f".{sys.version_info.micro}"
        )
        return self._record(
            "[OK] Python",
            ok=True,
            status=f"Python {ver}",
        )

    def _check_r(self) -> dict:
        rscript = _find_rscript()
        if not rscript:
            return self._record(
                "[ERROR] R",
                ok=False,
                status="NOT FOUND",
                description="Rscript is not on PATH — R must be installed",
            )
        _, out = _run([rscript, "--version"])
        version = out.splitlines()[0] if out else "unknown"
        return self._record("[OK] R", ok=True, status=version)

    def _check_quarto(self) -> dict:
        quarto = _find_quarto()
        if not quarto:
            return self._record(
                "[ERROR] Quarto",
                ok=False,
                status="NOT FOUND",
                description="quarto is not on PATH — Quarto must be installed",
            )
        _, out = _run([quarto, "--version"])
        version = out.strip() if out else "unknown"
        return self._record("[OK] Quarto", ok=True, status=f"Quarto {version}")

    def _check_tinytex(self) -> dict:
        tlmgr = _find_tlmgr()
        if not tlmgr:
            return self._record(
                "[ERROR] TinyTeX",
                ok=False,
                status="NOT FOUND",
                description="tlmgr is not on PATH — run: quarto install tinytex",
            )
        # .bat files on Windows need cmd /c to execute correctly via subprocess
        if sys.platform == "win32" and tlmgr.lower().endswith(".bat"):
            cmd = ["cmd", "/c", tlmgr, "--version"]
        else:
            cmd = [tlmgr, "--version"]
        _, out = _run(cmd)
        version = out.splitlines()[0] if out else "unknown"
        return self._record("[OK] TinyTeX", ok=True, status=version)

    def _check_r_packages(self) -> dict:
        rscript = _find_rscript()
        if not rscript:
            self.checks.append(
                {
                    "item": "[SKIP] R packages",
                    "status": "SKIPPED — R not available",
                    "description": "",
                }
            )
            self.warnings.append("R packages not checked (R not available)")
            return {"ok": False, "version": None}

        pkg_list = ", ".join(f'"{p}"' for p in _R_PACKAGES)

        # Pass the bundled R library path (installed by the setup script) so
        # requireNamespace() finds packages even if R_LIBS is not set.
        r_lib = _r_lib_path()
        lib_expr = f'"{r_lib}"' if r_lib else "NULL"
        script = (
            f"lib <- {lib_expr}; "
            "if (!is.null(lib)) .libPaths(c(lib, .libPaths())); "
            f"pkgs <- c({pkg_list}); "
            "bad <- pkgs[!sapply(pkgs, requireNamespace, quietly=TRUE)]; "
            "if (length(bad) == 0) cat('OK') "
            "else cat('MISSING:', paste(bad, collapse=', '))"
        )

        _, out = _run([rscript, "-e", script])
        ok = out.strip() == "OK"
        if ok:
            return self._record(
                "[OK] R packages",
                ok=True,
                status=f"All {len(_R_PACKAGES)} required packages installed and loadable",
            )
        missing = out.replace("MISSING:", "").strip()
        return self._record(
            "[WARNING] R packages",
            ok=False,
            status="Missing packages",
            description=f"Missing: {missing}",
            warning_only=True,
        )
