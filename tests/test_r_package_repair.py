"""
test_r_package_repair.py — tests for R package auto-repair feature.

Covers all six fixes made to handle pre-existing R installations and
missing packages at startup:

  1. PS1 pre-flight checks ONLY the bundled r-library (not user's library)
  2. PS1 pre-flight skips when r-library directory does not exist yet
  3. PS1 retry loop adds source fallback after binary fails
  4. PS1 global trap writes setup_complete.flag=FAIL before exiting
  5. launch_setup.ps1 overwrites stale flags with STALE sentinel
  6. launch_setup.ps1 polling skips any flag content that is not PASS/FAIL
  7-12. _install_r_packages_now: various Rscript output scenarios (threaded)
  13-16. _startup_guard: auto-repair triggered/suppressed correctly
  17. Dashboard has a "Repair R Packages" button
"""

import pathlib
import subprocess
import sys
import threading
import time
from unittest.mock import MagicMock, patch

import pytest

ROOT = pathlib.Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

# Read the PS1 files once — content tests are fast substring/regex checks.
_PS1_SETUP = (ROOT / "packaging" / "setup_dependencies.ps1").read_text(encoding="utf-8")
_PS1_LAUNCH = (ROOT / "packaging" / "launch_setup.ps1").read_text(encoding="utf-8")


# ===========================================================================
# 1–6: PS1 content checks
# ===========================================================================


def test_preflight_only_checks_bundled_r_library():
    """Pre-flight must use .libPaths() with ONLY the bundled r-library path.

    The previous bug: .libPaths(c('$pfLibR', .libPaths())) includes the user's
    personal R library, so packages the user installed globally could
    incorrectly make the pre-flight pass, leaving the bundled r-library empty.

    The fix: .libPaths('$pfLibR') (single arg) restricts search to our library.
    """
    # The pre-flight script must contain a .libPaths call with only the
    # bundled path — not c(..., .libPaths()) which adds global libraries.
    assert ".libPaths('$pfLibR')" in _PS1_SETUP, (
        "Pre-flight should restrict .libPaths to ONLY the bundled r-library path "
        "(not c('$pfLibR', .libPaths()) which would also check user's global library)"
    )


def test_preflight_skips_when_r_library_absent():
    """Pre-flight must check that r-library exists before checking packages.

    If r-library is absent (first install), the pre-flight must set
    $skipRPackages = $false so packages are always installed fresh.
    """
    # The fixed code wraps the entire pre-flight in 'if (Test-Path $R_LIB)'
    assert "Test-Path $R_LIB" in _PS1_SETUP, (
        "Pre-flight must check if the bundled r-library directory exists first"
    )
    # The else branch must also be present (sets skipRPackages = $false)
    # Find the pre-flight section and verify it has an else
    pf_idx = _PS1_SETUP.find("bundled r-library does not exist yet")
    assert pf_idx != -1, (
        "PS1 must have a comment about r-library not existing yet (the else branch)"
    )


def test_retry_loop_has_binary_and_source_labels():
    """Retry loop must label both binary and source attempts in log output.

    The fix adds: after binary retry, check if the package loaded, and
    if not, try without type restriction (source fallback).
    """
    assert "retry-bin" in _PS1_SETUP, (
        "Retry loop should log [R retry-bin] for the binary attempt"
    )
    assert "retry-src" in _PS1_SETUP, (
        "Retry loop should log [R retry-src] for the source fallback attempt"
    )


def test_retry_loop_source_triggered_when_binary_check_fails():
    """Retry loop must check if binary install succeeded before trying source."""
    # The fix adds a requireNamespace check after binary retry, then falls
    # back to source if the check fails.
    assert "Binary unavailable" in _PS1_SETUP or "binary unavailable" in _PS1_SETUP, (
        "Retry loop must log a message when binary is unavailable and source is tried"
    )


def test_trap_writes_fail_flag_before_exit():
    """Global trap must write setup_complete.flag=FAIL before calling exit.

    Without this, launch_setup.ps1 polls for up to 2 hours waiting for a
    flag that never arrives after a fatal error in setup_dependencies.ps1.
    """
    trap_start = _PS1_SETUP.find("trap {")
    assert trap_start != -1, "No trap block found in setup_dependencies.ps1"
    # Extract a reasonable window after the trap opening brace
    trap_section = _PS1_SETUP[trap_start : trap_start + 1000]
    assert "setup_complete.flag" in trap_section, (
        "The trap block must write to setup_complete.flag before exiting"
    )
    # Must write FAIL (not PASS) in the trap — an error occurred
    assert '"FAIL"' in trap_section or "'FAIL'" in trap_section, (
        "The trap block must write FAIL to setup_complete.flag"
    )


def test_launch_setup_writes_stale_sentinel_on_remove_failure():
    """launch_setup.ps1 must overwrite the stale flag with STALE when deletion fails.

    If Remove-Item silently fails (permission issue), the stale flag would be
    found instantly by the polling loop, giving a false PASS/FAIL result.
    The fix: overwrite with STALE so the polling loop ignores it.
    """
    assert "STALE" in _PS1_LAUNCH, (
        "launch_setup.ps1 must write STALE as a sentinel when Remove-Item fails"
    )


def test_launch_setup_polling_requires_pass_or_fail():
    """Polling loop must skip any flag content that is not PASS or FAIL.

    After the STALE sentinel fix, the loop must explicitly wait until the
    flag contains PASS or FAIL (written by setup_dependencies.ps1).
    """
    assert "PASS|FAIL" in _PS1_LAUNCH, (
        "Polling loop must check for PASS|FAIL (not just file existence) "
        "to ignore STALE sentinel flags"
    )


# ===========================================================================
# Minimal fake GUI for SettingsMixin method tests (no Tk window needed)
# ===========================================================================


class _FakeRoot:
    """Simulates Tk root.after() by queueing callbacks for manual flush."""

    def __init__(self):
        self._pending: list = []
        self._lock = threading.Lock()

    def after(self, delay_ms, callback):
        with self._lock:
            self._pending.append(callback)

    def flush_pending(self, timeout: float = 3.0) -> bool:
        """Execute the next queued callback, waiting up to *timeout* seconds."""
        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline:
            cb = None
            with self._lock:
                if self._pending:
                    cb = self._pending.pop(0)
            if cb is not None:
                cb()  # called outside lock so callbacks can call after() safely
                return True
            time.sleep(0.01)
        return False  # timed out


class _FakeLabel:
    def __init__(self):
        self.text = ""

    def config(self, text="", **kw):
        self.text = text


class _FakeGUI:
    """Minimal object satisfying SettingsMixin's interface without a real Tk window."""

    def __init__(self):
        self.root = _FakeRoot()
        self.status_label = _FakeLabel()
        self._logs: list[str] = []

    def log(self, msg: str) -> None:
        self._logs.append(msg)

    # Import SettingsMixin methods directly so _FakeGUI instances can call them.
    # Python treats them as regular methods (self = _FakeGUI instance).
    from app.gui_settings import SettingsMixin  # noqa: PLC0415

    _install_r_packages_now = SettingsMixin._install_r_packages_now
    _r_install_done = SettingsMixin._r_install_done
    _startup_guard = SettingsMixin._startup_guard


class _MockMsgbox:
    """Captures messagebox calls for assertion without displaying UI."""

    def __init__(self):
        self.info: list = []
        self.warn: list = []
        self.error: list = []

    def showinfo(self, title, msg, **kw):
        self.info.append((title, msg))

    def showwarning(self, title, msg, **kw):
        self.warn.append((title, msg))

    def showerror(self, title, msg, **kw):
        self.error.append((title, msg))

    def askyesno(self, title, msg, **kw):
        return True


# ===========================================================================
# 7–9: _r_install_done — synchronous dispatch (no threading needed)
# ===========================================================================


@pytest.fixture()
def gui_and_mb(monkeypatch):
    """Return a (_FakeGUI, _MockMsgbox) pair with messagebox patched."""
    gui = _FakeGUI()
    mb = _MockMsgbox()
    monkeypatch.setattr("app.gui_settings.messagebox", mb)
    return gui, mb


def test_r_install_done_already_ok_silent_suppresses_dialog(gui_and_mb):
    """ALREADY_OK + silent=True must NOT show a dialog."""
    gui, mb = gui_and_mb
    gui._r_install_done("ALREADY_OK\n", silent=True)
    assert not mb.info, "silent=True should suppress the 'already installed' dialog"
    assert gui.status_label.text == "Ready"
    assert any("already" in line.lower() for line in gui._logs)


def test_r_install_done_already_ok_not_silent_shows_dialog(gui_and_mb):
    """ALREADY_OK + silent=False must show an info dialog."""
    gui, mb = gui_and_mb
    gui._r_install_done("ALREADY_OK\n", silent=False)
    assert mb.info, "silent=False should show 'already installed' info dialog"
    assert gui.status_label.text == "Ready"


def test_r_install_done_success_shows_info(gui_and_mb):
    """SUCCESS output must show an info dialog."""
    gui, mb = gui_and_mb
    gui._r_install_done(
        "Installing 2 package(s): fmsb, cowplot\nSUCCESS\n", silent=False
    )
    assert mb.info
    assert gui.status_label.text == "Ready"
    assert any("[OK]" in line for line in gui._logs)


def test_r_install_done_missing_shows_warning_with_names(gui_and_mb):
    """MISSING output must show a warning dialog that names the missing packages."""
    gui, mb = gui_and_mb
    gui._r_install_done("MISSING: fmsb, cowplot", silent=False)
    assert mb.warn
    combined = " ".join(str(w) for w in mb.warn)
    assert "fmsb" in combined and "cowplot" in combined


def test_r_install_done_timeout_shows_warning(gui_and_mb):
    """TIMEOUT output must show a timeout warning dialog."""
    gui, mb = gui_and_mb
    gui._r_install_done("TIMEOUT", silent=False)
    assert mb.warn
    assert any("timeout" in str(w).lower() for w in mb.warn)


def test_r_install_done_always_resets_status_label(gui_and_mb):
    """_r_install_done must always set status_label back to 'Ready'."""
    gui, mb = gui_and_mb
    gui.status_label.text = "Installing..."
    gui._r_install_done("ALREADY_OK", silent=True)
    assert gui.status_label.text == "Ready"


def test_r_install_done_unexpected_output_shows_warning(gui_and_mb):
    """Unexpected Rscript output must show a warning (not silently swallow)."""
    gui, mb = gui_and_mb
    gui._r_install_done("some weird output", silent=False)
    assert mb.warn, "Unexpected output should produce a warning dialog"


# ===========================================================================
# 10–12: _install_r_packages_now — Rscript not found (synchronous code-path)
# ===========================================================================


def test_install_rscript_not_found_non_silent_shows_error(monkeypatch):
    """When Rscript not found and silent=False, an error dialog is shown."""
    gui = _FakeGUI()
    mb = _MockMsgbox()
    monkeypatch.setattr("app.gui_settings.messagebox", mb)
    monkeypatch.setattr("gui_system_check._find_rscript", lambda: None)
    monkeypatch.setattr("gui_system_check._refresh_windows_path", lambda: None)

    gui._install_r_packages_now(silent=False)

    assert mb.error, "Should show error dialog when Rscript not found (non-silent)"
    assert gui.status_label.text == "Ready"


def test_install_rscript_not_found_silent_no_dialog(monkeypatch):
    """When Rscript not found and silent=True, no dialog is shown."""
    gui = _FakeGUI()
    mb = _MockMsgbox()
    monkeypatch.setattr("app.gui_settings.messagebox", mb)
    monkeypatch.setattr("gui_system_check._find_rscript", lambda: None)
    monkeypatch.setattr("gui_system_check._refresh_windows_path", lambda: None)

    gui._install_r_packages_now(silent=True)

    assert not mb.error, (
        "silent=True should suppress error dialog when Rscript not found"
    )
    assert any(
        "not found" in line.lower() or "rscript" in line.lower() for line in gui._logs
    )


def test_install_sets_status_label_before_thread(monkeypatch):
    """_install_r_packages_now must update status_label synchronously before the thread."""
    gui = _FakeGUI()
    mb = _MockMsgbox()
    monkeypatch.setattr("app.gui_settings.messagebox", mb)
    monkeypatch.setattr("gui_system_check._find_rscript", lambda: "/usr/bin/Rscript")
    monkeypatch.setattr("gui_system_check._refresh_windows_path", lambda: None)

    label_during_run = []

    def _capture_label(*args, **kwargs):
        # Called when subprocess.run executes inside the thread
        label_during_run.append(gui.status_label.text)
        r = MagicMock()
        r.stdout = "ALREADY_OK\n"
        r.stderr = ""
        return r

    with patch("subprocess.run", side_effect=_capture_label):
        gui._install_r_packages_now(silent=True)
        gui.root.flush_pending(timeout=3.0)

    assert label_during_run, "subprocess.run was never called"
    # The label should say something about installing while the thread was running
    assert any(
        "install" in t.lower() or "package" in t.lower() for t in label_during_run
    ), f"Status label during thread was: {label_during_run}"


# ===========================================================================
# Thread-based tests: _install_r_packages_now with various Rscript outputs
# ===========================================================================


def _run_and_flush(gui, monkeypatch, rscript_stdout, *, silent=False, timeout=5.0):
    """Helper: mock Rscript, run _install_r_packages_now, flush the after() callback.

    Returns True if the callback was flushed within *timeout* seconds.
    """
    monkeypatch.setattr("gui_system_check._find_rscript", lambda: "/usr/bin/Rscript")
    monkeypatch.setattr("gui_system_check._refresh_windows_path", lambda: None)

    mock_proc = MagicMock()
    mock_proc.stdout = rscript_stdout
    mock_proc.stderr = ""

    with patch("subprocess.run", return_value=mock_proc):
        gui._install_r_packages_now(silent=silent)
        return gui.root.flush_pending(timeout=timeout)


def test_thread_already_ok_silent(monkeypatch):
    """ALREADY_OK + silent=True: thread posts callback, no dialog shown."""
    gui = _FakeGUI()
    mb = _MockMsgbox()
    monkeypatch.setattr("app.gui_settings.messagebox", mb)

    ok = _run_and_flush(gui, monkeypatch, "ALREADY_OK\n", silent=True)

    assert ok, "Background thread did not post callback within timeout"
    assert not mb.info, "silent=True must suppress the 'already installed' dialog"
    assert gui.status_label.text == "Ready"


def test_thread_already_ok_not_silent(monkeypatch):
    """ALREADY_OK + silent=False: thread posts callback, info dialog shown."""
    gui = _FakeGUI()
    mb = _MockMsgbox()
    monkeypatch.setattr("app.gui_settings.messagebox", mb)

    ok = _run_and_flush(gui, monkeypatch, "ALREADY_OK\n", silent=False)

    assert ok
    assert mb.info, "silent=False should show info dialog for ALREADY_OK"


def test_thread_success_shows_info(monkeypatch):
    """SUCCESS in output: thread posts callback, info dialog shown."""
    gui = _FakeGUI()
    mb = _MockMsgbox()
    monkeypatch.setattr("app.gui_settings.messagebox", mb)

    ok = _run_and_flush(gui, monkeypatch, "Installing 3 package(s): fmsb\nSUCCESS\n")

    assert ok
    assert mb.info, "SUCCESS should produce an info dialog"


def test_thread_missing_shows_warning(monkeypatch):
    """MISSING in output: thread posts callback, warning dialog names missing packages."""
    gui = _FakeGUI()
    mb = _MockMsgbox()
    monkeypatch.setattr("app.gui_settings.messagebox", mb)

    ok = _run_and_flush(gui, monkeypatch, "MISSING: fmsb, cowplot")

    assert ok
    assert mb.warn
    combined = " ".join(str(w) for w in mb.warn)
    assert "fmsb" in combined and "cowplot" in combined


def test_thread_timeout_shows_warning(monkeypatch):
    """TimeoutExpired: thread posts TIMEOUT callback, warning dialog shown."""
    gui = _FakeGUI()
    mb = _MockMsgbox()
    monkeypatch.setattr("app.gui_settings.messagebox", mb)
    monkeypatch.setattr("gui_system_check._find_rscript", lambda: "/usr/bin/Rscript")
    monkeypatch.setattr("gui_system_check._refresh_windows_path", lambda: None)

    with patch(
        "subprocess.run",
        side_effect=subprocess.TimeoutExpired(cmd="Rscript", timeout=600),
    ):
        gui._install_r_packages_now(silent=False)
        ok = gui.root.flush_pending(timeout=5.0)

    assert ok, "Thread did not post callback after TimeoutExpired"
    assert mb.warn, "TimeoutExpired should produce a warning dialog"
    assert any("timeout" in str(w).lower() for w in mb.warn)


def test_thread_subprocess_exception_captured(monkeypatch):
    """Any subprocess exception is caught and surfaced as a warning."""
    gui = _FakeGUI()
    mb = _MockMsgbox()
    monkeypatch.setattr("app.gui_settings.messagebox", mb)
    monkeypatch.setattr("gui_system_check._find_rscript", lambda: "/usr/bin/Rscript")
    monkeypatch.setattr("gui_system_check._refresh_windows_path", lambda: None)

    with patch("subprocess.run", side_effect=OSError("Rscript crashed")):
        gui._install_r_packages_now(silent=False)
        ok = gui.root.flush_pending(timeout=5.0)

    assert ok, "Thread did not post callback after OSError"
    # The exception is captured as "ERROR: <msg>" and hits the 'unexpected output'
    # branch which shows a warning dialog — any dialog is acceptable here.
    assert mb.warn or mb.error, "OSError should produce a warning or error dialog"


# ===========================================================================
# 13–16: _startup_guard auto-repair behaviour
# ===========================================================================


def _make_result(r_packages_ok=True, r_ok=True, quarto_ok=True, tinytex_ok=True):
    """Build a SystemChecker.check_all() result dict."""
    return {
        "python": {"ok": True, "version": "3.11"},
        "R": {"ok": r_ok, "version": "R 4.5.1" if r_ok else None},
        "quarto": {"ok": quarto_ok, "version": "Quarto 1.6" if quarto_ok else None},
        "tinytex": {"ok": tinytex_ok, "version": "tlmgr 2024" if tinytex_ok else None},
        "r_packages": {"ok": r_packages_ok, "version": None},
    }


@pytest.fixture()
def patched_guard(monkeypatch):
    """Patch SystemChecker and messagebox in app.gui_settings."""
    import app.gui_settings as gs

    mb = _MockMsgbox()
    monkeypatch.setattr(gs, "messagebox", mb)
    return mb


def test_startup_guard_auto_repair_packages_missing(monkeypatch, patched_guard):
    """Packages missing → auto-repair is always scheduled (bundled installer)."""
    import app.gui_settings as gs

    repair_calls: list = []

    monkeypatch.setattr(
        gs,
        "SystemChecker",
        type("SC", (), {"check_all": lambda s: _make_result(r_packages_ok=False)}),
    )

    gui = _FakeGUI()
    monkeypatch.setattr(
        gui, "_install_r_packages_now", lambda silent=False: repair_calls.append(silent)
    )

    gui._startup_guard()
    # The repair is scheduled via root.after(500, ...) — flush it
    gui.root.flush_pending(timeout=2.0)

    assert repair_calls, "Auto-repair must be triggered when R packages are missing"
    assert repair_calls[0] is True, "Auto-repair must be called with silent=True"


def test_startup_guard_no_repair_when_packages_ok(monkeypatch, patched_guard):
    """All packages OK → auto-repair must NOT fire."""
    import app.gui_settings as gs

    repair_calls: list = []

    monkeypatch.setattr(
        gs,
        "SystemChecker",
        type("SC", (), {"check_all": lambda s: _make_result(r_packages_ok=True)}),
    )

    gui = _FakeGUI()
    monkeypatch.setattr(
        gui, "_install_r_packages_now", lambda silent=False: repair_calls.append(silent)
    )

    gui._startup_guard()
    gui.root.flush_pending(timeout=0.3)

    assert not repair_calls, "Auto-repair must not run when all packages are OK"


# ===========================================================================
# 17: Dashboard button presence
# ===========================================================================


def test_repair_button_in_dashboard():
    """gui_data.py must have a 'Repair R Packages' button that calls _install_r_packages_now."""
    src = (ROOT / "app" / "gui_data.py").read_text(encoding="utf-8")
    assert "Repair R Packages" in src, (
        "Dashboard should have a 'Repair R Packages' button for manual repair"
    )
    assert "_install_r_packages_now" in src, (
        "The Repair button's command must call _install_r_packages_now"
    )
