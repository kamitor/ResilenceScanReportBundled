# REVIEW.md

## Code Quality Analysis: ResilenceScanReportBuilder

**Project**: Desktop Tkinter app for generating PDF reports via Quarto + R + TinyTeX
**Analysis Date**: 2026-03-10
**Analyst**: Independent automated review (Claude Code)

---

## 1. Dead Code

### 1.1 `generate_executive_dashboard()` — broken stub button
**File**: `app/main.py`
**Severity**: MEDIUM
**Description**: Function is wired to a visible Dashboard button but the referenced
`ExecutiveDashboard.qmd` does not exist and is never copied by `_sync_template()`.
Will always fail in the frozen app with a "file not found" Quarto error.
**Suggested fix**: Either (a) implement the template and add it to `_sync_template()`,
or (b) disable/hide the button until the feature is complete.

---

## 2. Unused Imports

### 2.1 `generate_all_reports.py`
**Severity**: LOW
Imports `pandas as pd` and `csv` at module level but the actual CSV read on line ~81 uses
`pd.read_csv()` directly. `csv` is used only for `csv.Sniffer` inside one helper; the import
can remain but is not immediately obvious.

---

## 3. Security Issues

### 3.1 Hardcoded test email address
**File**: `send_email.py`
**Severity**: MEDIUM
`TEST_EMAIL = "cg.verhoef@windesheim.nl"` is a real person's address committed to source.
If `TEST_MODE` is inadvertently left True in a release build, test sends go to this address.
**Suggested fix**: Load from environment variable or config; fall back to an obviously fake
address like `test@example.com`.

### 3.2 No output-folder validation
**File**: `app/main.py` — `browse_output_folder()`
**Severity**: LOW
User-selected output folder is stored without verifying it is a real, writable directory.
Subsequent `Path(output_folder_var.get()) / filename` can produce surprising paths if the
Tkinter dialog is bypassed by direct `output_folder_var.set()` calls.
**Suggested fix**: Validate that the resolved path is a directory and is writable before
accepting it.

### 3.3 SMTP connection has no timeout
**File**: `send_email.py`
**Severity**: LOW
`smtplib.SMTP(server, port)` is called without a timeout. If the server is unreachable the
send thread will hang indefinitely, blocking the UI.
**Suggested fix**: `smtplib.SMTP(server, port, timeout=30)`.

---

## 4. Thread-Safety Issues

### 4.1 Direct widget updates from `generate_reports_thread()`
**File**: `app/main.py`
**Severity**: HIGH
`generate_reports_thread()` runs in a `threading.Thread` (daemon). Several lines inside
the loop call `self.gen_current_label.config(text=...)`, `self.progress_var.set(...)`, etc.
directly from the background thread. Tkinter is not thread-safe; these calls can corrupt
widget state or crash silently on Linux/macOS.
**Suggested fix**: Replace every direct widget update inside the thread with
`self.root.after(0, lambda: widget.config(...))`.

### 4.2 `self._gen_proc` race condition
**File**: `app/main.py`
**Severity**: MEDIUM
`self._gen_proc` is written by the background generation thread and read by
`cancel_generation()` on the main thread. The v0.21.23 fix added a try/except but the
underlying TOCTOU race (check-then-use without a lock) remains.
**Suggested fix**: Protect with `threading.Lock()`.

### 4.3 `self.is_generating` unprotected boolean flag
**File**: `app/main.py`
**Severity**: LOW
`self.is_generating` is set from the main thread and read from the generation thread without
any synchronisation primitive. CPython's GIL makes this safe in practice, but a
`threading.Event` is the correct abstraction.

---

## 5. Frozen-App Path Correctness

### 5.1 `generate_executive_dashboard()` uses `ROOT_DIR` as `cwd`
**File**: `app/main.py`
**Severity**: HIGH
`subprocess.run(cmd, cwd=ROOT_DIR, ...)` — in a frozen app `ROOT_DIR` resolves to
`sys._MEIPASS` (`_internal/`), which is read-only under Program Files. Quarto will fail
immediately trying to create `.quarto/` there.
**Suggested fix**: `cwd=str(_DATA_ROOT)` — the same fix applied everywhere else in M15.

### 5.2 `view_cleaning_report()` uses relative paths
**File**: `app/main.py`
**Severity**: MEDIUM
`Path("./data/cleaning_report.txt")` is relative to cwd, which is unpredictable in a
frozen app. `clean_data.py` writes the report to `_DATA_ROOT / "data" / ...`.
**Suggested fix**: `_DATA_ROOT / "data" / "cleaning_report.txt"`.

### 5.3 Standalone scripts use `Path(__file__).parent` as root
**Files**: `generate_all_reports.py`, `clean_data.py`, `validate_reports.py`
**Severity**: LOW (dev/CLI scripts, not called from the frozen app directly)
These scripts hardcode `ROOT = Path(__file__).resolve().parent` which is correct for dev
but would break if imported from the frozen app. Document that they are dev-only CLI tools.

---

## 6. Error Handling Gaps

### 6.1 `update_checker.check_for_update()` — bare `except Exception: pass`
**File**: `update_checker.py`
**Severity**: MEDIUM
Network errors, JSON parse failures, and API changes are all silently swallowed. The GUI
simply never shows an update notification, giving no indication that the check failed.
**Suggested fix**: Log to stderr (non-frozen) or a debug log file.

### 6.2 `send_email.py` — single except for all SMTP errors
**File**: `send_email.py`
**Severity**: MEDIUM
`except Exception as e: failed_count += 1` catches auth errors, TLS failures, and
connection resets identically. Users cannot distinguish a config problem from a transient
network error.
**Suggested fix**: Catch `smtplib.SMTPAuthenticationError` separately and surface it
prominently.

### 6.3 Missing `finally` in `generate_all_reports.py` temp-file cleanup
**File**: `generate_all_reports.py`
**Severity**: MEDIUM
If `shutil.move(temp_path, final_path)` raises an exception, the temp PDF is left behind
in the reports directory with no cleanup.
**Suggested fix**: Add `finally: temp_path.unlink(missing_ok=True)`.

### 6.4 Silent `pass` in several except blocks
**File**: `app/main.py`
**Severity**: LOW
Multiple `except Exception: pass` blocks (update checker, cancel generation, misc widget
guards) suppress information that would be useful during debugging. Replace with at least
a `log(f"[DEBUG] suppressed: {e}")` call where the log is already available.

### 6.5 SMTP port not validated before `int()` cast
**File**: `app/main.py`
**Severity**: LOW
`int(self.smtp_port_var.get() or 587)` raises `ValueError` if the user types a non-numeric
value. No try/except around this cast; the exception propagates to the thread exception
handler and disables the send button silently.
**Suggested fix**: Validate on the main thread before launching the send thread.

---

## 7. Code Duplication

### 7.1 `safe_filename()` / `safe_display_name()` repeated in 4 files
**Files**: `generate_all_reports.py`, `send_email.py`, `app/main.py` (×2), `validate_reports.py`
**Severity**: MEDIUM
Near-identical filename-sanitisation logic is copy-pasted. Any bug fix must be applied in
all four places.
**Suggested fix**: Extract to `utils/filename_utils.py` and import.

### 7.2 User-data base-dir detection repeated in 5 files
**Files**: `convert_data.py`, `clean_data.py`, `email_tracker.py`, `send_email.py`,
`validate_reports.py`
**Severity**: MEDIUM
Each file contains nearly identical frozen/dev path detection:
```python
if getattr(sys, "frozen", False):
    if sys.platform == "win32":
        _user_base = Path(os.environ.get("APPDATA", ...)) / "ResilienceScan"
    else:
        _user_base = Path.home() / ".local" / "share" / "resiliencescan"
else:
    _user_base = Path(__file__).resolve().parent
```
If the platform logic changes, all five files must be updated in sync.
**Suggested fix**: Extract to `utils/path_utils.py`.

### 7.3 Score-column constant defined twice
**Files**: `clean_data.py` and `app/main.py`
**Severity**: LOW
The list of score columns (`up__r`, `up__c`, ..., `do__a`) is hardcoded in two places.
**Suggested fix**: Extract to `constants.py` and import in both.

---

## 8. Test Coverage Gaps

### 8.1 No tests for thread-safety
**Severity**: HIGH
The generation thread and email thread are both untested for concurrent access.
Race conditions in `_gen_proc` handling are not exercised by any test.

### 8.2 No tests for frozen-app path logic
**Severity**: HIGH
`_asset_root()`, `_data_root()`, and `_sync_template()` are untested. These are the most
fragile parts of the codebase (several past bugs originated here) yet rely entirely on
manual Windows/Linux install testing.
**Suggested fix**: Use `monkeypatch` to set `sys.frozen` and `sys._MEIPASS` in unit tests.

### 8.3 No tests for CSV validation / missing-column handling
**Severity**: MEDIUM
`load_data_file()` assumes required columns exist; no test exercises the error path for a
malformed CSV.

### 8.4 Email sending not tested end-to-end
**Severity**: MEDIUM
No test verifies the email send thread's success path. The tracker update and status-display
update are not tested together.

---

## 9. General Observations

### 9.1 Monolithic `app/main.py` (~3800 lines)
The entire GUI, generation logic, email sending, data loading, and system-check code lives
in a single class in a single file. This is intentionally suppressed from ruff linting.
Introduces high cognitive load and makes it difficult to write unit tests without spinning
up a full Tkinter root.
Addressed by milestone M26 (Refactor).

### 9.2 Inconsistent log format
Some modules use `print("[ERROR] ...")`, some use `self.log(...)`, some raise exceptions.
No central log level abstraction. Makes it hard for users to find or filter errors in logs.

### 9.3 No explicit encoding on all file opens
Several `open(path)` calls lack `encoding="utf-8"`. On Windows systems where the default
ANSI code page is not UTF-8, files with non-ASCII content (company names, person names)
can cause `UnicodeDecodeError`. Most critical files now specify encoding, but a few
legacy scripts still omit it.

### 9.4 `data/cleaned_master.csv` has no schema validation
The CSV is the single source of truth for all downstream processing. There is no schema
check at load time — missing columns, wrong dtypes, or extra columns are discovered late
(during quarto render, not at load). A validation step at `load_data_file()` would catch
problems early and give users actionable error messages.

---

## Summary

| Category | Issues | High | Medium | Low |
|---|---|---|---|---|
| Dead Code | 1 | 0 | 1 | 0 |
| Unused Imports | 1 | 0 | 0 | 1 |
| Security | 3 | 0 | 1 | 2 |
| Thread Safety | 3 | 1 | 1 | 1 |
| Frozen-App Paths | 3 | 1 | 1 | 1 |
| Error Handling | 5 | 0 | 3 | 2 |
| Code Duplication | 3 | 0 | 2 | 1 |
| Test Gaps | 4 | 2 | 2 | 0 |
| Observations | 4 | 0 | 1 | 3 |
| **TOTAL** | **27** | **4** | **12** | **11** |

### Recommended priority

**Immediate (before next Windows install test):**
- Fix `generate_reports_thread()` widget updates to use `root.after()`
- Fix `generate_executive_dashboard()` `cwd=ROOT_DIR` → `cwd=_DATA_ROOT`

**Soon (M25 — dead-code cleanup):**
- Remove/fix `generate_executive_dashboard()` stub
- Fix `view_cleaning_report()` relative paths
- Replace bare `except: pass` with specific exception types

**M26 — Refactor:**
- Extract `safe_filename()` and path-detection into shared utilities
- Break `app/main.py` into logical modules
- Add monkeypatched frozen-app tests
