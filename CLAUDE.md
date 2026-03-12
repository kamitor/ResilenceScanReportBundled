# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

```bash
# Set up Python environment
python -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt

# Run the GUI (main application)
python app/main.py

# Run pipeline steps individually
python clean_data.py                 # clean data/cleaned_master.csv in-place
python generate_all_reports.py       # render one PDF per row → reports/
python validate_reports.py           # validate generated PDFs against CSV values
python send_email.py                 # send PDFs (TEST_MODE=True by default)

# Lint and test (CI)
pip install pytest ruff pyyaml PyPDF2
ruff check .
ruff format --check .
pytest
pytest tests/test_smoke.py::test_import_main_module   # single test
pytest tests/test_pipeline_sample.py                  # anonymised fixture tests

# Regenerate the anonymised test fixture
python scripts/make_sample_data.py
```

## Release workflow

Bump `version` in `pyproject.toml` and push to `main`. CI detects no git tag `v<version>` exists and fires the build matrix. Do **not** create tags manually. macOS is not a target — only Windows and Linux matter.

---

## Architecture

`app/main.py` is the canonical entry point (Tkinter GUI + PyInstaller target).

```
app/main.py
  ├── imports convert_data          → Excel/ODS/XML/CSV/TSV → data/cleaned_master.csv
  ├── imports clean_data            → cleans and validates CSV in-place
  ├── imports email_tracker         → tracks per-recipient send status
  ├── imports gui_system_check      → verifies R/Quarto/TinyTeX are present at runtime
  ├── imports update_checker        → background GitHub release check
  └── imports dependency_manager    → stub (installation handled by the installer)
```

### Path resolution (frozen vs dev)

| Variable | Dev | Frozen (installed) |
|----------|-----|--------------------|
| `ROOT_DIR` / `_asset_root()` | repo root | `sys._MEIPASS` (`_internal/`) |
| `_DATA_ROOT` / `_data_root()` | repo root | `%APPDATA%\ResilienceScan` / `~/.local/share/resiliencescan` |
| `DATA_FILE` | `repo/data/cleaned_master.csv` | `APPDATA/data/cleaned_master.csv` |
| `REPORTS_DIR` | `repo/reports/` | `APPDATA/reports/` (temp write location only) |
| `DEFAULT_OUTPUT_DIR` | `repo/reports/` | `Documents\ResilienceScanReports\` |
| `TEMPLATE` | `repo/ResilienceReport.qmd` | `APPDATA/ResilienceScan/ResilienceReport.qmd` (copied from `_internal/` by `_sync_template()`) |

**Rule:** Any code that reads or displays reports must use `Path(self.output_folder_var.get())`, never `REPORTS_DIR`.

---

## Pipeline flow

```
data/*.xlsx  (or .xls, .ods, .xml, .tsv, .csv)
     │ convert_data.py
     ▼
data/cleaned_master.csv
     │ clean_data.py
     ▼
data/cleaned_master.csv  [validated & cleaned]
     │ generate_all_reports.py + ResilienceReport.qmd  (calls quarto render)
     ▼
reports/YYYYMMDD ResilienceScanReport (Company - Person).pdf
     │ validate_reports.py
     ▼
     │ send_email.py
     ▼
emails via Outlook COM (Windows) or SMTP fallback (Office365)
```

**Key data file:** `data/cleaned_master.csv`
**Score columns:** `up__r/c/f/v/a`, `in__r/c/f/v/a`, `do__r/c/f/v/a` — range 0–5
**PDF naming:** `YYYYMMDD ResilienceScanReport (Company Name - Firstname Lastname).pdf`

---

## Packaging strategy

**Staged installer** — the installer silently downloads and sets up all dependencies (R, Quarto, TinyTeX, R packages) during installation.

`ResilienceReport.qmd` is deeply LaTeX-dependent (TikZ, kableExtra, custom titlepage extension, custom fonts, raw `.tex` includes). The PDF engine **cannot** be switched to Typst or WeasyPrint — TinyTeX is required.

### Pinned dependency versions

| Dependency | Version | Notes |
|------------|---------|-------|
| R | 4.5.1 (pinned) | SYSTEM account has no network at install time; no auto-discovery |
| Quarto | 1.6.39 | GitHub releases |
| TinyTeX | Quarto-pinned | `quarto install tinytex` |
| Python | ≥ 3.11 | bundled by PyInstaller |

### R packages

`readr`, `dplyr`, `stringr`, `tidyr`, `ggplot2`, `knitr`, `fmsb`, `scales`, `viridis`, `patchwork`, `RColorBrewer`, `gridExtra`, `png`, `lubridate`, `kableExtra`, `rmarkdown`, `jsonlite`, `ggrepel`, `cowplot`

### LaTeX packages (tlmgr names)

`pgf`, `xcolor`, `colortbl`, `booktabs`, `multirow`, `float`, `wrapfig`, `pdflscape`, `geometry`, `preprint`, `graphics`, `tabu`, `threeparttable`, `threeparttablex`, `ulem`, `makecell`, `environ`, `trimspaces`, `caption`, `hyperref`, `setspace`, `fancyhdr`, `microtype`, `lm`, `needspace`, `varwidth`, `mdwtools`, `xstring`, `tools`

**Note:** `capt-of` is NOT installed via tlmgr (tar extraction fails on fresh TinyTeX). A minimal `capt-of.sty` stub is written directly by `setup_dependencies.ps1` / `setup_linux.sh` and registered with `mktexlsr`.

---

## Working rule

**Do not start the next milestone until the current one is fully verified by its gate condition.** Each gate must pass on a clean run before any work on the next milestone begins.

---

## Completed milestones

| Milestone | Description | Version |
|---|---|---|
| M1 | Fix CI, ship real app | v0.13.0 |
| M2 | Fix paths, consolidate cleaners | v0.14.0 |
| M3 | Implement data conversion | v0.15.0 |
| M4 | End-to-end report generation | v0.16.0 |
| M5 | Fix validation + email tracker | v0.17.0 |
| M6 | Email sending | v0.18.0 |
| M7 | Startup system check guard | v0.19.0 |
| M8 | Complete installer: R + Quarto + TinyTeX | v0.20.5 |
| M9 | Fix Windows installer: R path, LaTeX packages, capt-of | v0.20.14 |
| M10 | Fix report generation in installed app (frozen path split) | v0.21.0 |
| M11 | Anonymised sample dataset + pipeline smoke tests | v0.21.0 |
| M12 | End-to-end CI pipeline test (e2e.yml) | v0.21.0 |
| M13 | In-app update checker | v0.21.0 |
| M14 | README download badges + CI auto-update | v0.21.0 |
| M15 | Fix frozen app render failures (.quarto/ PermissionDenied, TinyTeX Quarto 1.4+, R_LIBS) | v0.21.4–v0.21.7 |
| M16 | Cross-platform test runner (platform.yml — Ubuntu + Windows on every push) | v0.21.14 |
| M17 | e2e CI passes on both platforms | v0.21.17 |
| M18 | Installer/version consistency tests; setup_linux.sh ASCII fix | v0.21.18 |
| M19 | Windows real-machine testing (Write-Log order, R pin, _version.py, output folder, cancel race, email folder) | v0.21.19–v0.21.25 |
| M20 | Setup completion feedback (sentinel flags, in-app polling, desktop notifications) | v0.21.26 |
| M21 | Fix email sending (thread-safe logging, send_config dict, except handler, tracker display) | v0.21.27 |
| M22 | R installer hardening + multi-format import (ODS/XML/CSV/TSV) | v0.21.28–v0.21.29 |
| M24 | Independent code analysis → `REVIEW.md` (27 findings) | v0.21.29 |
| M23 | SCROL matrix report template (`SCROLReport.qmd`); template dropdown; `_sync_template()` copies both QMDs | v0.21.30 |
| M25 | Thread-safety: `threading.Event` cancel, `threading.Lock` proc guard, all widget updates via `root.after(0,…)` | v0.21.31 |
| M26 | Frozen-app path fixes: `view_cleaning_report` uses `_DATA_ROOT`; `generate_executive_dashboard` removed (M27) | v0.21.31 |
| M27 | Dead-code removal: `generate_executive_dashboard()` + toolbar button deleted | v0.21.31 |
| M28 | Error handling: specific SMTP exceptions, SMTP port validation, temp PDF `finally` cleanup, debug logging in update_checker | v0.21.31 |
| M29 | Security: `TEST_EMAIL` → `test@example.com` default (env var override); SMTP timeout=30 | v0.21.31 |
| M30 | Extract shared utilities: `utils/path_utils`, `utils/filename_utils`, `utils/constants`; wire into 6 files | v0.21.32 |
| M31 | Test coverage: `test_frozen_paths.py` (6), `test_csv_validation.py` (11), `test_shared_utils.py` (26), `test_email_send.py` (13), `test_thread_safety.py` (13) | v0.21.33 |
| M32 | Refactor `app/main.py` → 224 lines; 5 mixin modules (`gui_data`, `gui_generate`, `gui_email`, `gui_settings`, `gui_logs`); ruff suppression removed | v0.21.34 |

**Current version: v0.21.34 — 201 tests, ruff clean**

---

## Active milestones

### ✅ M25 — Thread-safety fixes (REVIEW.md §4)

Fix all remaining direct widget access from background threads in `app/main.py`.

**Findings to fix (see `REVIEW.md` §4):**

| # | Finding | Severity |
|---|---|---|
| 4.1 | `generate_reports_thread()` calls `self.gen_current_label.config(...)`, `self.progress_var.set(...)` etc. directly from background thread — not thread-safe | HIGH |
| 4.2 | `self._gen_proc` TOCTOU race: written by generation thread, read/killed by main thread without a lock | MEDIUM |
| 4.3 | `self.is_generating` boolean shared across threads without `threading.Event` | LOW |

**Rules:**
- All widget updates inside a `threading.Thread` must go through `self.root.after(0, callback)`.
- Protect `self._gen_proc` with `threading.Lock()`.
- Replace `self.is_generating` flag with `threading.Event()`.

**Gate:** All 132+ tests pass; no Tkinter thread-safety warnings on Linux; cancel during generation does not crash.

---

### ⏳ M26 — Frozen-app path fixes (REVIEW.md §5)

Fix remaining `ROOT_DIR` / relative-path misuse in the frozen app.

**Findings to fix (see `REVIEW.md` §5):**

| # | Finding | File | Severity |
|---|---|---|---|
| 5.1 | `generate_executive_dashboard()` uses `cwd=ROOT_DIR` — read-only in frozen app; Quarto will crash creating `.quarto/` | `app/main.py` | HIGH |
| 5.2 | `view_cleaning_report()` uses `Path("./data/cleaning_report.txt")` — relative path fails in frozen app | `app/main.py` | MEDIUM |
| 5.3 | Standalone scripts (`generate_all_reports.py`, `clean_data.py`, `validate_reports.py`) use `Path(__file__).parent` as root — document as dev-only CLI tools | various | LOW |

**Rules:**
- Any `subprocess.run(..., cwd=...)` targeting a QMD render must use `cwd=str(_DATA_ROOT)`.
- Any file read/write outside of subprocess must use `_DATA_ROOT / ...` not relative paths.

**Gate:** All tests pass; ruff clean; frozen-app test (or manual install) confirms no PermissionDenied on the affected code paths.

---

### ✅ M27 — Dead-code and stub cleanup (REVIEW.md §1, §2)

Remove broken stubs and unreachable code identified in the code review.

**Findings to fix (see `REVIEW.md` §1–2):**

| # | Finding | File | Severity |
|---|---|---|---|
| 1.1 | `generate_executive_dashboard()` — references non-existent `ExecutiveDashboard.qmd`; button visible to users but always fails | `app/main.py` | MEDIUM |
| 2.1 | Unused imports in `generate_all_reports.py` | `generate_all_reports.py` | LOW |

**Rule:** Only remove code confirmed dead — no refactoring, no feature changes. If a stub button is removed, also remove its menu/toolbar entry.

**Gate:** All 132+ tests pass; ruff clean; no visible broken buttons in the UI.

---

### ✅ M28 — Error handling hardening (REVIEW.md §6)

Replace silent failures with specific, logged errors.

**Findings to fix (see `REVIEW.md` §6):**

| # | Finding | File | Severity |
|---|---|---|---|
| 6.1 | `update_checker.check_for_update()` — bare `except Exception: pass`; network/JSON errors silently ignored | `update_checker.py` | MEDIUM |
| 6.2 | `send_email.py` — single `except Exception` for all SMTP errors; auth vs network vs config errors indistinguishable | `send_email.py` | MEDIUM |
| 6.3 | `generate_all_reports.py` — no `finally` to clean up temp PDF if `shutil.move()` raises | `generate_all_reports.py` | MEDIUM |
| 6.4 | Multiple `except Exception: pass` blocks in `app/main.py` swallow debugging info | `app/main.py` | LOW |
| 6.5 | SMTP port `int()` cast not guarded — `ValueError` if user types non-numeric value | `app/main.py` | LOW |

**Rules:**
- Catch specific exception types (`smtplib.SMTPAuthenticationError`, `OSError`, `ValueError`, etc.).
- Log suppressed exceptions at DEBUG level where a log is available.
- Add `finally: temp_path.unlink(missing_ok=True)` in generation loops.

**Gate:** All tests pass; ruff clean; email auth error shown to user distinctly from network error.

---

### ✅ M29 — Security hardening (REVIEW.md §3)

Address security findings from the code review.

**Findings to fix (see `REVIEW.md` §3):**

| # | Finding | File | Severity |
|---|---|---|---|
| 3.1 | `TEST_EMAIL` hardcoded to a real person's address — if `TEST_MODE` is accidentally True, sends to a real person | `send_email.py` | MEDIUM |
| 3.2 | Output folder not validated for writability before use | `app/main.py` | LOW |
| 3.3 | `smtplib.SMTP()` has no timeout — hangs indefinitely if server unresponsive | `send_email.py` | LOW |

**Rules:**
- `TEST_EMAIL` must default to `test@example.com` (or load from env var `RESILIENCESCAN_TEST_EMAIL`).
- Output folder: validate it exists and is writable when the user changes it; show error if not.
- Add `timeout=30` to `smtplib.SMTP()`.

**Gate:** All tests pass; `test@example.com` default verified; no real email address in source.

---

### ✅ M30 — Extract shared utilities (REVIEW.md §7)

Eliminate copy-pasted code blocks by extracting shared helpers.

**Findings to fix (see `REVIEW.md` §7):**

| # | Finding | Files affected | Severity |
|---|---|---|---|
| 7.1 | `safe_filename()` / `safe_display_name()` duplicated in 4 files | `generate_all_reports.py`, `send_email.py`, `app/main.py` (×2), `validate_reports.py` | MEDIUM |
| 7.2 | User-data base-dir detection (`_user_base`) duplicated in 5 files | `convert_data.py`, `clean_data.py`, `email_tracker.py`, `send_email.py`, `validate_reports.py` | MEDIUM |
| 7.3 | Score-column constant (`up__r`, ...) defined in 2 files | `clean_data.py`, `app/main.py` | LOW |

**Target layout:**
```
utils/
  filename_utils.py   — safe_filename(), safe_display_name()
  path_utils.py       — get_user_base_dir() frozen/dev detection
  constants.py        — SCORE_COLUMNS, REQUIRED_COLUMNS
```

**Rule:** Import the shared helper everywhere; delete all copies. No behaviour change.

**Gate:** All 132+ tests pass; ruff clean; `utils/` is importable from all callers.

---

### ✅ M31 — Test coverage gaps (REVIEW.md §8)

Add tests for the untested areas identified in the code review.

**Findings to address (see `REVIEW.md` §8):**

| # | Finding | Priority |
|---|---|---|
| 8.1 | No tests for thread-safety: generation thread, cancel race, email thread | HIGH |
| 8.2 | No tests for frozen-app path logic (`_asset_root`, `_data_root`, `_sync_template`) with `monkeypatch` | HIGH |
| 8.3 | No tests for CSV validation / missing-column error path in `load_data_file()` | MEDIUM |
| 8.4 | Email send thread success path and tracker update not tested end-to-end | MEDIUM |

**Target test files:**
- `tests/test_thread_safety.py` — cancel race, `_gen_proc` guard, widget-update scheduling
- `tests/test_frozen_paths.py` — `monkeypatch` `sys.frozen`, `sys._MEIPASS`, platform checks
- `tests/test_csv_validation.py` — missing columns, wrong dtypes, extra columns
- `tests/test_email_send.py` — mocked SMTP, tracker state after send, status display update

**Gate:** Coverage on `gui_system_check.py`, `update_checker.py`, and path-resolution helpers reaches ≥ 80%; all new tests pass.

---

### ✅ M32 — Refactor `app/main.py` (REVIEW.md §9.1)

Break the ~3800-line monolithic file into focused modules.

**Scope (see `REVIEW.md` §9):**
- Extract generation logic → `app/gui_generate.py`
- Extract email logic → `app/gui_email.py`
- Extract data-load / filter logic → `app/gui_data.py`
- Extract system-check / settings logic → `app/gui_settings.py`
- `app/main.py` becomes the thin entry point: creates root, composes modules, starts mainloop.
- Remove the `"app/main.py" = ["E", "W", "F", "UP", "B"]` ruff suppression line once the file is small enough to lint normally.

**Rule:** No behaviour changes — tests must pass without modification. Gate condition requires the ruff suppression for `app/main.py` to be removed or significantly narrowed.

**Gate:** All tests pass; `app/main.py` ≤ 300 lines; ruff suppression removed or reduced to specific codes only.

---

### ⏳ M33 — Encoding safety (REVIEW.md §9.3) ← NEXT

Add `encoding="utf-8"` to all `open()` calls in pipeline scripts that are missing it.

**Findings:**
- Pipeline scripts (`generate_all_reports.py`, `clean_data.py`, `validate_reports.py`, `send_email.py`, etc.) have `open(path)` calls without `encoding=` — on Windows systems with non-UTF-8 ANSI code page (e.g. cp1252) this causes silent mojibake or `UnicodeDecodeError` when names contain non-ASCII characters.

**Rule:** Add `encoding="utf-8"` to every `open()` call that reads or writes text in pipeline scripts and the new `app/` mixin files. Binary opens (`"rb"`, `"wb"`) are exempt.

**Gate:** `grep -rn "open(" --include="*.py" | grep -v "encoding=" | grep -v '"rb"\|"wb"\|"ab"\|# noqa'` returns no hits in non-test pipeline files; all 201+ tests pass.

---

### ⏳ M34 — Output folder writability validation (REVIEW.md §3.2)

Validate the output folder before starting generation or sending emails.

**Finding:** `browse_output_folder()` stores the path without checking it exists and is writable. Generation silently fails with a confusing OS error if the folder is read-only or on a disconnected network share.

**Rule:** After any output folder change (browse or direct entry), resolve the path, attempt `mkdir(parents=True, exist_ok=True)`, and verify writable with a probe write. Surface a clear error via `messagebox.showerror` if not writable.

**Gate:** All tests pass; ruff clean; manual test: set output folder to `/tmp/readonly_test` (chmod 555) → error dialog shown before generation starts.

---

### ⏳ M35 — Log format standardisation (REVIEW.md §9.2)

Unify the log format across all pipeline scripts.

**Finding:** Pipeline scripts mix `print("[ERROR] ...")`, `print(f"[WARN] ...")`, raw `print(...)` with no timestamp, and `raise RuntimeError(...)` — no consistent level/timestamp scheme. The GUI swallows stdout from subprocesses into the generation log, making it hard for users to distinguish errors from progress messages.

**Rule:** All pipeline scripts (`generate_all_reports.py`, `clean_data.py`, `validate_reports.py`, `send_email.py`, `convert_data.py`) must use the format `[LEVEL] message` where LEVEL ∈ `{INFO, WARN, ERROR, OK}`. No change to the GUI log infrastructure. No new dependencies (no `logging` module required).

**Gate:** All tests pass; ruff clean; `grep -rn "^print(" --include="*.py"` in pipeline scripts shows no bare unlabelled prints.
