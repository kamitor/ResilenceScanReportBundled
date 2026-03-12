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
| M30 | Extract shared utilities: `utils/path_utils`, `utils/filename_utils`, `utils/constants`; wire into 6 files | v0.21.36 |
| M31 | Test coverage: `test_frozen_paths.py` (6), `test_csv_validation.py` (11), `test_shared_utils.py` (26), `test_email_send.py` (13), `test_thread_safety.py` (13) | v0.21.36 |
| M32 | Refactor `app/main.py` → 224 lines; 5 mixin modules (`gui_data`, `gui_generate`, `gui_email`, `gui_settings`, `gui_logs`); ruff suppression removed | v0.21.36 |
| M33 | Encoding safety — `encoding="utf-8"` on all text `open()` calls in pipeline + app files | v0.21.36 |
| M34 | Output folder writability validation: `_validate_output_folder()` called before generation starts | v0.21.38 |
| M35 | Log format standardisation in pipeline scripts — all prints use `{INFO,WARN,ERROR,OK}` | v0.21.39 |
| — | Installer hardening: `requireNamespace()` package checks; r-library ACL fix; `requirements_check.log` version validation | v0.21.37 |
| — | Installer bug-fixes from real-world test: Rscript version regex; pre-flight skip; SID-based ACLs; tlmgr self-update; Linux SETUP_RESULT; CODENAME fallback | v0.21.40 |
| — | Round-2 independent code review → `REVIEW2.md` (19 findings) | v0.21.40 |

**Current version: v0.21.40 — 201 tests, ruff clean**

---

## Active milestones

### ⏳ M36 — Frozen-app path fixes (REVIEW2.md §6) ← NEXT

Fix two remaining `ROOT_DIR` / relative-path misuses that will crash the frozen app.

**Findings to fix (see `REVIEW2.md` §6):**

| # | Finding | File | Severity |
|---|---|---|---|
| 6.1 | `email_template.json` saved to / loaded from `ROOT_DIR` — read-only `_internal/` in frozen app; `PermissionError` for any non-admin user who customises the template | `app/gui_email.py:478, 492` | HIGH |
| 6.2 | `run_integrity_validation()` reads results from `Path("./data/...")` — relative path fails in frozen app where CWD is not the data dir | `app/gui_data.py:652–653` | MEDIUM |

**Rules:**
- Replace `ROOT_DIR` with `_DATA_ROOT` in both `save_email_template()` and `load_email_template()`.
- Replace `Path("./data/integrity_validation_report.*")` with `_DATA_ROOT / "data" / "integrity_validation_report.*"`.
- Add `_DATA_ROOT` to the imports in each affected file.

**Gate:** All tests pass; ruff clean; frozen-app manual test confirms template saves to `APPDATA/ResilienceScan/` not `_internal/`.

---

### ⏳ M37 — Thread-safety residuals in generation + email threads (REVIEW2.md §1.1, §1.2, §1.3, §2.1, §2.2)

Fix remaining direct Tkinter widget writes in background threads and race conditions missed by M25.

**Findings to fix (see `REVIEW2.md` §1, §2):**

| # | Finding | File | Severity |
|---|---|---|---|
| 1.1 | `generate_reports_thread()`: unreachable duplicate `except Exception` at line 863; progress bar and label never updated on error; dead code has `# noqa: B025` suppression | `app/gui_generate.py:854–874` | HIGH |
| 2.1 | `generate_reports_thread()`: direct widget writes at lines 620–621 (`progress["maximum"]`, `progress["value"]`), 635–642 (`gen_current_label.config`), 858–874 (progress in exception handlers) — all outside `root.after()` | `app/gui_generate.py:620–874` | HIGH |
| 1.2 | `_send_emails_impl()` reads `self.df` directly from background thread — `self.df` can be replaced on the main thread at any time | `app/gui_email.py:1058–1074` | MEDIUM |
| 1.3 | `finalize()` closure accesses `self.df.columns` without a `None` guard — raises `AttributeError` if user reloads data during a send | `app/gui_email.py:1441` | MEDIUM |
| 2.2 | `is_generating` and `is_sending_emails` are plain booleans written from background threads; contradicts the M25 `threading.Event` pattern | `app/gui_generate.py`, `app/gui_email.py` | MEDIUM |

**Rules:**
- Collapse the duplicate `except Exception` into one; wrap all widget updates in `self.root.after(0, lambda: ...)`.
- Capture `self.df` as a local variable on the main thread before starting the email thread; pass it into `_send_emails_impl`.
- Add `if self.df is not None and` guard before `self.df.columns` in `finalize()`.
- Reset `is_generating` / `is_sending_emails` via `root.after(0, ...)` so writes happen on the main thread.

**Gate:** All tests pass; ruff clean; `# noqa: B025` suppression removed; no direct widget writes in background thread code paths.

---

### ⏳ M38 — Error handling + resource leaks in GUI email/generate (REVIEW2.md §1.4, §3, §5)

Harden the GUI email and generation paths to match the standards set for `send_email.py` in M28.

**Findings to fix (see `REVIEW2.md` §1.4, §3, §5):**

| # | Finding | File | Severity |
|---|---|---|---|
| 1.4 | `save_config()` and `load_config()` call `yaml.dump` / `yaml.safe_load` without checking `yaml is None` — raises `AttributeError` with no helpful error if PyYAML not installed | `app/gui_email.py:439, 451` | MEDIUM |
| 3.1 | `smtplib.SMTP()` constructed without `timeout=30` in the GUI send path (two sites) | `app/gui_email.py:1331, 1372` | MEDIUM |
| 3.2 | SMTP port `int()` cast not guarded with `try/except ValueError` in `save_config()` and `start_email_all()` | `app/gui_email.py:430, 942` | MEDIUM |
| 5.1 | SMTP `server` object not closed if `send_message()` raises — connection leaks until socket timeout | `app/gui_email.py:1331–1343, 1372–1376` | MEDIUM |
| 5.2 | Temp PDF not cleaned up in `generate_single_report_worker()` if `shutil.move()` fails (M28 fixed `generate_all_reports.py` but not the GUI path) | `app/gui_generate.py:342–480` | MEDIUM |
| 3.3 | Single broad `except Exception` covers all SMTP errors in the send loop — auth vs network vs config indistinguishable | `app/gui_email.py:1391` | LOW |

**Rules:**
- Add `if yaml is None: messagebox.showerror(...); return` at the top of `save_config()` and `load_config()`.
- Add `timeout=30` to both `smtplib.SMTP()` calls in `_send_emails_impl`.
- Wrap port cast in `try/except ValueError` at both sites; show `messagebox.showerror` on bad input.
- Use `with smtplib.SMTP(...) as server:` or add `try/finally: server.quit()` to both SMTP blocks.
- Wrap `subprocess.run` + `shutil.move` in `generate_single_report_worker` with `try/finally: temp_path.unlink(missing_ok=True)`.
- Add specific `except smtplib.SMTPAuthenticationError / smtplib.SMTPException / OSError` before the catch-all.

**Gate:** All tests pass; ruff clean; temp PDF always cleaned up; SMTP connection always closed.

---

### ⏳ M39 — Dead code, duplicate helpers, pandas encoding (REVIEW2.md §8, §10)

Remove remaining duplicate code and fix pandas encoding omissions.

**Findings to fix (see `REVIEW2.md` §8, §10):**

| # | Finding | File | Severity |
|---|---|---|---|
| 8.1 | `safe_filename()` / `safe_display_name()` still defined locally inside two methods in `gui_generate.py` — M30 extracted these to `utils/filename_utils.py` but missed this file | `app/gui_generate.py:284–302, 656–676` | MEDIUM |
| 10.1 | `pd.read_csv()` and `pd.to_csv()` calls lack `encoding="utf-8"` — on Windows cp1252 systems, names with accented characters cause mojibake or `UnicodeDecodeError` | multiple files | MEDIUM |
| 8.2 | `use_outlook = True` is set and never changed; the `else:` branch (direct SMTP when Outlook disabled) is ~30 lines of unreachable code | `app/gui_email.py:1193–1376` | LOW |
| 8.3 | `update_time()` and `show_about()` are defined in `DataMixin` but have no relation to data operations | `app/gui_data.py:1356–1381` | LOW |
| 10.2 | No comment in `generate_all_reports.py` documenting that it is a dev-only CLI tool (paths relative to repo root) | `generate_all_reports.py:13–15` | LOW |

**Rules:**
- Remove local `safe_filename` / `safe_display_name` from `gui_generate.py`; import from `utils.filename_utils`.
- Add `encoding="utf-8"` to all `pd.read_csv()` / `pd.to_csv()` calls that operate on `cleaned_master.csv` or any user-data CSV.
- Remove `use_outlook` variable and the unreachable `else:` SMTP block.
- Move `update_time()` and `show_about()` to `app/main.py`.
- Add `# NOTE: dev-only CLI tool` comment near path constants in `generate_all_reports.py`.

**Gate:** All tests pass; ruff clean; `grep -n "safe_filename\|safe_display_name" app/gui_generate.py` shows only the import line, not local definitions.

---

### ⏳ M40 — Test coverage gaps (REVIEW2.md §7)

Add tests for critical untested logic paths identified in the round-2 review.

**Findings to address (see `REVIEW2.md` §7):**

| # | Finding | Priority |
|---|---|---|
| 7.1 | No tests for `_sync_template()` copy logic — conditional mtime check, skip-when-current, run-when-dst-missing | MEDIUM |
| 7.2 | No tests for email filename parser in `_send_emails_impl()` — legacy names, company names with ` - `, SCROL filenames | MEDIUM |
| 7.3 | No tests for `_check_r_packages_ready()` — OK path, missing-package path, subprocess timeout | LOW |
| 7.4 | `test_send_emails_smtp_auth_error` assertion uses `or "FAIL"` fallback — too weak | LOW |

**Target test files:**
- `tests/test_frozen_paths.py` — extend with `_sync_template()` mtime scenarios
- `tests/test_email_send.py` — add filename parser tests; tighten existing auth-error assertion
- `tests/test_app_paths.py` (new) — mocked-subprocess tests for `_check_r_packages_ready()`

**Gate:** All new tests pass; `test_send_emails_smtp_auth_error` asserts `"Authentication error" in captured.out` without the `or "FAIL"` fallback; ruff clean.

---

### ⏳ M41 — Security + installer residuals (REVIEW2.md §4, §9.2)

Address remaining low-priority security and installer findings.

**Findings to fix (see `REVIEW2.md` §4, §9.2):**

| # | Finding | File | Severity |
|---|---|---|---|
| 4.1 | Hardcoded institution email addresses in Outlook account priority list | `app/gui_email.py:1206–1209` | LOW |
| 9.2 | `set -e` + `|| true` in `setup_linux.sh` silently swallows `Rscript not found` errors — `MISSING` stays empty, packages appear to pass | `packaging/setup_linux.sh:190, 195` | LOW |

**Rules:**
- Move `priority_accounts` list to `config.yml` under an `outlook_accounts` key; read at send time via `send_config`.
- In `setup_linux.sh`, replace `2>/dev/null || true` on Rscript verify calls with error output redirected to the log; add `command -v Rscript` guard before package checks.

**Gate:** All tests pass; ruff clean; no hardcoded email addresses in Python source files.
