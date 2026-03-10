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
  ├── imports convert_data          → Excel/.xlsx → data/cleaned_master.csv
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
| `REPORTS_DIR` | `repo/reports/` | `APPDATA/reports/` |
| `TEMPLATE` | `repo/ResilienceReport.qmd` | `APPDATA/ResilienceScan/ResilienceReport.qmd` (copied from `_internal/` by `_sync_template()`) |

---

## Pipeline flow

```
data/*.xlsx  (or .xml)
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
| R | latest (auto-discovered from CRAN) | Current release URL, falls back to `/old/`; fallback pin: 4.5.1 |
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

## Milestones

### ✅ M1 — Fix CI, ship real app (v0.13.0)
### ✅ M2 — Fix paths, consolidate cleaners (v0.14.0)
### ✅ M3 — Implement data conversion (v0.15.0)
### ✅ M4 — End-to-end report generation (v0.16.0)
### ✅ M5 — Fix validation + email tracker (v0.17.0)
### ✅ M6 — Email sending (v0.18.0)
### ✅ M7 — Startup system check guard (v0.19.0)
### ✅ M8 — Complete installer: R + Quarto + TinyTeX (v0.20.5)
### ✅ M9 — Fix Windows installer: R path, LaTeX packages, capt-of (v0.20.14)
- R path: forward-slash conversion (`$R_LIB_R`) + single-quoted R package names in PS5.1
- LaTeX: corrected tlmgr package names; capt-of via stub + mktexlsr
- Startup guard: reads PATH from Windows registry + hardcoded fallbacks
- **Gate:** Fresh Windows VM — startup guard passes, all 19 R packages install, quarto render produces PDF ⏳ *pending re-test on v0.21.x build*

### ✅ M10 — Fix report generation in installed app (v0.21.0)
- File picker now accepts `.xlsx`/`.xls`; auto-copies + converts before loading
- Frozen path split: `_asset_root()` (QMD/images → `sys._MEIPASS`) vs `_data_root()` (CSV/reports/logs → APPDATA)
- Temp PDF output uses absolute `REPORTS_DIR` path (avoids writing to non-writable `_internal/`)
- **Gate:** Installed app generates correct PDF from real .xlsx — ⏳ *pending Windows test*

### ✅ M11 — Anonymised sample dataset (v0.21.0)
- `scripts/make_sample_data.py` + `tests/fixtures/sample_anonymized.xlsx` (3 fictional respondents)
- `tests/test_pipeline_sample.py` — 6 tests, all pass
- **Gate:** ✅ `pytest tests/test_pipeline_sample.py` passes on clean checkout

### ✅ M12 — End-to-end CI pipeline test (v0.21.0)
- `.github/workflows/e2e.yml` — `workflow_dispatch`, ubuntu + windows matrix
- Installs R/Quarto/TinyTeX, runs full pipeline on fixture, asserts PDFs produced
- **Gate:** ⏳ *pending first manual trigger*

### ✅ M13 — In-app update checker (v0.21.0)
- `update_checker.py` — GitHub releases API, semver compare, daemon thread
- Status bar shows clickable blue update link when newer version found
- Fixed startup crash: removed `ttk.Frame.cget("background")` (v0.21.1)
- **Gate:** ✅ App starts cleanly; update notification visible when version downgraded

### ✅ M14 — README download badges (v0.21.0)
- `Latest Release` shield badge + Downloads section with versioned links
- CI `update-readme` job patches links after each release and commits `[skip ci]`
- **Gate:** ✅ README updated by CI with correct v0.21.0 download URLs

### ✅ M15 — Fix frozen app render failures (v0.21.4 – v0.21.7)
**v0.21.4 — quarto .quarto/ PermissionDenied**
- **Root cause:** Quarto creates `.quarto/` next to the QMD. In frozen app the QMD is in `_internal/`
  (Windows: `Program Files\…\_internal\`; Linux: `/opt/…/_internal\`) — read-only → PermissionDenied.
- **Fix:** `_sync_template()` copies QMD + `img/`, `tex/`, `_extensions/`, `references.bib`,
  `QTDublinIrish.otf` from `_asset_root()` to `_data_root()` at startup (frozen only, skips if already
  up-to-date by mtime). Both render calls use `cwd=str(_DATA_ROOT)` and `selected_template = _DATA_ROOT / ...`.
- `stderr[:2000]` (was `[-500:]`) so root-cause error is visible in logs, not just the JS stack tail.
- Confirmed on Linux by reproducing `PermissionDenied: mkdir _internal/.quarto` with a read-only dir.

**v0.21.5 — TinyTeX detection for Quarto 1.4+**
- Quarto 1.4+ installs TinyTeX to `%APPDATA%\quarto\tools\tinytex\` (Windows) / `~/.local/share/quarto/tools/tinytex/` (Linux), not the legacy `TinyTeX\` location.
- Fixed in `gui_system_check._find_tlmgr()` (new candidates first), `setup_dependencies.ps1`, and `setup_linux.sh` (arch-dynamic path).

**v0.21.6 — e2e Windows TinyTeX PATH + stderr from start**
- `quarto install tinytex` only sets PATH for its own process. Added 'Add TinyTeX to PATH (Windows)' step in `e2e.yml` that writes the TinyTeX bin dir to `$GITHUB_PATH`.
- `generate_all_reports.py`: `stderr[:2000]` (was `[-1000:]`).

**v0.21.7 — R_LIBS missing from generate-single**
- `generate_single` was calling `subprocess.run` without `env=` so `R_LIBS` was never set.
- Single-report generation would fail to find R packages in the bundled `r-library/` while generate-all worked fine (it already built `gen_env`).
- Fix: build `single_env` with `R_LIBS` the same way generate-all does.

- **Why not caught earlier:** Dev mode + e2e both use system-wide R packages; bundled `r-library/` path only matters in the frozen installed app.
- **Gate:** Installed app generates correct PDF from real .xlsx on Windows ⏳ *pending Windows test*

---

## Next milestones

### ✅ M16 — Cross-platform test runner (v0.21.14)
`.github/workflows/platform.yml` — runs on every push + PR, no R/Quarto needed.

**Matrix:** `ubuntu-latest` × `windows-latest`

| Step | Ubuntu | Windows |
|---|---|---|
| `pytest` | ✅ | ✅ |
| App import smoke test | ✅ | ✅ |
| Pipeline dry run: convert → clean → verify CSV | ✅ | ✅ |
| PowerShell 5.1 syntax check (`shell: powershell`) | — | ✅ |

**Why it matters:**
- `pytest` previously only ran on Ubuntu — Windows path/encoding bugs could ship silently.
- Pipeline dry run validates the Python-only steps on both platforms on every push.
- `shell: powershell` invokes genuine PS5.1 (not PS7/pwsh); catches encoding and syntax issues
  that the existing PS7 check in `ci.yml` would miss (the v0.21.13 em-dash bug would have been
  caught here before reaching a real machine).

**Gate:** ✅ Both matrix jobs green on first run.

### ✅ M17 — e2e CI passes on both platforms (v0.21.17)
Full end-to-end pipeline (R/Quarto/TinyTeX) now reliably passes on Ubuntu + Windows.

**Key fixes applied:**
- YAML block scalar column-0 bug: printf/array join replace heredocs (run 4)
- bash pipefail + find|head: `|| true` on all pipelines (run 4)
- Windows cp1252 Unicode: `>=` not `≥`, `--` not `—` in validate_reports.py (run 4)
- GitHub API rate limit on TinyTeX: `GITHUB_TOKEN` env var (run 5)
- PS5.1 swallowing output: `shell: pwsh` (PS7) in Configure TinyTeX step (run 4)
- tlmgr symlink miss: removed `-type f` from `find`, added `which tlmgr` and `$HOME/.TinyTeX` to search (run 7)

**New tests added (89 total):**
- `test_package_sync.py` (6): R+LaTeX drift between e2e.yml and setup_dependencies.ps1
- `test_qmd_integrity.py` (6): QMD YAML, R packages in install lists, params
- `test_report_generation.py` (13): quarto command structure (mocked subprocess)
- `test_unicode_safety.py` (18): non-ASCII in print()/raise in pipeline scripts
- `test_workflow_yaml.py` (16): YAML validity, structure, column-0 bug detection
- `test_installer_sanity.py` (16): assets, version, PS1 consistency

**Gate:** ✅ e2e.yml run 7 — both Ubuntu and Windows generate PDF artifacts.

### ✅ M18 — Installer/version consistency tests + setup_linux.sh fix (v0.21.18)

**New tests added (121 total, +32):**
- `test_version_consistency.py` (9): Quarto version in sync across setup_linux.sh,
  setup_dependencies.ps1, e2e.yml; setup_linux.sh ASCII-only; Linux R packages vs e2e.yml and PS1
- `test_update_checker.py` (23): `_parse_version` semver, version comparison, `check_for_update`
  with mocked network responses, `start_background_check` callback

**Bug found and fixed by new tests:**
- `setup_linux.sh` contained 322 non-ASCII chars (U+2014 em dash, U+2500 box-drawing, U+2192 arrow)
  — could cause encoding errors on C-locale systems. Fixed with ASCII equivalents.

**Gate:** ✅ 121 tests pass; pushed to main; CI building.

### ✅ M19 — Windows real-machine testing + installer/app fixes (v0.21.19–v0.21.25)

Fresh Windows install of v0.21.18 revealed a series of bugs. All fixed and shipped.

**v0.21.19 — Write-Log crash on line 58**
- `[FATAL] The term 'Write-Log' is not recognized` — PS5.1 executes top-to-bottom; Write-Log was called at lines 55/58 (inside R-discovery try/catch) before its `function` block at line 87.
- Fix: moved Write-Log definition to before the first call.
- Regression test: `test_ps1_write_log_defined_before_first_use` in `test_installer_sanity.py`.

**v0.21.20 — Remove broken R auto-discovery**
- SYSTEM account has no outbound network → CRAN auto-discovery always failed with `[FATAL] Could not auto-discover R version`.
- Fix: removed try/catch block entirely; simplified to `$R_VERSION = "4.5.1"` (pinned).

**v0.21.21 — App shows v0.0.0**
- CI never wrote `app/_version.py` before PyInstaller. `_current_version()` falls through to `"0.0.0"`.
- Fix: added "Write _version.py" step in `ci.yml` build job before PyInstaller runs.
- Added `app/_version.py` to `.gitignore`.

**v0.21.22 — Reports saved to hidden AppData folder**
- Both `generate_single_report_worker` and `generate_reports_thread` hardcoded `REPORTS_DIR` (hidden `AppData\Roaming\ResilienceScan\reports\`), ignoring the output folder shown in the UI.
- Fix: both now use `out_dir = Path(self.output_folder_var.get())`.
- Added `[INFO] Output folder:` log at batch start.

**v0.21.23 — NoneType kill race + corrupt temp PDFs**
- `cancel_generation()` set `self._gen_proc = None` from main thread while generation thread called `.kill()` → `AttributeError: 'NoneType' object has no attribute 'kill'`.
- xelatex sometimes left a partial PDF (e.g. exit code 3221225786 = STATUS_CONTROL_C_EXIT) that appeared briefly then made quarto crash on rename.
- Fix: capture `self._gen_proc` in local var with try/except guard; unlink `temp_path` on non-zero returncode.

**v0.21.24 — Default output folder → Documents\ResilienceScanReports**
- `AppData\Roaming` is hidden by default; users couldn't find their reports.
- Added `_default_output_dir()` → `Documents\ResilienceScanReports\` (frozen Win/Linux) or `reports/` (dev).
- `self.output_folder_var` now defaults to `DEFAULT_OUTPUT_DIR`.
- Success log now shows full path: `[OK] Saved: C:\Users\...\file.pdf`.

**v0.21.25 — Email tab + stats use configured output folder**
- Email status display, prerequisite check, send loop, preview attachment lookup, and both stats counters all hardcoded `REPORTS_DIR` — showed 0 reports even after generation succeeded.
- Fix: all replaced with `Path(self.output_folder_var.get())`.
- Error messages now show the actual folder being searched.

**Gate:** ⏳ Re-test v0.21.25 on Windows — confirm reports appear in Documents, email tab finds them.

### ✅ M20 — Setup completion feedback (v0.21.26)

Background installer gave no feedback — users saw confusing "R NOT FOUND" errors if they opened the app during the 5-20 min setup window.

**Sentinel files:**
- Both setup scripts write `setup_running.flag` immediately on start.
- On exit (normal or crash), write `setup_complete.flag` with `PASS` or `FAIL` content, remove running flag.
- Windows: `C:\ProgramData\ResilienceScan\` — Linux: `/opt/ResilenceScanReportBuilder/`

**`gui_system_check.setup_status()`:**
- Reads the flags, returns `running` / `complete_pass` / `complete_fail` / `unknown`.

**`_startup_guard()` context-aware dialogs:**
- `running` → friendly "Setup In Progress" info dialog (not a scary warning).
- `complete_fail` → "Setup Failed" warning with log path.
- `unknown` / dev → existing generic "Missing Components" warning unchanged.

**In-app polling (`_poll_setup_completion`):**
- Fires every 30 s while `running`; updates status bar when flag changes.
- Status bar shows "Installing dependencies... (5-20 min)" while running, "Setup complete — all dependencies ready." on success, resets to "Ready" after 10 s.

**Desktop notifications (best-effort):**
- Windows: `msg *` from SYSTEM (works on Pro/Enterprise, silently skipped on Home).
- Linux: `notify-send` via `sudo -u $LOGGED_USER` with D-Bus session path.

**Gate:** ✅ 122 tests pass; ruff clean.

---

## Next milestones

### ✅ M21 — Fix email sending (v0.21.27)

Three bugs caused emails to silently fail and stay "pending":

**Bug 1 — Thread-unsafe Tkinter widget access (root cause of "nothing logged")**
- `log()`, `log_email()`, `log_gen()` called `widget.insert()` directly from the background send thread.
- Tkinter is single-threaded; this crashes the thread silently (especially on Linux).
- Fix: all three now check `threading.current_thread()` and schedule widget updates via `root.after(0, ...)` when called from a non-main thread.

**Bug 2 — Widget values read from background thread**
- `_send_emails_impl()` called `self.smtp_server_var.get()`, `self.email_body_text.get()`, etc. from the background thread — not thread-safe.
- Fix: `start_sending_emails()` captures a `send_config` dict from all widget vars on the main thread before launching the thread. `send_emails_thread(send_config)` and `_send_emails_impl(send_config)` use the dict.

**Bug 3 — No `except` in `send_emails_thread`**
- Any exception in `_send_emails_impl()` killed the thread silently: `is_sending_emails` stayed `True`, send button stayed disabled, no dialog shown.
- Fix: added `except Exception as exc:` that logs the full traceback to the email log, shows an error dialog, and resets the send button.

**Bug 4 — Status display ignored `email_tracker`**
- `update_email_status_display()` only checked the CSV `reportsent` column — never the tracker. Test-mode sends update the tracker but not the CSV, so status always showed "pending".
- Fix: display now checks `email_tracker._recipients` first; falls back to CSV only if no tracker entry. Stats label now also shows "Failed" count.

**Gate:** ✅ 122 tests pass; ruff clean.

---

## Next milestones

*Re-test email sending on Windows — confirm emails send, log shows progress, status updates to sent/failed.*

---

### ✅ M22 — R installer hardening + multi-format import (v0.21.28)

**R installer hardening (Windows + Linux):**
- Version comparison bug: unknown installed version was treated as "OK to skip" → now forces reinstall.
- Stale `r-library` cleanup: when R was upgraded, old binary packages would silently fail to load (ABI mismatch). Both setup scripts now wipe `r-library` when `R_UPGRADED=true`.
- R library writable check: added write-test before `install.packages()`; logs clear `ERROR` if not writable instead of silent partial install.
- Windows: `icacls` grants `SYSTEM:F`, `Administrators:F`, `Users:RX` on `r-library`.
- Linux: `chmod -R u+w` + `chown -R root:root` fallback; `chmod -R a+rX` after install.

**Multi-format data import:**
- `convert_data.py` extended to accept `.xlsx`, `.xls`, `.ods`, `.xml`, `.tsv`, `.csv` (raw survey exports).
- All formats normalise to `data/cleaned_master.csv` via the same column-mapping logic.
- `app/main.py` file picker updated to show all supported extensions.
- `requirements.txt`: added `odfpy>=1.4.1` for ODS support.
- `tests/test_convert_formats.py` (10 new tests): extension constant, dispatcher, XML strategies, CSV with header offset, `convert_and_save` with explicit path, `_find_source_file`.

**Gate:** ✅ 132 tests pass; ruff clean.

---

### ⏳ M23 — SCROL matrix report template (v0.22.0)

Second Quarto template based on the Supply Chain Resilience Opportunities & Limitations (SCROL) matrix from `Resilience - Dashboard V3.5.xlsm`.

**DO NOT modify `ResilienceReport.qmd`.**

**Template structure (`SCROLReport.qmd`):**
- Same params as `ResilienceReport.qmd` (`company`, `person`, `data_file`).
- SCROL matrix table: rows = Supply Chain (overall avg), Upstream, Downstream, Internal/Process, Internal/Product; columns = Redundancy (R), Collaboration (C), Flexibility (F), Visibility (V), Agility (A) + Row avg.
- Column mapping: `up__r/c/f/v/a`, `do__r/c/f/v/a`, `in__r/c/f/v/a` (all on 0-5 scale).
- Colour-coded cells: red (<2), orange (2-3), yellow (3-4), green (≥4).
- Spider/radar chart per dimension.
- Sector benchmarking where data is available in the CSV.

**App wiring:**
- Template dropdown in the app shows `ResilienceReport.qmd` and `SCROLReport.qmd`.
- `_sync_template()` copies both QMD files + shared assets on startup.
- Generate single / generate all use whichever template is selected.

**Gate:** App generates correct SCROL PDF for a sample row; original ResilienceReport.qmd unmodified.

---

### ⏳ M24 — Independent code analysis (REVIEW.md)

Thorough, independent review of the entire codebase documented in `REVIEW.md`.

**Scope:**
- Dead code: functions defined but never called; UI buttons wired to stubs.
- Unused imports and variables.
- Security issues: command injection, path traversal, hardcoded credentials.
- Thread-safety issues: any remaining direct widget access from non-main threads.
- Frozen-app path correctness: any residual `ROOT_DIR` / `sys._MEIPASS` / `_internal/` misuse.
- Data-flow gaps: CSV columns read but never validated; error paths that swallow exceptions.
- Code duplication: copy-pasted blocks that could share a helper.
- Test coverage gaps.

**Gate:** `REVIEW.md` committed; no new bugs introduced.

---

### ⏳ M25 — Dead-code cleanup

Remove all stubs, fake buttons, and unreachable code identified in M24.

**Rule:** Only delete code confirmed dead in M24 review — no refactoring, no feature changes.

**Gate:** All 132+ tests still pass; ruff clean; no functionality regression.

---

### ⏳ M26 — Refactor

Structured refactor of `app/main.py` (currently ~3750 lines, all in one class).

**Scope:**
- Extract logical sections into separate modules: `gui_generate.py`, `gui_email.py`, `gui_settings.py`.
- Replace copy-pasted path/env-building blocks with shared helpers.
- No behaviour changes — tests must pass unchanged.

**Gate:** All tests pass; ruff clean; app starts and all features work on both platforms.
