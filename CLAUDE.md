# CLAUDE.md

Guidance for Claude Code when working in this repository.

## Commands

```bash
# Set up Python environment
python -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt

# Run the GUI
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

# Regenerate anonymised test fixture
python scripts/make_sample_data.py
```

## Release workflow

Bump `version` in `pyproject.toml` and push to `main`. CI detects no git tag `v<version>` and fires the build matrix. **Do not create tags manually.** macOS is not a target — only Windows and Linux.

---

## Architecture

`app/main.py` is the entry point (Tkinter GUI + PyInstaller target, 255 lines). It inherits from five mixins:

| Mixin | File | Responsibility |
|---|---|---|
| `DataMixin` | `app/gui_data.py` | Data tab, CSV load/convert, quality analysis |
| `GenerationMixin` | `app/gui_generate.py` | Generation tab, PDF render, thread-safe cancel |
| `EmailMixin` | `app/gui_email.py` | Email tabs, SMTP send, template editor |
| `SettingsMixin` | `app/gui_settings.py` | Startup guard, system check, installer |
| `LogsMixin` | `app/gui_logs.py` | Logs tab, log/log_gen/log_email helpers |

Shared infrastructure:

| Module | Purpose |
|---|---|
| `app/app_paths.py` | All path constants + `_sync_template()`, `_check_r_packages_ready()` |
| `utils/path_utils.py` | `get_user_base_dir()` for pipeline scripts |
| `utils/filename_utils.py` | `safe_filename()`, `safe_display_name()` |
| `utils/constants.py` | `SCORE_COLUMNS`, `REQUIRED_COLUMNS` |
| `gui_system_check.py` | R / Quarto / TinyTeX runtime checks |
| `email_tracker.py` | Per-recipient send-status JSON store |
| `update_checker.py` | Background GitHub release check |

### Path resolution (frozen vs dev)

| Constant | Dev | Frozen (installed) |
|---|---|---|
| `ROOT_DIR` / `_asset_root()` | repo root | `sys._MEIPASS` (`_internal/`) — **read-only** |
| `_DATA_ROOT` / `_data_root()` | repo root | `%APPDATA%\ResilienceScan` (Win) / `~/.local/share/resiliencescan` (Linux) |
| `DATA_FILE` | `repo/data/cleaned_master.csv` | `APPDATA/data/cleaned_master.csv` |
| `REPORTS_DIR` | `repo/reports/` | `APPDATA/reports/` — temp write location only |
| `DEFAULT_OUTPUT_DIR` | `repo/reports/` | `Documents\ResilienceScanReports\` |

**Rule:** Code that reads or displays reports must use `Path(self.output_folder_var.get())`, never `REPORTS_DIR`.

`_sync_template()` runs at import time and copies QMDs + assets from `_asset_root()` to `_data_root()` so Quarto can write `.quarto/` next to them (frozen `_internal/` is read-only).

---

## Pipeline flow

```
data/*.xlsx / .xlsm / .xls (incl. SpreadsheetML) / .ods / .xml / .json / .jsonl / .csv / .tsv
     │ convert_data.py  — reads → normalises columns → upserts into cleaned_master.csv
     ▼                    (new records first; reportsent preserved for existing rows)
data/cleaned_master.csv
     │ clean_data.py
     ▼
data/cleaned_master.csv  [validated & cleaned]
     │ generate_all_reports.py + ResilienceReport.qmd or SCROLReport.qmd
     ▼
reports/YYYYMMDD <TemplateName> (Company Name - Firstname Lastname).pdf
     │ validate_reports.py
     │ send_email.py
     ▼
emails via Outlook COM (Windows) or SMTP fallback (Office365)
```

**Key data file:** `data/cleaned_master.csv`
**Score columns:** `up__r/c/f/v/a`, `in__r/c/f/v/a`, `do__r/c/f/v/a` — range 0–5
**PDF naming:** `YYYYMMDD <TemplateName> (Company Name - Firstname Lastname).pdf`

---

## Packaging strategy

**Staged installer** — NSIS (Windows) / postinst (Linux) silently downloads and installs R, Quarto, TinyTeX, and R packages during installation. Python is bundled by PyInstaller (`--onedir`).

`ResilienceReport.qmd` and `SCROLReport.qmd` are deeply LaTeX-dependent (TikZ, kableExtra, custom titlepage extension, custom fonts, raw `.tex` includes). **The PDF engine cannot be switched to Typst or WeasyPrint** — TinyTeX is required.

**Do not modify `.qmd` templates** — they contain interdependent LaTeX/R/Quarto logic that is fragile to whitespace and encoding changes.

### Pinned dependency versions

| Dependency | Version | Notes |
|---|---|---|
| R | 4.5.1 | Pinned — SYSTEM account has no network at install time |
| Quarto | 1.6.39 | GitHub releases |
| TinyTeX | Quarto-pinned | `quarto install tinytex` |
| Python | ≥ 3.11 | Bundled by PyInstaller |

### R packages

`readr`, `dplyr`, `stringr`, `tidyr`, `ggplot2`, `knitr`, `fmsb`, `scales`, `viridis`, `patchwork`, `RColorBrewer`, `gridExtra`, `png`, `lubridate`, `kableExtra`, `rmarkdown`, `jsonlite`, `ggrepel`, `cowplot`

### LaTeX packages (tlmgr)

`pgf`, `xcolor`, `colortbl`, `booktabs`, `multirow`, `float`, `wrapfig`, `pdflscape`, `geometry`, `preprint`, `graphics`, `tabu`, `threeparttable`, `threeparttablex`, `ulem`, `makecell`, `environ`, `trimspaces`, `caption`, `hyperref`, `setspace`, `fancyhdr`, `microtype`, `lm`, `needspace`, `varwidth`, `mdwtools`, `xstring`, `tools`

**Note:** `capt-of` is NOT installed via tlmgr — a minimal stub is written by the installer scripts directly and registered with `mktexlsr`.

---

## Working rule

**Do not start the next milestone until the current one is fully verified by its gate condition.**

---

## Milestone history (summary)

| Range | What was built |
|---|---|
| M1–M7 | Core app: CI, paths, data conversion, report generation, email, startup guard |
| M8–M9 | Windows installer (R, Quarto, TinyTeX, LaTeX packages, capt-of) |
| M10–M14 | Frozen-app path split, smoke tests, e2e CI, update checker, README badges |
| M15–M19 | Frozen-app render fixes, cross-platform CI, Windows real-machine testing |
| M20–M23 | Setup completion feedback, email fixes, R hardening, SCROL template |
| M24 | Independent code review → REVIEW.md (27 findings) |
| M25–M29 | Thread-safety, frozen paths, dead-code removal, error handling, security |
| M30–M35 | Shared utils extraction, test coverage (69 tests), main.py refactor to 5 mixins, encoding safety, output folder validation, log standardisation |
| M36–M41 | Round-2 review fixes: path correctness, thread-safety, error handling, dead code, test coverage, security |
| M42–M45 | Installer hardening: stale-flag cleanup, smoke test CI, R repair in-app |
| M46 | SpreadsheetML XLS support + upsert (new records first, reportsent preserved) |
| M47 | JSON / JSONL / XLSM format support; dummy fixtures in tests/fixtures/ |
| M48 | Non-GUI quick wins: encoding fix, dead code, constants, pandas removal from filename_utils |
| M49 | GUI improvements: dead if/else, score constant, SMTP/Quarto timeout constants, silent log fix |

**Current version: v0.21.49 — 268 tests, ruff clean**

---

## Active milestones

---

### M49 — GUI improvements

Gate: all tests pass; ruff clean; version → v0.21.49.

Each task is independent and can be done in any order.

**T1 — Fix dead if/else in report name logic** (`app/gui_generate.py:291–295, 654–659`)
Both branches of the `if template_name.startswith("Report"):` block assign the same value. Remove the conditional:
```python
# Before (both branches identical — dead code)
if template_name.startswith("Report"):
    report_name = template_name
else:
    report_name = template_name

# After
report_name = Path(self.template_var.get()).stem
```
Apply the fix in both occurrences (single-report worker at line 291 and batch thread at line 654).

**T2 — Narrow score-validation exception** (`app/gui_generate.py:553`)
`except Exception: pass` in the score float-conversion loop catches everything. Only `float()` can raise here:
```python
# Before
except Exception:
    pass

# After
except (ValueError, TypeError):
    pass
```

**T3 — Replace inline score-column list with shared constant** (`app/gui_generate.py:520–550`)
`validate_record_for_report()` hardcodes the full score-column list, duplicating `utils/constants.SCORE_COLUMNS`. Replace with:
```python
from utils.constants import SCORE_COLUMNS
# ... then use SCORE_COLUMNS directly in the validation loop
```

**T4 — Add SMTP timeout constant** (`app/gui_email.py:1354`, `send_email.py:199`)
Both files hardcode `timeout=30` for SMTP connections. Centralise:
```python
# utils/constants.py
SMTP_TIMEOUT_SECONDS = 30
```
Import and use in both files.

**T5 — Fix silent log-write failures** (`app/gui_logs.py:65, 128`)
`except Exception: pass` swallows disk-full and permission errors silently. Replace with:
```python
except OSError as e:
    import sys
    print(f"[gui_logs] Log write failed: {e}", file=sys.stderr)
```

**T6 — Add Quarto render timeout constant to GUI** (`app/gui_generate.py:368`)
`timeout=300` in the GUI render call is a duplicate of the value in `generate_all_reports.py`. After M48 adds `QUARTO_TIMEOUT_SECONDS` to `constants.py`, import and use it here too.

**T7 — Verify PDF output filename stem** (`app/gui_generate.py:290, 652`)
After T1 removes the dead if/else, `report_name` will equal `Path(self.template_var.get()).stem`, producing e.g. `"ResilienceReport"`. Confirm this matches the filename convention users expect (`YYYYMMDD ResilienceReport (Company - Person).pdf`). If historical PDFs used a different stem, add a mapping dict in `constants.py` and apply it here.

---

### M50 — Module splitting

Gate: all tests pass; ruff clean; version → v0.21.50.

**T1 — Split `app/gui_email.py` (1,480 lines)**

Extract into three files:

| New file | New mixin | Contents |
|---|---|---|
| `app/gui_email_template.py` | `EmailTemplateMixin` | Template editor tab, load/save/reset/preview, placeholder logic |
| `app/gui_email_send.py` | `EmailSendMixin` | Sending tab, SMTP thread, Outlook COM, per-row send loop, stop button |
| `app/gui_email.py` | `EmailMixin` | Inherits both above + tracker display, mark-sent/pending, status table |

Update `app/main.py`: add `EmailTemplateMixin`, `EmailSendMixin` to `ResilienceScanGUI` inheritance list.

**T2 — Extract `app/gui_quality.py` from `gui_data.py` (1,365 lines)**

Move quality-analysis methods (~400 lines) into a new `QualityMixin`:
- `run_quality_dashboard()`, `analyze_data_quality()`, `run_data_cleaner()`
- `analyze_duplicates()`, `show_row_details()`, `export_filtered_data()`

`DataMixin` inherits `QualityMixin`; `ResilienceScanGUI` inheritance list unchanged.

---

### M51 — Exception hardening + type hints

Gate: all tests pass; ruff clean; version → v0.21.51.

- **5.3** Narrow high-value `except Exception` in pipeline scripts (`clean_data.py`, `generate_all_reports.py`, `send_email.py`) to specific exception types (`OSError`, `ValueError`, `subprocess.CalledProcessError`)
- **6.4** Add return type hints (`-> str | None`, `-> bool`, `-> Path | None`) to all public functions in `gui_system_check.py`
