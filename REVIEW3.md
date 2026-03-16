# REVIEW3.md — Code Review & Refactoring Analysis (Round 3)

**Date:** 2026-03-16
**Version reviewed:** v0.21.47
**Scope:** All Python source files (`.qmd` templates excluded)
**Tests at review time:** 268 passing, ruff clean

---

## Executive summary

The codebase is in good shape overall. Architecture is sound, thread-safety is correctly implemented, frozen-app paths are correct, and all REVIEW.md / REVIEW2.md findings have been addressed. This round identifies **18 new findings** across six categories: dead code, encoding gaps, module size, hardcoded constants, broad exception handling, and minor design issues.

**Critical (fix before next release):** 2
**Medium (fix in next milestone):** 8
**Low (fix when convenient):** 8

---

## §1 Dead code

### 1.1 Useless if/else in `generate_single_report_worker()` — MEDIUM

**File:** `app/gui_generate.py:291–295`

```python
template_name = Path(self.template_var.get()).stem
if template_name.startswith("Report"):
    report_name = template_name   # ← both branches
else:
    report_name = template_name   # ← identical
```

The condition was likely intended to produce different names for ResilienceReport vs SCROLReport, but both branches assign the same value. The `if/else` is dead and the comment `# For standard reports, use the template name as is` is misleading.

**Same pattern appears at:** `app/gui_generate.py:654–659` (inside `generate_reports_thread()`).

**Fix:** Remove the conditional entirely:
```python
report_name = Path(self.template_var.get()).stem
```

---

### 1.2 `_config_path()` duplicates `_data_root()` logic — LOW

**File:** `app/app_paths.py:109–118`

`_config_path()` contains an inline copy of the frozen/dev/platform detection that `_data_root()` already encapsulates:

```python
def _config_path() -> Path:
    if getattr(sys, "frozen", False):
        if sys.platform == "win32":
            _base = Path(os.environ.get("APPDATA", ...)) / "ResilienceScan"
        else:
            _base = Path.home() / ".local" / "share" / "resiliencescan"
    else:
        _base = ROOT_DIR
    return _base / "config.yml"
```

**Fix:** Replace the body with `return _data_root() / "config.yml"`. One fewer place to update if the directory structure ever changes.

---

### 1.3 `pd.isna()` import for a one-liner in `filename_utils.py` — LOW

**File:** `utils/filename_utils.py:7`

```python
import pandas as pd   # only used for pd.isna() on line 13
```

Importing the entire pandas library to call `pd.isna(name)` once. `pd.isna` returns `True` for `None`, `float('nan')`, and `pd.NA`. For the string inputs this function receives, a plain `not name` or `name is None or str(name).strip() == ""` covers all realistic cases without the import.

**Fix:** Remove the pandas import and replace `pd.isna(name) or name == ""` with `not name or str(name).strip() == ""`.

---

## §2 Encoding gaps

### 2.1 `email_tracker.py:78` missing `encoding="utf-8"` — CRITICAL

**File:** `email_tracker.py:78`

```python
df = pd.read_csv(path, low_memory=False)   # ← no encoding
```

This is the one remaining `pd.read_csv()` call without `encoding="utf-8"`. On Windows with a non-UTF-8 ANSI code page, respondent names containing accented characters (é, ü, ñ) will be read as mojibake or raise `UnicodeDecodeError`, silently corrupting the send-tracking display.

All other `pd.read_csv()` calls in the codebase correctly specify `encoding="utf-8"`.

**Fix:** `df = pd.read_csv(path, low_memory=False, encoding="utf-8")`

---

### 2.2 `generate_all_reports.py` CSV reader only tries `utf-8` and `latin-1` — LOW

**File:** `generate_all_reports.py:37–44`

The `_load_csv()` helper tries `utf-8` then `latin-1`. Since `cleaned_master.csv` is always written with `encoding="utf-8"` by `convert_data.py`, the `latin-1` fallback is dead code in normal use. The real risk is that this standalone script is a dev-only CLI tool but is not documented as such in the file itself.

**Fix:** Add a comment at the top of the path-constants block:
```python
# NOTE: dev-only CLI tool — reads from repo/data/, not the frozen-app data dir.
```

---

## §3 Module size — candidates for splitting

### 3.1 `app/gui_email.py` is 1,480 lines — MEDIUM

The file contains three logically distinct concerns that have grown together:

| Concern | Approximate lines | Suggested module |
|---|---|---|
| Email template editor (UI widgets, preview, load/save) | ~300 | `app/gui_email_template.py` |
| Email sending (thread, SMTP, Outlook, per-row logic) | ~600 | `app/gui_email_send.py` (or keep in `gui_email.py`) |
| Email status display (tracker table, mark sent/pending) | ~200 | stay in `gui_email.py` |

No behaviour change is needed — just move method groups into separate mixin classes and add them to `ResilienceScanGUI`'s inheritance list in `main.py`.

**Gate:** All 268 tests pass after split; ruff clean.

---

### 3.2 `app/gui_data.py` is 1,365 lines — MEDIUM

Contains data loading, quality analysis, column selection, duplicate detection, integrity validation, and export — plus the Dashboard tab. The Dashboard quick-action buttons call methods from every other mixin; separating them is tricky. A lighter refactor is to extract:

| Concern | Approximate lines | Suggested module |
|---|---|---|
| Data quality analysis (quality dashboard, duplicate detection) | ~400 | `app/gui_quality.py` |
| Rest (data tab, load/save, preview) | ~965 | stays in `gui_data.py` |

---

## §4 Hardcoded constants that should be centralised

### 4.1 Quarto render timeout duplicated — MEDIUM

**Files:** `app/gui_generate.py:368`, `generate_all_reports.py:144`

Both hardcode `timeout=300` (5 minutes per report). If this needs changing it must be updated in two places.

**Fix:** Add to `utils/constants.py`:
```python
QUARTO_TIMEOUT_SECONDS = 300
```
Import in both files.

---

### 4.2 SMTP timeout duplicated — LOW

**Files:** `app/gui_email.py:1354`, `send_email.py:199`

Both hardcode `timeout=30`. Same issue as §4.1.

**Fix:** Add `SMTP_TIMEOUT_SECONDS = 30` to `utils/constants.py`.

---

### 4.3 R-check subprocess timeout duplicated — LOW

**Files:** `app/app_paths.py:164`, `gui_system_check.py:225`

Both hardcode `timeout=30` for Rscript subprocess calls.

**Fix:** Add `R_SUBPROCESS_TIMEOUT = 30` to `utils/constants.py`.

---

## §5 Exception handling

### 5.1 `except Exception: pass` in score validation — MEDIUM

**File:** `app/gui_generate.py:553`

```python
try:
    float_val = float(str(val).replace(",", "."))
    if 0 <= float_val <= 5:
        available_scores += 1
except Exception:
    pass
```

Only `float()` can raise here; the exception is always `ValueError` (or `TypeError` for exotic inputs). Catching broad `Exception` masks programming errors in the surrounding code.

**Fix:** `except (ValueError, TypeError): pass`

---

### 5.2 `except Exception: pass` in log file write — MEDIUM

**File:** `app/gui_logs.py:65, 128`

Silent failures on disk-full, permission errors, or read-only log file. The UI will show nothing and the user won't know logging broke.

**Fix:** At minimum, print to stderr:
```python
except OSError as e:
    print(f"[gui_logs] Could not write log: {e}", file=sys.stderr)
```

---

### 5.3 Twenty-plus `except Exception` clauses — LOW

The grep across all source files shows **53 instances** of `except Exception` (or bare `except`). The majority are in UI code where broad catch is appropriate (prevent crash), but several in pipeline scripts should be narrowed. The most impactful to fix are listed in §5.1 and §5.2. For the rest, a pragmatic standard is:

> In background threads: keep `except Exception` but always log the error.
> In pipeline scripts: narrow to the specific exception type that can realistically occur.

---

## §6 Minor design issues

### 6.1 `validate_record_for_report()` has an inline score-column list — LOW

**File:** `app/gui_generate.py:520–550`

```python
score_columns = [
    "up__r", "up__c", "up__f", "up__v", "up__a",
    "in__r", ...
    "do__r", ...
]
```

This is a duplicate of `utils/constants.SCORE_COLUMNS`. A bug fix to one won't propagate.

**Fix:** `from utils.constants import SCORE_COLUMNS` and replace the inline list.

---

### 6.2 `update_checker.py` imports `sys` twice — LOW

**File:** `update_checker.py:51, 104`

```python
import sys          # line 51
...
import sys as _sys2 # line 104
```

The second import is inside a function and aliases `sys` to `_sys2`. This is unnecessary — `sys` is already available from the module-level import.

**Fix:** Remove the local `import sys as _sys2` and replace `_sys2` with `sys`.

---

### 6.3 Template name `if/else` should use `ResilienceScanReport` literal — LOW

**File:** `app/gui_generate.py:290, 652` (after fixing §1.1)

After removing the dead `if/else`, `report_name = Path(self.template_var.get()).stem` will produce `"ResilienceReport"` or `"SCROLReport"` — which differ from the historical PDF filename format (`"ResilienceScanReport"`). If there is a naming convention mismatch, it should be resolved explicitly rather than relying on the template stem.

**Investigate:** Check whether existing PDF filenames use `ResilienceReport` or `ResilienceScanReport` and align the output filename accordingly.

---

### 6.4 `gui_system_check.py` has no type hints on public functions — LOW

**File:** `gui_system_check.py` throughout

Functions like `_find_rscript()`, `_find_quarto()`, `_find_tlmgr()`, `_check_r_packages()` have no return type annotations. Given that they return `str | None` or `bool`, adding annotations makes the contracts explicit and enables mypy checking.

---

### 6.5 `filename_utils.py` hardcodes `"Unknown"` fallback — LOW

**File:** `utils/filename_utils.py:13`

```python
if not name or str(name).strip() == "":
    return "Unknown"
```

`"Unknown"` will appear in PDF filenames if a name is missing. This should either be documented as intentional or moved to `constants.py` as `UNKNOWN_NAME_PLACEHOLDER = "Unknown"`.

---

## Finding index

| # | File | Severity | Title |
|---|---|---|---|
| 1.1 | `app/gui_generate.py:291–295, 654–659` | MEDIUM | Dead if/else — both branches identical |
| 1.2 | `app/app_paths.py:109–118` | LOW | `_config_path()` duplicates `_data_root()` logic |
| 1.3 | `utils/filename_utils.py:7` | LOW | pandas imported only for `pd.isna()` |
| 2.1 | `email_tracker.py:78` | **CRITICAL** | `pd.read_csv()` missing `encoding="utf-8"` |
| 2.2 | `generate_all_reports.py:1–15` | LOW | Dev-only CLI not documented in file |
| 3.1 | `app/gui_email.py` (1,480 lines) | MEDIUM | Module too large — split into template + send |
| 3.2 | `app/gui_data.py` (1,365 lines) | MEDIUM | Module too large — extract quality analysis |
| 4.1 | `app/gui_generate.py:368`, `generate_all_reports.py:144` | MEDIUM | Quarto timeout 300s duplicated |
| 4.2 | `app/gui_email.py:1354`, `send_email.py:199` | LOW | SMTP timeout 30s duplicated |
| 4.3 | `app/app_paths.py:164`, `gui_system_check.py:225` | LOW | R subprocess timeout 30s duplicated |
| 5.1 | `app/gui_generate.py:553` | MEDIUM | `except Exception` should be `ValueError, TypeError` |
| 5.2 | `app/gui_logs.py:65, 128` | MEDIUM | Silent `except Exception: pass` on log write |
| 5.3 | Codebase-wide | LOW | 53 broad `except Exception` clauses |
| 6.1 | `app/gui_generate.py:520–550` | LOW | Inline score-column list duplicates `constants.SCORE_COLUMNS` |
| 6.2 | `update_checker.py:51, 104` | LOW | `sys` imported twice |
| 6.3 | `app/gui_generate.py:290, 652` | LOW | PDF filename stem may not match historical convention |
| 6.4 | `gui_system_check.py` | LOW | No type hints on public functions |
| 6.5 | `utils/filename_utils.py:13` | LOW | `"Unknown"` hardcoded fallback |

---

## Recommended milestone plan

### M48 — Quick wins (single session)
- Fix **2.1** (`email_tracker.py` encoding) — critical, 1 line
- Fix **1.1** (dead if/else in `gui_generate.py`) — 4 lines, both occurrences
- Fix **1.2** (`_config_path()` simplification)
- Fix **1.3** (remove pandas from `filename_utils.py`)
- Fix **5.1** (`except Exception` → `except (ValueError, TypeError)`)
- Fix **6.1** (inline score list → `constants.SCORE_COLUMNS`)
- Fix **6.2** (`sys` double-import in `update_checker.py`)
- Fix **4.1, 4.2, 4.3** (centralise timeout constants)

### M49 — Module splitting
- Split `app/gui_email.py` into template + send mixins (finding **3.1**)
- Optionally extract quality analysis from `app/gui_data.py` (finding **3.2**)

### M50 — Exception hardening + type hints
- Fix **5.2** (silent log failures)
- Narrow high-value `except Exception` clauses in pipeline scripts (finding **5.3**)
- Add type hints to `gui_system_check.py` (finding **6.4**)
- Add dev-only comment to `generate_all_reports.py` (finding **2.2**)
