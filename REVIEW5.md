# REVIEW5.md — Round-5 Code Review
Date: 2026-03-19
Reviewer: Claude Code (thorough read of all 23 source files + tests)
Baseline: v0.21.57, 268 tests, ruff clean

---

## Summary

8 findings. 1 critical, 2 high, 3 medium, 2 low.
No security regressions since REVIEW4. Thread safety and resource management from M52–M53 hold.
The critical finding (progress bar never advances) is a functional UI regression.

---

## Findings

### F001 — CRITICAL | Progress bar never updates on successful generation
**File:** `app/gui_generate.py`, lines 586–799
**Severity:** Critical — the progress bar stays at 0% throughout all successful report generations

The progress bar is initialised to `maximum=total, value=0` at line 587.
Inside the `for` loop, the happy-path increments `success` or `skipped` and logs, but never schedules a `gen_progress.configure(value=…)` call.
The `gen_progress_label` ("Progress: i/t | …") is also only updated inside the `except Exception` block at lines 791–799, meaning both the bar and the count label are frozen at 0 unless an exception fires.

**Fix:** Move the `root.after(…, gen_progress.configure(value=i))` and `gen_progress_label` update to the end of the try block (after each record is processed, whether success or skip), not only on exception.

---

### F002 — HIGH | `askyesnocancel` result checked with `if not response:` (single-file generate)
**File:** `app/gui_generate.py`, line 306
**Severity:** High — the truthiness check is semantically fragile and misleads future maintainers

```python
response = messagebox.askyesnocancel("File Exists", "…Overwrite?")
if not response:          # True when response is False (No) OR None (Cancel)
    self.log_gen("…cancelled…")
    return
```

`askyesnocancel` returns `True` (Yes), `False` (No), or `None` (Cancel/close).
The current code cancels on both No and Cancel, which is the correct *behaviour*, but the check is non-obvious. Any future change that adds logic between the three states (e.g. "skip this one but continue batch") will silently break.

**Fix:** `if response is not True:`

---

### F003 — HIGH | `_current_version()` imported and called twice in `main.py`
**File:** `app/main.py`, lines 46–48 and 147–151
**Severity:** High — redundant I/O and import at startup; also the first call's result is dead code

`__init__` imports `_current_version`, calls it, stores the result in local `_APP_VERSION`, then uses it only for `root.title(…)`. The result is never stored on `self`.
`create_header()` then re-imports and re-calls `_current_version()` for the subtitle label.

Two file reads at startup where one would do. `_APP_VERSION` in `__init__` is unreachable after startup.

**Fix:** Store the version once as `self._app_version = _current_version()` in `__init__`, remove the second import+call in `create_header()`, use `self._app_version` throughout.

---

### F004 — MEDIUM | `_find_row()` returning `None` is silent — no log warning
**File:** `app/gui_email_send.py`, lines 490–500
**Severity:** Medium — data inconsistency (PDF on disk but not in CSV) is silently swallowed

When `_find_row()` returns `None`, the caller substitutes `""` for the email address and `False` for `reportsent`. The PDF still appears in the send list but with a blank email field. No message is logged, so there is no visible indication that the CSV and the reports folder are out of sync.

**Fix:** When `row is None`, log a warning: `self.log_email(f"[WARN] No CSV record found for {company} – {person}; email will be blank")`.

---

### F005 — MEDIUM | Test-email validation accepts structurally invalid addresses
**File:** `app/gui_email_send.py`, line 319
**Severity:** Medium — users can proceed with malformed test addresses that will fail at send time

```python
if not test_email or "@" not in test_email:
    # show error
```

This accepts `@`, `foo@`, `@bar`, and `user name@domain` as valid. The error surface is small (test mode only), but the failure mode is confusing — the send thread starts, the first SMTP `RCPT TO` is rejected, and the error message comes from the SMTP server rather than the validation layer.

**Fix:** Replace the single `"@" in` check with a minimal structural check:

```python
parts = test_email.split("@")
if len(parts) != 2 or not parts[0] or not parts[1] or "." not in parts[1]:
    # show error
```

---

### F006 — MEDIUM | `_match_key()` called twice on identical DataFrames in `convert_data.py`
**File:** `convert_data.py`, lines 377–389
**Severity:** Medium — redundant string operations on potentially large DataFrames

```python
def _match_key(df):  # performs astype(str) + str.lower() + str.strip()
    …

new_keys = _match_key(new_df)   # line 388
old_keys = _match_key(old_df)   # line 389
```

Each call runs three pandas string-method chains. For large files (thousands of rows) this doubles the work. The function is defined as a closure inside `_upsert_records` and is never reused anywhere else.

**Fix:** Inline the key computation once each, or rename to make clear they are different DataFrames. No algorithmic change needed — just avoid calling on the same DataFrame twice. (They *are* called on different DataFrames — new_df and old_df — so the calls are correct. The real concern is ensuring the closure is not accidentally called again; no code change strictly needed, but making `_match_key` a module-level function with a clear name like `_build_match_key(df)` prevents accidental re-call.)

**Revised severity: Low** — the calls are on different DataFrames (new_df ≠ old_df), so there is no actual redundancy. Retaining as a Low finding for clarity/naming only.

---

### F007 — LOW | `validate_columns` builds a list for O(n) membership checks
**File:** `clean_data.py`, lines 91–103
**Severity:** Low — performance, not correctness

```python
df_cols_lower = [col.lower() for col in df.columns]   # list
for req_col in REQUIRED_COLUMNS:
    if req_col.lower() not in df_cols_lower:           # O(n) per check
```

The same list is iterated again at line 102 for SCORE_COLUMNS. Converting `df_cols_lower` to a `set` makes each `in` check O(1). For typical files (tens to hundreds of columns) this is unnoticeable, but it is straightforwardly better.

**Fix:** `df_cols_lower = {col.lower() for col in df.columns}`

---

### F008 — LOW | Progress counter label font inconsistency
**File:** `app/gui_email_send.py`, line 51 (email stats label)
**File:** `app/main.py`, lines 143, 172 (header uses Segoe UI)
**Severity:** Low — cosmetic; Arial vs Segoe UI after the M56 theme upgrade

After M56 standardised the header on `"Segoe UI"`, the email status label at line 51 still uses `font=("Arial", 10, "bold")`. On systems where Arial differs from Segoe UI this produces a visual inconsistency.

**Fix:** Change to `font=("Segoe UI", 10, "bold")` — matches the rest of the interface.

---

## Milestone plan

| Milestone | Finding(s) | Description |
|-----------|-----------|-------------|
| M58 | F001 | Fix progress bar — update bar and label on every iteration (success, skip, fail) |
| M59 | F002, F003 | Fix `askyesnocancel` check + deduplicate `_current_version()` |
| M60 | F004, F005 | Log warning for missing CSV record; stronger email address validation |
| M61 | F007, F008 | Set for column lookup; font consistency |

F006 is reclassified Low and folded into M61 as a naming-only clarification (no functional change required).

---

## Not found / confirmed clean

- Thread safety: all shared state still properly guarded (M52 locks hold)
- Keyring integration: no password leaks to config.yml (M53 holds)
- Path resolution: frozen/dev split correct throughout (REVIEW3 fix holds)
- SMTP socket: quit/close fallback in place (M52 fix holds)
- ExcelFile context managers: all `with pd.ExcelFile(…) as xl:` (M52 fix holds)
- `_cell_text()` returns `""` not `None` (M54 fix holds)
- `TEST_MODE_LABEL` constant used consistently (M54 fix holds)
- No plaintext passwords anywhere in source
- No bare `except:` without specific re-raise or logged message
- CI / NSIS / postinst: no new issues
