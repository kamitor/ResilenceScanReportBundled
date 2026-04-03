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

`app/main.py` is the entry point (Tkinter GUI + PyInstaller target). It inherits from five mixins:

| Mixin | File | Responsibility |
|---|---|---|
| `DataMixin(QualityMixin)` | `app/gui_data.py` | Data tab, CSV load/convert, quality analysis display |
| `GenerationMixin` | `app/gui_generate.py` | Generation tab, PDF render, thread-safe cancel |
| `EmailMixin(EmailTemplateMixin, EmailSendMixin)` | `app/gui_email.py` | Email tabs, tracker display, mark-sent/pending |
| `SettingsMixin` | `app/gui_settings.py` | Startup guard, system check, installer |
| `LogsMixin` | `app/gui_logs.py` | Logs tab, log/log_gen/log_email helpers |

Sub-mixins:

| File | Responsibility |
|---|---|
| `app/gui_quality.py` | `QualityMixin`: `analyze_data_quality()` — inline quality metrics |
| `app/gui_email_template.py` | `EmailTemplateMixin`: template editor, SMTP config, sender profiles, load/save/preview |
| `app/gui_email_send.py` | `EmailSendMixin`: sending tab, SMTP thread, Outlook COM, per-row send |

Shared infrastructure:

| Module | Purpose |
|---|---|
| `app/app_paths.py` | All path constants + `_sync_template()`, `_check_r_packages_ready()` |
| `utils/path_utils.py` | `get_user_base_dir()` for pipeline scripts |
| `utils/filename_utils.py` | `safe_filename()`, `safe_display_name()` |
| `utils/constants.py` | `SCORE_COLUMNS`, `REQUIRED_COLUMNS`, timeout constants, `TEST_MODE_LABEL` |
| `gui_system_check.py` | R / Quarto / TinyTeX runtime checks |
| `email_tracker.py` | Per-recipient send-status JSON store (thread-safe, `threading.Lock`) |
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

**Bundled installer** — R, Quarto, TinyTeX, and all R packages are downloaded and pre-installed during the **CI build**, then shipped inside the release artifact alongside the Python app. No network access or post-install scripts are needed on the end-user machine.

- Python is bundled by PyInstaller (`--onedir`) as before.
- R portable, Quarto portable, TinyTeX, and the R package library are baked into the build artifact at CI time and placed under `_internal/` (frozen) / `vendor/` (dev convention).
- The staged installer approach (NSIS post-install scripts, `setup_linux.sh`, `setup_macos.sh`, `setup_dependencies.ps1`, `postinst.sh`) is **replaced** by pre-bundled deps. These scripts are retained temporarily for reference but will be removed once the bundled build is verified.
- End-user installer is larger (~400–600 MB) but truly self-contained and offline-capable.

`ResilienceReport.qmd` and `SCROLReport.qmd` are deeply LaTeX-dependent (TikZ, kableExtra, custom titlepage extension, custom fonts, raw `.tex` includes). **The PDF engine cannot be switched to Typst or WeasyPrint** — TinyTeX is required.

**Do not modify `.qmd` templates** — they contain interdependent LaTeX/R/Quarto logic that is fragile to whitespace and encoding changes.

### Pinned dependency versions

| Dependency | Version | Notes |
|---|---|---|
| R | 4.5.1 | Pinned — downloaded in CI, bundled into artifact |
| Quarto | 1.6.39 | GitHub releases — bundled into artifact |
| TinyTeX | Quarto-pinned | `quarto install tinytex` — bundled into artifact |
| Python | ≥ 3.11 | Bundled by PyInstaller |

### Bundle path resolution (frozen app)

| Component | Frozen path | Dev path |
|---|---|---|
| Rscript | `{_MEIPASS}/r/bin/Rscript` (Win: `.../R/bin/Rscript.exe`) | system PATH or `vendor/r/bin/Rscript` |
| quarto | `{_MEIPASS}/quarto/bin/quarto` | system PATH or `vendor/quarto/bin/quarto` |
| tlmgr / pdflatex | `{_MEIPASS}/tinytex/bin/<arch>/tlmgr` | system PATH or `vendor/tinytex/...` |
| R library | `{_MEIPASS}/r-library/` | `vendor/r-library/` |

`app/app_paths.py` must expose `R_BIN`, `QUARTO_BIN`, `TINYTEX_BIN` constants that resolve bundle paths first, falling back to system PATH for dev convenience.

### R packages

`readr`, `dplyr`, `stringr`, `tidyr`, `ggplot2`, `knitr`, `fmsb`, `scales`, `viridis`, `patchwork`, `RColorBrewer`, `gridExtra`, `png`, `lubridate`, `kableExtra`, `rmarkdown`, `jsonlite`, `ggrepel`, `cowplot`

### LaTeX packages (tlmgr)

`pgf`, `xcolor`, `colortbl`, `booktabs`, `multirow`, `float`, `wrapfig`, `pdflscape`, `geometry`, `preprint`, `graphics`, `tabu`, `threeparttable`, `threeparttablex`, `ulem`, `makecell`, `environ`, `trimspaces`, `caption`, `hyperref`, `setspace`, `fancyhdr`, `microtype`, `lm`, `needspace`, `varwidth`, `mdwtools`, `xstring`, `tools`

**Note:** `capt-of` stub is written at build time and bundled — not installed at end-user machine.

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
| M50 | Module splitting: gui_email.py → template+send+tracker; gui_data.py → QualityMixin extracted |
| M51 | Exception narrowing in pipeline scripts; type hints on gui_system_check.py functions |
| M52 | Thread safety + resource leaks: EmailTracker lock, log file lock, SMTP socket, ExcelFile context manager |
| M53 | Security: SMTP password moved from plaintext config.yml to OS keyring |
| M54 | Code quality: _find_row() helper, _cell_text() returns "", TEST_MODE_LABEL constant |
| M55 | GUI audit: removed dead "Run Quality Dashboard" / "Run Data Cleaner" buttons (scripts missing) |
| M56 | GUI visual upgrade: sv-ttk Sun Valley theme, card-style stats, Segoe UI fonts |
| M57 | Email sender profiles: named profiles in config.yml, profile selector in SMTP tab, "Sending from:" in Send tab |
| M58–M61 | Round-5 review fixes (REVIEW5.md): progress bar, askyesnocancel check, _current_version dedup, CSV-miss warning, email validation, set lookup, font consistency |
| M62 | macOS (Apple Silicon) build target: darwin paths, setup_macos.sh, DMG CI, 134 new tests (268 → 402 total) |
| M63 | **Architecture pivot: bundled installer** — pre-bake R + Quarto + TinyTeX + R packages into CI artifact; `utils/bin_paths.py` for bundle-first resolution; remove staged post-install scripts; add `R_BIN`/`QUARTO_BIN`/`TINYTEX_BIN` to `app_paths.py`; 420 tests |

**Current version: v0.21.63 — 420 tests, ruff clean**

---

## Active milestones

All milestones M1–M63 complete.

---

## Proposed next milestones

### M64 — Verify CI in new repo + fix repo-wide stale references
**Goal:** Confirm the M63 bundled-build CI passes end-to-end in `kamitor/ResilenceScanReportBundled`; remove every remaining reference to the old repo.

**Tasks:**
1. Check GitHub Actions run for sha `a6a3618` — confirm `test`, `version-check`, `build` (Windows/Linux/macOS), `installer-smoke`, `build-smoke-macos`, and `publish` all pass.
2. Grep codebase for `Windesheim-A-I-Support`, `ResilenceScanReportBuilder` (old name), and `Windesheim` — fix any remaining occurrences in source files, README, packaging metadata, and desktop entries.
3. Update `README.md` entirely: remove all staged-installer documentation (PowerShell task scheduler, postinst, setup_macos.sh instructions), replace with bundled-installer install steps (just run the installer — no setup wait), update download links to new repo.
4. Update `nfpm.yaml` maintainer/homepage fields if they reference the old org.

**Gate:** CI green on new repo; `grep -r "Windesheim-A-I-Support\|ResilenceScanReportBuilder[^B]" .` returns nothing outside `CLAUDE.md` history entries.

---

### M65 — End-to-end render smoke test in CI
**Goal:** After bundling, actually invoke `quarto render` on a minimal fixture inside CI to confirm the bundled R + Quarto + TinyTeX stack produces a valid PDF — not just that the binaries exist.

**Tasks:**
1. Create `tests/fixtures/smoke_report.qmd` — a minimal Quarto document that loads one R package (e.g. `ggplot2`) and renders to PDF.
2. Add a CI step (in the `build` job, after vendoring completes) that runs:
   ```bash
   vendor/quarto/bin/quarto render tests/fixtures/smoke_report.qmd --to pdf
   ```
   using `build_r_env()` for the subprocess environment.
3. Assert the output PDF exists and is > 1 KB.
4. Run this on all three platforms (Windows, Linux, macOS) as part of the existing build matrix.

**Gate:** CI produces a rendered PDF on all three platforms; job fails if `quarto render` exits non-zero.

---

### M66 — Bundle size audit + artifact optimisation
**Goal:** Measure and reduce the size of the bundled release artifacts (expected ~400–600 MB) to the minimum necessary for correct operation.

**Tasks:**
1. After a successful M64 CI run, record artifact sizes for each platform.
2. Audit `vendor/r/` for unneeded components: source files, static libraries, docs, demos, Tcl/Tk if unused, 32-bit libs on 64-bit Windows.
3. Audit `vendor/r-library/` — strip compiled `.o`/`.a` files from R package builds (not needed at runtime).
4. Audit `vendor/tinytex/` — identify and remove LaTeX packages not in the required `tlmgr` list.
5. Re-measure artifact sizes; target < 350 MB per platform.
6. Document final sizes in CLAUDE.md.

**Gate:** All three platform artifacts < 350 MB; `quarto render` smoke test (M65) still passes after stripping.

---

### M67 — Round-6 code review (post-M63 surface)
**Goal:** Fresh independent review of the code changes introduced in M63 (bin_paths, gui_settings refactor, ci.yml rewrite, app_paths changes) to catch regressions, edge cases, or design issues.

**Tasks:**
1. Write `REVIEW6.md` covering: `utils/bin_paths.py`, `app/app_paths.py`, `app/gui_settings.py`, `app/gui_generate.py`, `generate_all_reports.py`, `.github/workflows/ci.yml`.
2. For each finding: severity (critical / warning / info), description, proposed fix.
3. Implement all critical and warning findings; document info findings for later.
4. Update test suite if new edge cases are identified.

**Gate:** `REVIEW6.md` complete; all critical/warning findings resolved; 420+ tests passing.

---

### M68 — GitHub Pages website polish
**Goal:** Make the `docs/index.html` promotional site production-quality with real screenshots, accurate download links (auto-updated by CI), and a link from the repo README.

**Tasks:**
1. Add a screenshot of the GUI (Data tab, Generation tab) to `docs/assets/`.
2. Add a sample PDF thumbnail to `docs/assets/` showing the report format.
3. Wire CI `update-readme` job to also patch the version number in `docs/index.html` after each release (replace hardcoded `v0.21.63` with the live version).
4. Add a "Website" link to `README.md` pointing at `https://kamitor.github.io/ResilenceScanReportBundled/`.
5. Add Open Graph meta tags (`og:title`, `og:description`, `og:image`) to `docs/index.html`.

**Gate:** Site live at GitHub Pages URL; version number auto-updates on release; screenshots present.

---

### M69 — macOS code signing (Gatekeeper hardening)
**Goal:** Eliminate the "app is damaged" Gatekeeper error on macOS by signing and notarising the DMG.

**Tasks:**
1. Obtain an Apple Developer ID Application certificate (requires paid Apple Developer account).
2. Add GitHub Actions secrets: `APPLE_CERT_BASE64`, `APPLE_CERT_PASSWORD`, `APPLE_TEAM_ID`, `APPLE_NOTARIZE_USER`, `APPLE_NOTARIZE_PASSWORD`.
3. Add codesign step in the macOS build job: sign the `.app` bundle and all bundled binaries (Rscript, quarto, tlmgr).
4. Add `xcrun notarytool submit` + `xcrun stapler staple` steps for the DMG.
5. Update README macOS section — remove the `xattr -cr` workaround once signing is in place.

**Gate:** DMG installs and opens on a clean macOS machine without any Gatekeeper dialog; notarisation ticket stapled to DMG.

---

### M70 — Windows installer code signing
**Goal:** Sign the `.exe` installer with an EV code-signing certificate to prevent SmartScreen warnings.

**Tasks:**
1. Obtain an EV code-signing certificate (DigiCert / Sectigo or equivalent).
2. Add GitHub Actions secret: `WINDOWS_CERT_BASE64`, `WINDOWS_CERT_PASSWORD`.
3. Add `signtool sign` step in the Windows build job after NSIS produces the `.exe`.
4. Update README Windows section to note the signed installer.

**Gate:** Installer passes Windows SmartScreen without warning on a clean Windows machine.
