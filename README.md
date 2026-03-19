# ResilienceScan Report Builder

[![Latest Release](https://img.shields.io/github/v/release/Windesheim-A-I-Support/ResilenceScanReportBuilder?label=latest)](https://github.com/Windesheim-A-I-Support/ResilenceScanReportBuilder/releases/latest)

A Windows/Linux desktop application that generates personalised PDF resilience reports for survey respondents and distributes them by email.  Built with Python (Tkinter GUI) + R + Quarto + TinyTeX.

---

## Downloads

<!-- DOWNLOAD_LINKS_START -->
| Platform | Download |
|----------|----------|
| Windows | [Windows Installer (.exe)](https://github.com/Windesheim-A-I-Support/ResilenceScanReportBuilder/releases/download/v0.21.52/ResilenceScanReportBuilder-0.21.52-windows-setup.exe) |
| Windows | [Portable ZIP](https://github.com/Windesheim-A-I-Support/ResilenceScanReportBuilder/releases/download/v0.21.52/ResilenceScanReportBuilder-0.21.52-windows-portable.zip) |
| Linux | [.deb (Ubuntu/Debian)](https://github.com/Windesheim-A-I-Support/ResilenceScanReportBuilder/releases/download/v0.21.52/ResilenceScanReportBuilder-0.21.52-amd64.deb) |
| Linux | [AppImage](https://github.com/Windesheim-A-I-Support/ResilenceScanReportBuilder/releases/download/v0.21.52/ResilenceScanReportBuilder-0.21.52-x86_64.AppImage) |
| Linux | [Tarball (.tar.gz)](https://github.com/Windesheim-A-I-Support/ResilenceScanReportBuilder/releases/download/v0.21.52/ResilenceScanReportBuilder-0.21.52-linux-amd64.tar.gz) |
<!-- DOWNLOAD_LINKS_END -->

> Direct download links are updated automatically after each release by CI.  If the links above point to the releases page rather than a specific file, a new release is in progress.

---

## What it does

1. **Import** — reads respondent data from `.xlsx` (or `.xml`) and converts it to a clean CSV
2. **Generate** — renders one PDF per respondent via `quarto render ResilienceReport.qmd` (R/LaTeX pipeline)
3. **Validate** — checks every generated PDF against the source CSV values
4. **Send** — emails each PDF to the right recipient via Outlook COM (Windows) or SMTP (Linux/Office365)

The GUI tracks send status per respondent across sessions so re-runs only send to people who haven't received their report yet.

---

## Architecture

```
app/main.py                          ← Tkinter GUI (entry point + PyInstaller target)
  ├── convert_data.py                ← Excel/XML → data/cleaned_master.csv
  ├── clean_data.py                  ← validates & cleans CSV in-place
  ├── generate_all_reports.py        ← loops CSV rows, calls quarto render per row
  │     └── ResilienceReport.qmd    ← Quarto template (R + LaTeX/TikZ)
  ├── validate_reports.py            ← checks PDF values against CSV
  ├── send_email.py                  ← Outlook COM (Windows) or SMTP fallback
  ├── email_tracker.py               ← persists per-recipient send status (JSON)
  ├── gui_system_check.py            ← verifies R / Quarto / TinyTeX at startup
  └── dependency_manager.py          ← stub (installation handled by installer)
```

### Key data files

| File | Description |
|------|-------------|
| `data/cleaned_master.csv` | Master respondent data — **never commit** |
| `reports/` | Generated PDFs — **never commit** |
| `%APPDATA%\ResilienceScan\config.yml` | SMTP credentials (Windows) |
| `~/.local/share/resiliencescan/config.yml` | SMTP credentials (Linux) |
| `C:\ProgramData\ResilienceScan\setup.log` | Installer setup progress (Windows) |

### Score columns in CSV

`up__r/c/f/v/a`, `in__r/c/f/v/a`, `do__r/c/f/v/a` — values 0–5

### PDF naming convention

```
YYYYMMDD ResilienceScanReport (Company Name - Firstname Lastname).pdf
```

---

## Pipeline flow

```
data/*.xlsx  (or .xml)
     │ convert_data.py
     ▼
data/cleaned_master.csv
     │ clean_data.py
     ▼
data/cleaned_master.csv  [validated]
     │ generate_all_reports.py  +  ResilienceReport.qmd
     │ (calls: quarto render --to pdf, uses R + TinyTeX)
     ▼
reports/YYYYMMDD ResilienceScanReport (Company - Person).pdf
     │ validate_reports.py  →  validate_single_report.py
     ▼
     │ send_email.py  +  email_tracker.py
     ▼
Outlook COM (Windows) or SMTP (Linux / Office365)
```

---

## PDF template

`ResilienceReport.qmd` is deeply LaTeX-dependent and **cannot be switched to Typst or WeasyPrint**.  It uses:

- TikZ (radar charts, custom graphics)
- kableExtra (styled tables)
- `_extensions/nmfs-opensci/titlepage` (custom title/cover pages)
- `QTDublinIrish.otf` (custom font)
- Raw `.tex` include files in `tex/`
- `references.bib` (bibliography)

The PDF engine is **TinyTeX** (pdflatex).  Required R packages: `readr dplyr stringr tidyr ggplot2 knitr fmsb scales viridis patchwork RColorBrewer gridExtra png lubridate kableExtra rmarkdown jsonlite ggrepel cowplot`.

---

## Packaging & installer

The application ships as a **staged installer** — the user runs a single setup wizard that silently downloads and installs all heavy dependencies (R, Quarto, TinyTeX) in the background.

### How the Windows installer works

```
NSIS .exe installer
  │ 1. Extracts PyInstaller bundle → C:\Program Files\ResilenceScanReportBuilder\
  │ 2. Runs launch_setup.ps1 via nsExec (returns in < 1 second)
  │      └── Registers + starts ResilienceScanSetup Task Scheduler task as SYSTEM
  │            └── Runs setup_dependencies.ps1 as SYSTEM in background
  │                  ├── Downloads + installs R 4.3.2 (CRAN /old/ archive)
  │                  ├── Downloads + installs Quarto 1.6.39 (MSI)
  │                  ├── quarto install tinytex  → SYSTEM AppData
  │                  │     └── Grants BUILTIN\Users RX on TinyTeX root
  │                  │     └── Adds TinyTeX bin to machine PATH (registry)
  │                  ├── tlmgr install <latex packages>
  │                  └── Rscript install.packages(<19 R packages> → r-library\)
  └── 3. User can launch app immediately; setup continues in background
```

**Setup logs** (Windows):

| File | Contents |
|------|----------|
| `C:\ProgramData\ResilienceScan\setup.log` | Step-by-step progress |
| `C:\ProgramData\ResilienceScan\setup_transcript.log` | Full stdout/stderr |
| `C:\ProgramData\ResilienceScan\setup_error.log` | Errors only |

### How the Linux installer works

`.deb` package with `postinst.sh` that launches `setup_linux.sh` via `nohup` + `disown` (deferred to avoid dpkg lock deadlock).  `setup_linux.sh` installs R, Quarto, TinyTeX, system libs for kableExtra, and all R packages.

### Pinned dependency versions

| Dependency | Version | Notes |
|------------|---------|-------|
| R | 4.3.2 | CRAN `/old/` archive URL |
| Quarto | 1.6.39 | GitHub releases |
| TinyTeX | Quarto-pinned | `quarto install tinytex` |
| Python | ≥ 3.11 | bundled by PyInstaller |

### Build artifacts per release

| Platform | Artifact | Notes |
|----------|----------|-------|
| Windows | `*-windows-setup.exe` | NSIS installer, ~200 MB download |
| Windows | `*-windows-portable.zip` | No installer, manual setup |
| Linux | `*-amd64.deb` | Debian/Ubuntu |
| Linux | `*-x86_64.AppImage` | Universal Linux |
| Linux | `*-linux-amd64.tar.gz` | Raw bundle |

### R library path (frozen app)

R packages are installed to `<InstallDir>\r-library\` (Windows) or `/opt/ResilenceScanReportBuilder/r-library/` (Linux).  At runtime the app injects `R_LIBS=<r-library path>` into every `quarto render` and `Rscript` subprocess so the bundled packages are always found regardless of system R library state.

---

## Runtime startup check

`gui_system_check.py` is called at every launch via `_startup_guard()`.  It checks:

1. **R** — `Rscript --version` (PATH + fallback: `C:\Program Files\R\R-*\bin\Rscript.exe`)
2. **Quarto** — `quarto --version` (PATH + fallback: `C:\Program Files\Quarto\bin\quarto.exe`)
3. **TinyTeX** — `tlmgr --version` (PATH + fallback: SYSTEM AppData + current user AppData)
4. **R packages** — `installed.packages()` with `R_LIBS` pointing at bundled library
5. **Python** — always passes (in-process)

On Windows, the checker first reads the machine PATH from the Windows registry (`HKLM\SYSTEM\...\Environment`) so it picks up paths added by the setup script after the user session started — without requiring a reboot.

If any critical check fails the app shows a dialog listing what is missing and exits.

---

## Local development setup

```bash
# Python environment
python -m venv .venv && source .venv/bin/activate   # Linux/macOS
python -m venv .venv && .venv\Scripts\activate       # Windows

pip install -r requirements.txt
pip install pytest ruff pyyaml PyPDF2                # dev tools

# Run the GUI
python app/main.py

# Run pipeline steps individually
python clean_data.py                 # clean data/cleaned_master.csv in-place
python generate_all_reports.py       # render one PDF per row → reports/
python validate_reports.py           # validate PDFs against CSV values
python send_email.py                 # send PDFs (TEST_MODE=True by default)

# Lint + test (mirrors CI)
ruff check .
ruff format --check .
pytest
```

### Prerequisites for local development

- Python ≥ 3.11
- R ≥ 4.3 with all 19 required packages installed
- Quarto ≥ 1.6
- TinyTeX (install via `quarto install tinytex`)

---

## Release workflow

```bash
# 1. Bump version in pyproject.toml
# 2. Push to main — CI runs lint + tests, then builds and publishes the release
git add pyproject.toml && git commit -m "v0.X.Y: description"
git push origin main
```

**Do not create tags manually** — CI detects that `v<version>` does not exist and creates it.  macOS is not a target.

### CI workflows

| Workflow | Trigger | Purpose |
|----------|---------|---------|
| `ci.yml` | every push / PR | lint (ruff check + format), pytest, version-check, build matrix, publish release |
| `codeql.yml` | push / PR / weekly | security analysis |

---

## Troubleshooting

### Startup guard fails on Windows after fresh install

The setup script runs as SYSTEM in the background after the installer exits.  **Wait for setup to finish** before launching the app (typically 5–15 minutes on a slow connection).  Check `C:\ProgramData\ResilienceScan\setup.log` — the last line should read `=== Dependency setup complete ===`.

If setup is complete but the guard still fails, the machine PATH update may not have reached the app's process.  v0.20.13+ reads PATH from the Windows registry directly, so this should be resolved.  If the problem persists, log out and back in so Explorer inherits the updated PATH.

### `quarto render` fails with "No valid input files passed to render"

`ResilienceReport.qmd` and its companion assets (`_extensions/`, `img/`, `tex/`, fonts, `references.bib`) must be reachable from the working directory passed to quarto.  In the frozen app, these are bundled inside `_internal/`.  See **Milestone 10** in `CLAUDE.md`.

### R packages not found at runtime

The app injects `R_LIBS=<InstallDir>\r-library` into all R subprocesses.  If packages still aren't found, check that `setup_dependencies.ps1` completed without errors (`setup_error.log` should be empty or contain only stale entries from previous runs).

### SMTP email not sending

Edit `%APPDATA%\ResilienceScan\config.yml` (Windows) or `~/.local/share/resiliencescan/config.yml` (Linux) with valid SMTP credentials, or use the **Configuration** tab in the GUI to save them.  On Windows the app tries Outlook COM first and falls back to SMTP.

---

## Project structure

```
├── app/
│   └── main.py                      # Tkinter GUI — entry point + PyInstaller target
├── packaging/
│   ├── launch_setup.ps1             # NSIS calls this; registers Task Scheduler task
│   ├── setup_dependencies.ps1       # runs as SYSTEM; installs R/Quarto/TinyTeX/packages
│   ├── setup_linux.sh               # Linux equivalent of setup_dependencies.ps1
│   ├── postinst.sh                  # .deb postinst — deferred setup_linux.sh via nohup
│   └── template.desktop             # Linux desktop entry
├── _extensions/                     # Quarto titlepage extension (committed)
│   └── nmfs-opensci/titlepage/
├── img/                             # Report images (logo, background, etc.)
├── tex/                             # Raw LaTeX include files for the template
├── tests/
│   └── test_smoke.py
├── ResilienceReport.qmd             # Quarto/R/LaTeX report template
├── references.bib                   # Bibliography for template
├── QTDublinIrish.otf               # Custom font used by template
├── clean_data.py                    # CSV cleaner / validator
├── convert_data.py                  # Excel/XML → CSV converter
├── generate_all_reports.py          # Batch report generator
├── generate_single_report.py        # Single-report generator (used by GUI)
├── validate_reports.py              # Batch validator
├── validate_single_report.py        # Single-report validator (shared logic)
├── send_email.py                    # Email sender
├── email_tracker.py                 # Per-recipient send status tracker
├── gui_system_check.py              # Runtime dependency checker
├── dependency_manager.py            # Stub (installation is installer's job)
├── nfpm.yaml                        # Linux package metadata (.deb / .rpm / AppImage)
├── pyproject.toml                   # Python project config + version
├── requirements.txt                 # Runtime Python dependencies
├── CLAUDE.md                        # Development milestones and architecture notes
└── CHANGELOG.md
```

---

## Milestone status

See [`CLAUDE.md`](CLAUDE.md) for the full milestone plan with gate conditions.

| Milestone | Description | Status |
|-----------|-------------|--------|
| M1 | Fix CI, ship real app | ✅ v0.13.0 |
| M2 | Fix paths, consolidate cleaners | ✅ v0.14.0 |
| M3 | Implement data conversion | ✅ v0.15.0 |
| M4 | End-to-end report generation | ✅ v0.16.0 |
| M5 | Fix validation | ✅ v0.17.0 |
| M6 | Email sending | ✅ v0.18.0 |
| M7 | Startup system check guard | ✅ v0.19.0 |
| M8 | Complete installer (R + Quarto + TinyTeX) | ✅ v0.20.5 |
| M9 | Fix Windows installer: R path, LaTeX packages, capt-of | 🔧 v0.20.12–14 |
| M10 | Fix report generation in installed app (.xlsx input + quarto path) | ⏳ TODO |
| M11 | Anonymised sample dataset + pipeline smoke test | ⏳ TODO |
| M12 | End-to-end CI pipeline test (Windows + Linux, real render) | ⏳ TODO |
| M13 | In-app update checker | ⏳ TODO |
| M14 | README download badges (auto-updated on release) | ⏳ TODO |
