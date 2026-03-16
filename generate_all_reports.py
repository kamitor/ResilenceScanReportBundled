import subprocess
import csv
import shutil
from datetime import datetime
from pathlib import Path

import pandas as pd

from utils.constants import QUARTO_TIMEOUT_SECONDS
from utils.filename_utils import safe_display_name, safe_filename

# NOTE: dev-only CLI tool — paths are relative to repo root and will not work
# in the frozen/installed app.  Use the GUI (app/main.py) for production use.
ROOT = Path(__file__).resolve().parent
TEMPLATE = ROOT / "ResilienceReport.qmd"
DATA = ROOT / "data" / "cleaned_master.csv"
OUTPUT_DIR = ROOT / "reports"
COLUMN_MATCH_COMPANY = "company_name"
COLUMN_MATCH_PERSON = "name"


def load_csv(path):
    """Load CSV with encoding and delimiter detection"""
    encodings = ["utf-8", "cp1252", "latin1"]
    for enc in encodings:
        try:
            with open(path, encoding=enc) as f:
                # Read enough lines for reliable delimiter detection
                lines = []
                for _ in range(5):
                    line = f.readline()
                    if not line:
                        break
                    lines.append(line)
                sample = "".join(lines)
                try:
                    sep = csv.Sniffer().sniff(sample, delimiters=",;\t|").delimiter
                    print(f"[OK] Delimiter '{sep}' with encoding '{enc}'")
                except Exception:
                    sep = ","
                    print(f"[WARN] Using fallback delimiter ',' with encoding '{enc}'")
                return pd.read_csv(path, encoding=enc, sep=sep)
        except Exception as e:
            print(f"[WARN] Failed with encoding {enc}: {e}")
    raise RuntimeError("[ERROR] Could not read CSV.")


def generate_reports():
    """Generate individual PDF reports for each person/company entry"""

    print("=" * 70)
    print("[INFO] RESILIENCE SCAN REPORT GENERATOR")
    print("=" * 70)

    # Load data
    df = load_csv(DATA)
    df.columns = df.columns.str.lower().str.strip()

    # Find required columns
    company_col = next((col for col in df.columns if COLUMN_MATCH_COMPANY in col), None)
    person_col = next((col for col in df.columns if COLUMN_MATCH_PERSON in col), None)

    if not company_col:
        raise ValueError(f"[ERROR] No column matching '{COLUMN_MATCH_COMPANY}'")

    print("\n[INFO] Found columns:")
    print(f"   Company: {company_col}")
    print(f"   Person: {person_col if person_col else 'Not found (will use Unknown)'}")

    # Create output directory
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    # Get current date for filename
    date_str = datetime.now().strftime("%Y%m%d")

    # Count total and generate reports
    total_entries = len(df[df[company_col].notna()])
    print(f"\n[OUTPUT] Total entries to process: {total_entries}")
    print("=" * 70)

    generated = 0
    skipped = 0
    failed = 0

    for idx, row in df.iterrows():
        company = row[company_col]

        # Skip rows without company name
        if pd.isna(company) or str(company).strip() == "":
            continue

        # Get person name (fallback to "Unknown")
        if person_col and person_col in row and not pd.isna(row[person_col]):
            person = row[person_col]
        else:
            person = "Unknown"

        # Create safe filenames for temp file (underscore-based)
        safe_company = safe_filename(company)
        safe_person = safe_filename(person)

        # Create safe display names for final filename (readable but safe)
        display_company = safe_display_name(company)
        display_person = safe_display_name(person)

        # New naming format: YYYYMMDD ResilienceScanReport (COMPANY NAME - Firstname Lastname).pdf
        output_filename = f"{date_str} ResilienceScanReport ({display_company} - {display_person}).pdf"
        output_file = OUTPUT_DIR / output_filename

        # Check if already exists
        if output_file.exists():
            print(f"[INFO] Skipping {company} - {person} (already exists)")
            skipped += 1
            continue

        print(f"\n[INFO] Generating report {generated + 1}/{total_entries}:")
        print(f"   Company: {company}")
        print(f"   Person: {person}")
        print(f"   Output: {output_filename}")

        # Build quarto command with both company and person parameters
        temp_output = f"temp_{safe_company}_{safe_person}.pdf"
        cmd = [
            "quarto",
            "render",
            str(TEMPLATE),
            "-P",
            f"company={company}",
            "-P",
            f"person={person}",
            "--to",
            "pdf",
            "--output",
            temp_output,
        ]

        # Execute quarto render with verbose output
        print("   [INFO] Running: quarto render...")
        try:
            result = subprocess.run(
                cmd,
                cwd=ROOT,
                capture_output=True,
                text=True,
                timeout=QUARTO_TIMEOUT_SECONDS,
            )

            if result.returncode == 0:
                temp_path = ROOT / temp_output
                if temp_path.exists():
                    shutil.move(str(temp_path), str(output_file))
                    print(f"   [OK] Saved: {output_file}")
                    generated += 1
                else:
                    print("   [ERROR] Output file not found after successful render")
                    print(
                        f"   [INFO] stdout: {result.stdout[-500:]}"
                        if result.stdout
                        else ""
                    )
                    print(
                        f"   [WARN]  stderr: {result.stderr[-500:]}"
                        if result.stderr
                        else ""
                    )
                    failed += 1
            else:
                print(f"   [ERROR] Failed (exit code: {result.returncode})")
                print("   [INFO] Error output:")
                if result.stderr:
                    # Show first 2000 chars so root cause error is visible
                    error_text = (
                        result.stderr[:2000]
                        if len(result.stderr) > 2000
                        else result.stderr
                    )
                    print(f"   {error_text}")
                if result.stdout:
                    # Show last 500 chars of stdout if available
                    stdout_text = (
                        result.stdout[-500:]
                        if len(result.stdout) > 500
                        else result.stdout
                    )
                    print(f"   [INFO] Output: {stdout_text}")
                failed += 1

        except subprocess.TimeoutExpired:
            print("   [ERROR] Failed: Timeout after 300 seconds")
            failed += 1
        except Exception as e:
            print(f"   [ERROR] Failed: {type(e).__name__}: {e}")
            failed += 1
        finally:
            temp_path = ROOT / temp_output
            if temp_path.exists():
                temp_path.unlink()

    # Summary
    print("\n" + "=" * 70)
    print("[INFO] GENERATION SUMMARY")
    print("=" * 70)
    print(f"   [OK] Generated: {generated}")
    print(f"   [SKIP] Skipped:   {skipped}")
    print(f"   [ERROR] Failed:    {failed}")
    print(f"   [INFO] Total:     {total_entries}")
    print("=" * 70)

    if generated > 0:
        print(f"\n[OK] Reports saved to: {OUTPUT_DIR}")


if __name__ == "__main__":
    try:
        generate_reports()
    except Exception as e:
        print(f"\n[ERROR] ERROR: {e}")
        raise
