"""
validate_reports.py — validates generated PDFs against cleaned_master.csv.

Scans the reports/ directory for PDF files, parses company and person names
from each filename, then delegates per-report validation to
validate_single_report.validate_report().  No external JSON file required.
"""

import re
from pathlib import Path

from validate_single_report import validate_report

from utils.path_utils import get_user_base_dir

_user_base = get_user_base_dir()

DATA_FILE = _user_base / "data" / "cleaned_master.csv"
REPORTS_DIR = _user_base / "reports"

# Matches the parenthesised "(Company - Person)" in the PDF filename.
# Handles both current and legacy report name variants.
_PARENS_RE = re.compile(r"\((.+)\)\.pdf$", re.IGNORECASE)


def _parse_pdf_filename(filename: str) -> tuple[str, str] | None:
    """Return (company, person) from a PDF filename, or None if unparseable.

    Expected format:
      YYYYMMDD ResilienceScanReport (Company Name - Firstname Lastname).pdf
    Legacy format also accepted:
      YYYYMMDD ResilienceReport (Company Name - Firstname Lastname).pdf
    """
    m = _PARENS_RE.search(filename)
    if not m:
        return None
    content = m.group(1)
    if " - " not in content:
        return None
    company, person = content.rsplit(" - ", 1)
    return company.strip(), person.strip()


def validate_all(
    reports_dir: Path = REPORTS_DIR,
    csv_path: Path = DATA_FILE,
) -> dict:
    """Validate every PDF in reports_dir against csv_path.

    Returns a summary dict:
      {"total": int, "passed": int, "failed": int, "errors": int, "pass_rate": float}
    """
    print("=" * 70)
    print("[INFO] PDF REPORT VALIDATION")
    print("=" * 70)

    empty = {"total": 0, "passed": 0, "failed": 0, "errors": 0, "pass_rate": 0.0}

    if not csv_path.exists():
        print(f"[ERROR] CSV not found: {csv_path}")
        return empty

    if not reports_dir.exists():
        print(f"[ERROR] Reports directory not found: {reports_dir}")
        return empty

    pdf_files = sorted(reports_dir.glob("*.pdf"))
    if not pdf_files:
        print("[INFO] No PDF files found in reports/ directory")
        return empty

    print(f"\n[INFO] Found {len(pdf_files)} PDF(s) in {reports_dir}")
    print(f"[INFO] CSV: {csv_path}\n")
    print("=" * 70)

    passed = 0
    failed = 0
    errors = 0

    for pdf_path in pdf_files:
        parsed = _parse_pdf_filename(pdf_path.name)
        if parsed is None:
            print(f"\n[SKIP] Cannot parse filename: {pdf_path.name}")
            errors += 1
            continue

        company, person = parsed
        print(f"\n[{company} - {person}]")
        print(f"   File: {pdf_path.name}")

        result = validate_report(
            pdf_path=str(pdf_path),
            csv_path=str(csv_path),
            company_name=company,
            person_name=person,
        )

        if result["success"]:
            print(f"   [OK] {result['message']}")
            passed += 1
        else:
            print(f"   [FAIL] {result['message']}")
            for info in result.get("details", {}).values():
                if not info.get("matches"):
                    exp = (
                        f"{info['expected']:.2f}"
                        if info.get("expected") is not None
                        else "N/A"
                    )
                    act = (
                        f"{info['actual']:.2f}"
                        if info.get("actual") is not None
                        else "N/A"
                    )
                    print(
                        f"      {info.get('label', '?')}: Expected={exp}, Actual={act}"
                    )
            failed += 1

    total = passed + failed + errors
    pass_rate = (passed / (passed + failed) * 100) if (passed + failed) > 0 else 0.0

    print("\n" + "=" * 70)
    print("[INFO] VALIDATION SUMMARY")
    print("=" * 70)
    print(f"   Total PDFs:  {total}")
    print(f"   Passed:      {passed}")
    print(f"   Failed:      {failed}")
    print(f"   Errors/skip: {errors}")
    print(f"   Pass rate:   {pass_rate:.1f}%  (of parseable reports)")

    if pass_rate >= 90.0:
        print("\n   [OK] Gate passed -- pass rate >= 90%")
    else:
        print("\n   [WARN] Gate not met -- pass rate < 90%")
    print("=" * 70)

    return {
        "total": total,
        "passed": passed,
        "failed": failed,
        "errors": errors,
        "pass_rate": pass_rate,
    }


if __name__ == "__main__":
    summary = validate_all()
    exit(0 if summary["pass_rate"] >= 90.0 else 1)
