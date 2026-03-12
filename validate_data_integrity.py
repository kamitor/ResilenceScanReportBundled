import pandas as pd
from pathlib import Path
import glob
import json
from datetime import datetime
import random

# Configuration
DATA_DIR = "/app/data"
CLEANED_CSV = "/app/outputs/cleaned_master.csv"
VALIDATION_OUTPUT = "/app/outputs/integrity_validation_report.json"
REPORT_OUTPUT = "/app/outputs/integrity_validation_report.txt"

# Excel file patterns (same as convert_data.py)
FILE_PATTERNS = [
    "Resilience - MasterDatabase*.xlsx",
    "Resilience - MasterDatabase*.xls",
    "MasterDatabase*.xlsx",
    "MasterDatabase*.xls",
    "*.xlsx",
    "*.xls",
]


class DataIntegrityValidator:
    def __init__(self):
        self.issues = []
        self.warnings = []
        self.info = []
        self.samples = []
        self.statistics = {
            "total_records_excel": 0,
            "total_records_csv": 0,
            "samples_validated": 0,
            "perfect_matches": 0,
            "acceptable_matches": 0,
            "mismatches": 0,
            "missing_in_csv": 0,
        }

    def log(self, level, message, details=None):
        """Log a message with level"""
        entry = {
            "level": level,
            "message": message,
            "timestamp": datetime.now().isoformat(),
        }
        if details:
            entry["details"] = details

        if level == "ERROR":
            self.issues.append(entry)
        elif level == "WARNING":
            self.warnings.append(entry)
        else:
            self.info.append(entry)

        symbol = "[X]" if level == "ERROR" else "[!]" if level == "WARNING" else "[i]"
        print(f"{symbol} [{level}] {message}")

    def find_excel_file(self):
        """Find the source Excel file"""
        print("\n" + "=" * 70)
        print("[INFO] FINDING SOURCE EXCEL FILE")
        print("=" * 70)

        # Check if data directory exists
        if not Path(DATA_DIR).exists():
            self.log(
                "ERROR",
                f"No data files found: data directory does not exist at {DATA_DIR}",
            )
            return None

        for pattern in FILE_PATTERNS:
            search_path = Path(DATA_DIR) / pattern
            matches = list(glob.glob(str(search_path)))

            if matches:
                matches.sort(key=lambda x: Path(x).stat().st_mtime, reverse=True)
                excel_file = matches[0]
                self.log("INFO", f"Found Excel file: {excel_file}")
                if len(matches) > 1:
                    self.log("INFO", "Multiple files found, using most recent")
                return excel_file

        self.log(
            "ERROR",
            "No data files found: no Excel files matching expected patterns in data directory",
        )
        return None

    def load_excel_data(self, excel_path):
        """Load data from Excel file"""
        print("\n" + "=" * 70)
        print("[INFO] LOADING EXCEL DATA")
        print("=" * 70)

        try:
            # Load with no header to detect it ourselves
            df_raw = pd.read_excel(excel_path, header=None)
            self.log(
                "INFO",
                f"Loaded raw Excel: {df_raw.shape[0]} rows x {df_raw.shape[1]} columns",
            )

            # Detect header row (same logic as convert_data.py)
            header_keywords = [
                "company",
                "name",
                "email",
                "submitdate",
                "up -",
                "in -",
                "do -",
            ]
            header_row_idx = 0

            for idx in range(min(10, len(df_raw))):
                row_strings = [
                    str(v).lower() if pd.notna(v) else "" for v in df_raw.iloc[idx]
                ]
                keyword_matches = sum(
                    any(keyword in s for keyword in header_keywords)
                    for s in row_strings
                )
                if keyword_matches >= 3:
                    header_row_idx = idx
                    self.log("INFO", f"Detected header at row {idx + 1}")
                    break

            # Extract data with proper header
            header = df_raw.iloc[header_row_idx].tolist()
            data = df_raw.iloc[header_row_idx + 1 :]
            df = pd.DataFrame(data.values, columns=header)

            # Clean column names (same as convert_data.py)
            import re

            cleaned_cols = []
            for col in df.columns:
                col_str = str(col).strip().lower()
                col_str = (
                    col_str.replace(" ", "_")
                    .replace("-", "")
                    .replace(":", "")
                    .replace("(", "")
                    .replace(")", "")
                )
                col_str = re.sub(r"[^\w_]", "", col_str)
                cleaned_cols.append(col_str)
            df.columns = cleaned_cols

            # Remove empty rows
            df = df.dropna(how="all")

            self.statistics["total_records_excel"] = len(df)
            self.log("INFO", f"Excel data rows: {len(df)}")

            return df

        except Exception as e:
            self.log("ERROR", f"Failed to load Excel: {e}")
            return None

    def load_csv_data(self):
        """Load cleaned CSV data"""
        print("\n" + "=" * 70)
        print("[INFO] LOADING CLEANED CSV")
        print("=" * 70)

        if not Path(CLEANED_CSV).exists():
            self.log(
                "ERROR",
                f"ERROR: missing input file - Cleaned CSV not found at {CLEANED_CSV}",
            )
            return None

        try:
            df = pd.read_csv(CLEANED_CSV)
            df.columns = df.columns.str.lower().str.strip()
            self.statistics["total_records_csv"] = len(df)
            self.log("INFO", f"CSV data rows: {len(df)}")
            return df

        except Exception as e:
            self.log(
                "ERROR", f"ERROR: failed to load input data - Failed to load CSV: {e}"
            )
            return None

    def create_record_key(self, row):
        """Create a unique key for a record"""
        company = str(row.get("company_name", "")).strip().lower()
        name = str(row.get("name", "")).strip().lower()
        email = str(row.get("email_address", "")).strip().lower()
        return f"{company}||{name}||{email}"

    def compare_score_values(self, excel_val, csv_val, tolerance=0.01):
        """Compare two score values with tolerance"""
        # Handle various representations of missing/invalid data
        excel_empty = pd.isna(excel_val) or str(excel_val).strip() in ["", "?", "nan"]
        csv_empty = pd.isna(csv_val) or str(csv_val).strip() in ["", "?", "nan"]

        if excel_empty and csv_empty:
            return "match", "both_empty"
        if excel_empty != csv_empty:
            return "mismatch", "one_empty"

        try:
            # Convert to float, handling European decimals
            excel_float = float(str(excel_val).replace(",", "."))
            csv_float = float(csv_val)

            # Check if values match within tolerance
            if abs(excel_float - csv_float) <= tolerance:
                return "match", "values_equal"
            else:
                return (
                    "mismatch",
                    f"values_differ: {excel_float:.2f} vs {csv_float:.2f}",
                )
        except:
            return "mismatch", "conversion_error"

    def validate_sample(self, excel_row, csv_row, sample_num):
        """Validate a single sampled record"""
        validation_result = {
            "sample_num": sample_num,
            "company": excel_row.get("company_name", "Unknown"),
            "person": excel_row.get("name", "Unknown"),
            "email": excel_row.get("email_address", "Unknown"),
            "fields_checked": 0,
            "fields_matched": 0,
            "mismatches": [],
            "status": "unknown",
        }

        # Check basic fields
        basic_fields = ["company_name", "name", "email_address"]
        for field in basic_fields:
            excel_val = str(excel_row.get(field, "")).strip()
            csv_val = str(csv_row.get(field, "")).strip()
            validation_result["fields_checked"] += 1

            if excel_val.lower() == csv_val.lower():
                validation_result["fields_matched"] += 1
            else:
                validation_result["mismatches"].append(
                    {"field": field, "excel": excel_val, "csv": csv_val}
                )

        # Check score columns
        score_columns = [
            "up__r",
            "up__c",
            "up__f",
            "up__v",
            "up__a",
            "in__r",
            "in__c",
            "in__f",
            "in__v",
            "in__a",
            "do__r",
            "do__c",
            "do__f",
            "do__v",
            "do__a",
        ]

        for col in score_columns:
            if col in excel_row.index and col in csv_row.index:
                validation_result["fields_checked"] += 1
                status, detail = self.compare_score_values(excel_row[col], csv_row[col])

                if status == "match":
                    validation_result["fields_matched"] += 1
                else:
                    validation_result["mismatches"].append(
                        {
                            "field": col,
                            "excel": str(excel_row[col]),
                            "csv": str(csv_row[col]),
                            "detail": detail,
                        }
                    )

        # Determine overall status
        if validation_result["fields_checked"] == 0:
            validation_result["status"] = "no_fields"
        elif validation_result["fields_matched"] == validation_result["fields_checked"]:
            validation_result["status"] = "perfect"
            self.statistics["perfect_matches"] += 1
        elif (
            validation_result["fields_matched"] / validation_result["fields_checked"]
            >= 0.9
        ):
            validation_result["status"] = "acceptable"
            self.statistics["acceptable_matches"] += 1
        else:
            validation_result["status"] = "mismatch"
            self.statistics["mismatches"] += 1

        return validation_result

    def validate_samples(self, excel_df, csv_df, num_samples=10):
        """Validate random samples from Excel against CSV"""
        print("\n" + "=" * 70)
        print(f"VALIDATING {num_samples} RANDOM SAMPLES")
        print("=" * 70)

        # Create lookup dictionary for CSV records
        csv_lookup = {}
        for idx, row in csv_df.iterrows():
            key = self.create_record_key(row)
            csv_lookup[key] = row

        self.log("INFO", f"Created CSV lookup with {len(csv_lookup)} unique records")

        # Sample random records from Excel
        if len(excel_df) < num_samples:
            num_samples = len(excel_df)
            self.log(
                "WARNING", f"Excel has fewer than {num_samples} records, sampling all"
            )

        sample_indices = random.sample(range(len(excel_df)), num_samples)

        for i, idx in enumerate(sample_indices, 1):
            excel_row = excel_df.iloc[idx]
            key = self.create_record_key(excel_row)

            print(
                f"\n[{i}/{num_samples}] Validating: {excel_row.get('company_name', 'Unknown')}"
            )

            if key not in csv_lookup:
                self.log(
                    "WARNING",
                    f"Record not found in CSV: {excel_row.get('company_name', 'Unknown')}",
                )
                self.statistics["missing_in_csv"] += 1
                self.samples.append(
                    {
                        "sample_num": i,
                        "company": excel_row.get("company_name", "Unknown"),
                        "status": "missing_in_csv",
                    }
                )
                continue

            csv_row = csv_lookup[key]
            validation_result = self.validate_sample(excel_row, csv_row, i)
            self.samples.append(validation_result)
            self.statistics["samples_validated"] += 1

            # Log result
            if validation_result["status"] == "perfect":
                self.log(
                    "INFO",
                    f"Perfect match: {validation_result['fields_matched']}/{validation_result['fields_checked']} fields",
                )
            elif validation_result["status"] == "acceptable":
                self.log(
                    "INFO",
                    f"Acceptable match: {validation_result['fields_matched']}/{validation_result['fields_checked']} fields",
                )
            else:
                self.log(
                    "WARNING",
                    f"Mismatches found: {len(validation_result['mismatches'])} issues",
                )

    def generate_report(self):
        """Generate human-readable report"""
        import os

        lines = []
        lines.append("=" * 70)
        lines.append("DATA INTEGRITY VALIDATION REPORT")
        lines.append(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        lines.append("=" * 70)

        lines.append("\nOVERVIEW:")
        lines.append(f"  Excel file records: {self.statistics['total_records_excel']}")
        lines.append(f"  CSV file records: {self.statistics['total_records_csv']}")
        lines.append(
            f"  Records removed during cleaning: {self.statistics['total_records_excel'] - self.statistics['total_records_csv']}"
        )

        lines.append("\nSAMPLE VALIDATION RESULTS:")
        lines.append(f"  Samples validated: {self.statistics['samples_validated']}")
        lines.append(f"  Perfect matches: {self.statistics['perfect_matches']}")
        lines.append(
            f"  Acceptable matches (>90%): {self.statistics['acceptable_matches']}"
        )
        lines.append(f"  Mismatches: {self.statistics['mismatches']}")
        lines.append(f"  Missing in CSV: {self.statistics['missing_in_csv']}")

        accuracy = 0
        if self.statistics["samples_validated"] > 0:
            accuracy = (
                (
                    self.statistics["perfect_matches"]
                    + self.statistics["acceptable_matches"]
                )
                / self.statistics["samples_validated"]
                * 100
            )
            lines.append(f"\n  Overall accuracy: {accuracy:.1f}%")

        lines.append("\nDETAILED SAMPLE RESULTS:")
        for sample in self.samples:
            if sample.get("status") == "missing_in_csv":
                lines.append(f"\n  [{sample['sample_num']}] {sample['company']}")
                lines.append("      Status: MISSING IN CSV")
            else:
                lines.append(
                    f"\n  [{sample['sample_num']}] {sample['company']} - {sample['person']}"
                )
                lines.append(f"      Status: {sample['status'].upper()}")
                lines.append(
                    f"      Matched: {sample['fields_matched']}/{sample['fields_checked']} fields"
                )

                if sample["mismatches"]:
                    lines.append("      Mismatches:")
                    for mm in sample["mismatches"][:5]:  # Limit to first 5
                        lines.append(
                            f"        - {mm['field']}: Excel='{mm['excel']}' vs CSV='{mm['csv']}'"
                        )
                    if len(sample["mismatches"]) > 5:
                        lines.append(
                            f"        ... and {len(sample['mismatches']) - 5} more"
                        )

        lines.append("\n" + "=" * 70)
        lines.append("VERDICT:")
        if self.statistics["samples_validated"] == 0:
            lines.append("  [!] No samples could be validated")
            lines.append("  [!] All sampled records were missing in CSV")
            lines.append("  [!] This suggests records were removed during cleaning")
        elif (
            self.statistics["mismatches"] == 0
            and self.statistics["missing_in_csv"] == 0
        ):
            lines.append("  [OK] Data integrity verified!")
            lines.append("  [OK] Cleaning process preserves data accurately")
        elif accuracy >= 90:
            lines.append("  [OK] Data integrity is acceptable (>90% accuracy)")
            lines.append("  [!] Minor discrepancies found - review details above")
        else:
            lines.append("  [!] Data integrity issues detected")
            lines.append("  [!] Significant discrepancies between Excel and CSV")
            lines.append("  [!] Review cleaning process")

        lines.append("=" * 70)

        report_text = "\n".join(lines)
        print("\n" + report_text)

        try:
            # Ensure output directory exists
            output_dir = Path(REPORT_OUTPUT).parent
            os.makedirs(output_dir, exist_ok=True)

            with open(REPORT_OUTPUT, "w", encoding="utf-8") as f:
                f.write(report_text)

            self.log("INFO", f"Report saved: {REPORT_OUTPUT}")
        except Exception as e:
            self.log(
                "ERROR", f"ERROR: failed to write output - Failed to save report: {e}"
            )

    def save_validation_log(self):
        """Save detailed JSON log"""
        import os

        log_data = {
            "timestamp": datetime.now().isoformat(),
            "statistics": self.statistics,
            "samples": self.samples,
            "errors": self.issues,
            "warnings": self.warnings,
            "info": self.info,
        }

        try:
            # Ensure output directory exists
            output_dir = Path(VALIDATION_OUTPUT).parent
            os.makedirs(output_dir, exist_ok=True)

            with open(VALIDATION_OUTPUT, "w", encoding="utf-8") as f:
                json.dump(log_data, f, indent=2)

            self.log("INFO", f"Validation log saved: {VALIDATION_OUTPUT}")
        except Exception as e:
            self.log(
                "ERROR",
                f"ERROR: failed to write output - Failed to save validation log: {e}",
            )


def main(num_samples=10):
    """Main validation function"""
    print("=" * 70)
    print("[INFO] DATA INTEGRITY VALIDATOR")
    print("[INFO] Validates that cleaning process preserves data accurately")
    print("=" * 70)

    validator = DataIntegrityValidator()

    # Find and load Excel file
    excel_file = validator.find_excel_file()
    if not excel_file:
        print("\n[ERROR] Cannot proceed without Excel file")
        return False

    excel_df = validator.load_excel_data(excel_file)
    if excel_df is None:
        print("\n[ERROR] Cannot load Excel data")
        return False

    # Load CSV file
    csv_df = validator.load_csv_data()
    if csv_df is None:
        print("\n[ERROR] Cannot load CSV data")
        return False

    # Validate samples
    validator.validate_samples(excel_df, csv_df, num_samples)

    # Generate reports
    validator.generate_report()
    validator.save_validation_log()

    print("\n" + "=" * 70)
    print("[OK] Data integrity validation finished")
    print("=" * 70)
    print(f"[INFO] Text report: {REPORT_OUTPUT}")
    print(f"[INFO] JSON log: {VALIDATION_OUTPUT}")
    print("=" * 70)

    return True


if __name__ == "__main__":
    import sys

    # Allow specifying number of samples via command line
    num_samples = 10
    if len(sys.argv) > 1:
        try:
            num_samples = int(sys.argv[1])
        except:
            print("Usage: python validate_data_integrity.py [num_samples]")
            sys.exit(1)

    success = main(num_samples)
    exit(0 if success else 1)
