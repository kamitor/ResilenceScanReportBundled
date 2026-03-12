"""
clean_data.py — validates and cleans data/cleaned_master.csv in-place.

Called by the GUI's "Clean Data" button via clean_and_fix() -> (bool, str).
Also runnable standalone: python clean_data.py
"""

import json
import shutil
import sys
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd

from utils.constants import REQUIRED_COLUMNS, SCORE_COLUMNS
from utils.path_utils import get_user_base_dir

_user_base = get_user_base_dir()

DATA_DIR = _user_base / "data"
INPUT_PATH = DATA_DIR / "cleaned_master.csv"
BACKUP_DIR = DATA_DIR / "backups"
VALIDATION_LOG = DATA_DIR / "cleaning_validation_log.json"
CLEANING_REPORT = DATA_DIR / "cleaning_report.txt"
REPLACEMENT_LOG = DATA_DIR / "value_replacements_log.csv"


class DataCleaningValidator:
    def __init__(self):
        self.issues = []
        self.warnings = []
        self.info = []
        self.removed_records = []
        self.statistics = {
            "initial_rows": 0,
            "final_rows": 0,
            "removed_rows": 0,
            "records_with_insufficient_data": 0,
            "records_with_no_scores": 0,
            "duplicates_removed": 0,
        }

    def log_issue(self, level, message, row_info=None):
        """Log an issue with details."""
        entry = {
            "level": level,
            "message": message,
            "timestamp": datetime.now().isoformat(),
        }
        if row_info:
            entry["row"] = row_info

        if level == "ERROR":
            self.issues.append(entry)
        elif level == "WARNING":
            self.warnings.append(entry)
        else:
            self.info.append(entry)

        symbol = "[X]" if level == "ERROR" else "[!]" if level == "WARNING" else "[i]"
        print(f"{symbol} [{level}] {message}")

    def create_backup(self, file_path):
        """Create a timestamped backup of a file."""
        try:
            if not Path(file_path).exists():
                self.log_issue("WARNING", f"Input file not found: {file_path}")
                return None

            BACKUP_DIR.mkdir(parents=True, exist_ok=True)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = Path(file_path).stem
            ext = Path(file_path).suffix
            backup_path = BACKUP_DIR / f"{filename}_{timestamp}{ext}"

            shutil.copy2(file_path, backup_path)
            self.log_issue("INFO", f"Backup created: {backup_path}")
            return backup_path
        except Exception as e:
            self.log_issue("ERROR", f"Failed to create backup: {e}")
            return None

    def validate_columns(self, df):
        """Validate that required columns exist."""
        print("\n" + "=" * 70)
        print("[INFO] COLUMN VALIDATION")
        print("=" * 70)

        df_cols_lower = [col.lower() for col in df.columns]
        missing_required = []

        for req_col in REQUIRED_COLUMNS:
            if req_col.lower() not in df_cols_lower:
                missing_required.append(req_col)
                self.log_issue("ERROR", f"Required column '{req_col}' not found")

        if missing_required:
            raise ValueError(f"Missing required columns: {missing_required}")

        missing_scores = [s for s in SCORE_COLUMNS if s.lower() not in df_cols_lower]
        if missing_scores:
            self.log_issue("WARNING", f"Missing score columns: {missing_scores}")

        self.log_issue("INFO", f"All required columns present: {REQUIRED_COLUMNS}")

    def validate_record_completeness(self, df):
        """Check each record for sufficient data to generate a report."""
        print("\n" + "=" * 70)
        print("[INFO] RECORD COMPLETENESS VALIDATION")
        print("=" * 70)

        records_to_keep = []

        for idx, row in df.iterrows():
            issues_found = []

            if (
                pd.isna(row.get("company_name"))
                or str(row.get("company_name", "")).strip() == ""
            ):
                issues_found.append("No company name")

            if pd.isna(row.get("name")) or str(row.get("name", "")).strip() == "":
                issues_found.append("No person name")

            if pd.isna(row.get("email_address")) or "@" not in str(
                row.get("email_address", "")
            ):
                issues_found.append("Invalid/missing email")

            # Count available scores
            available_scores = 0
            for score_col in SCORE_COLUMNS:
                if score_col in df.columns:
                    val = row[score_col]
                    if pd.notna(val) and val not in ["?", "", " "]:
                        try:
                            float_val = float(str(val).replace(",", "."))
                            if 0 <= float_val <= 5:
                                available_scores += 1
                        except ValueError:
                            pass

            min_scores_required = 5

            if issues_found:
                self.log_issue(
                    "WARNING",
                    f"Row {idx + 2}: {', '.join(issues_found)}",
                    {
                        "company": row.get("company_name", "N/A"),
                        "person": row.get("name", "N/A"),
                        "email": row.get("email_address", "N/A"),
                    },
                )
                self.removed_records.append(
                    {
                        "row": idx + 2,
                        "company": row.get("company_name", "N/A"),
                        "person": row.get("name", "N/A"),
                        "reason": ", ".join(issues_found),
                    }
                )
                continue

            if available_scores < min_scores_required:
                self.log_issue(
                    "WARNING",
                    f"Row {idx + 2}: Insufficient data ({available_scores}/15 scores)",
                    {
                        "company": row.get("company_name", "N/A"),
                        "person": row.get("name", "N/A"),
                        "available_scores": available_scores,
                    },
                )
                self.removed_records.append(
                    {
                        "row": idx + 2,
                        "company": row.get("company_name", "N/A"),
                        "person": row.get("name", "N/A"),
                        "reason": f"Only {available_scores} valid scores (need {min_scores_required})",
                    }
                )
                self.statistics["records_with_insufficient_data"] += 1
                continue

            records_to_keep.append(idx)

        if records_to_keep:
            df_clean = df.iloc[records_to_keep].copy()
            self.log_issue(
                "INFO",
                f"Kept {len(records_to_keep)}/{len(df)} records after validation",
            )
            return df_clean
        else:
            self.log_issue("ERROR", "No valid records remaining after validation!")
            return pd.DataFrame()

    def clean_score_columns(self, df):
        """Clean and convert score columns to numeric with detailed logging."""
        print("\n" + "=" * 70)
        print("[INFO] SCORE COLUMN CLEANING")
        print("=" * 70)

        total_replacements = 0
        replacement_log = []

        for col in SCORE_COLUMNS:
            if col not in df.columns:
                continue

            original_values = df[col].copy()
            df[col] = df[col].astype(str)

            invalid_mask = ~df[col].str.match(r"^[0-5](\.[0-9]+)?$", na=False)
            invalid_mask = invalid_mask & (df[col] != "nan")

            if invalid_mask.any():
                invalid_count = invalid_mask.sum()
                total_replacements += invalid_count

                for idx in df[invalid_mask].head(5).index:
                    original_val = original_values.loc[idx]
                    company = (
                        df.loc[idx, "company_name"]
                        if "company_name" in df.columns
                        else "Unknown"
                    )
                    person = df.loc[idx, "name"] if "name" in df.columns else "Unknown"
                    replacement_log.append(
                        {
                            "row": int(idx),
                            "company": company,
                            "person": person,
                            "column": col,
                            "original_value": str(original_val),
                            "action": "set_to_NaN (missing data)",
                        }
                    )

                self.log_issue(
                    "WARNING",
                    f"{col}: {invalid_count} invalid value(s) (e.g., '{original_values[invalid_mask].iloc[0]}')",
                )

            df[col] = df[col].replace(["?", "", " ", "nan"], np.nan)
            df[col] = df[col].str.replace(",", ".", regex=False)
            df[col] = df[col].str.replace(r"[^0-9.-]", "", regex=True)
            df[col] = pd.to_numeric(df[col], errors="coerce")
            df[col] = df[col].clip(lower=0, upper=5)

        if replacement_log:
            try:
                DATA_DIR.mkdir(parents=True, exist_ok=True)
                pd.DataFrame(replacement_log).to_csv(REPLACEMENT_LOG, index=False)
                self.log_issue(
                    "INFO",
                    f"Saved {len(replacement_log)} replacement details to: {REPLACEMENT_LOG}",
                )
            except Exception as e:
                self.log_issue("ERROR", f"Failed to save replacement log: {e}")

        if total_replacements > 0:
            self.log_issue(
                "WARNING", f"Total invalid values replaced: {total_replacements}"
            )
        else:
            self.log_issue("INFO", "All score values were valid")

        self.log_issue("INFO", f"Cleaned {len(SCORE_COLUMNS)} score columns")
        self.statistics["invalid_values_replaced"] = total_replacements

        return df

    def remove_duplicates(self, df):
        """Remove duplicate records."""
        print("\n" + "=" * 70)
        print("[INFO] DUPLICATE DETECTION")
        print("=" * 70)

        duplicates = df.duplicated(
            subset=["company_name", "email_address"], keep="first"
        )
        duplicate_count = duplicates.sum()

        if duplicate_count > 0:
            self.log_issue(
                "WARNING",
                f"Found {duplicate_count} duplicate records (keeping first occurrence)",
            )
            self.statistics["duplicates_removed"] = duplicate_count
            df = df[~duplicates]
        else:
            self.log_issue("INFO", "No duplicates found")

        return df

    def save_validation_log(self):
        """Save detailed validation log as JSON."""
        try:
            DATA_DIR.mkdir(parents=True, exist_ok=True)
            stats = {
                k: int(v) if isinstance(v, (np.int64, np.int32)) else v
                for k, v in self.statistics.items()
            }
            log_data = {
                "timestamp": datetime.now().isoformat(),
                "statistics": stats,
                "errors": self.issues,
                "warnings": self.warnings,
                "info": self.info,
                "removed_records": self.removed_records,
            }
            with open(VALIDATION_LOG, "w", encoding="utf-8") as f:
                json.dump(log_data, f, indent=2)
            self.log_issue("INFO", f"Validation log saved: {VALIDATION_LOG}")
        except Exception as e:
            self.log_issue("ERROR", f"Failed to save validation log: {e}")

    def generate_report(self):
        """Generate and save human-readable cleaning report."""
        lines = [
            "=" * 70,
            "DATA CLEANING REPORT",
            f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
            "=" * 70,
            "\nSTATISTICS:",
            f"  Initial rows: {self.statistics['initial_rows']}",
            f"  Final rows: {self.statistics['final_rows']}",
            f"  Removed rows: {self.statistics['removed_rows']}",
            f"  Duplicates removed: {self.statistics['duplicates_removed']}",
            f"  Records with insufficient data: {self.statistics['records_with_insufficient_data']}",
            "\nREMOVED RECORDS:",
        ]

        if self.removed_records:
            for record in self.removed_records:
                lines.append(
                    f"  Row {record['row']}: {record['company']} - {record['person']}"
                )
                lines.append(f"    Reason: {record['reason']}")
        else:
            lines.append("  None")

        lines += ["\nERRORS:"]
        lines += [f"  {i['message']}" for i in self.issues] or ["  None"]

        lines += ["\nWARNINGS:"]
        shown = self.warnings[:20]
        lines += [f"  {w['message']}" for w in shown]
        if len(self.warnings) > 20:
            lines.append(f"  ... and {len(self.warnings) - 20} more warnings")
        if not shown:
            lines.append("  None")

        lines += [
            "\n" + "=" * 70,
            "RECOMMENDATIONS:",
        ]
        if self.statistics["removed_rows"] > 0:
            lines.append(
                f"  [OK] {self.statistics['removed_rows']} records were excluded from the master CSV"
            )
            lines.append("  [OK] Review removed records above to ensure data quality")
        if self.statistics["final_rows"] == 0:
            lines.append("  [!] WARNING: No valid records remaining!")
        else:
            lines.append(
                f"  [OK] {self.statistics['final_rows']} valid records ready for report generation"
            )
        lines.append("=" * 70)

        report_text = "\n".join(lines)
        try:
            DATA_DIR.mkdir(parents=True, exist_ok=True)
            with open(CLEANING_REPORT, "w", encoding="utf-8") as f:
                f.write(report_text)
            print("\n" + report_text)
            self.log_issue("INFO", f"Cleaning report saved: {CLEANING_REPORT}")
        except Exception as e:
            print("\n" + report_text)
            self.log_issue("ERROR", f"Failed to save cleaning report: {e}")


def clean_and_fix():
    """
    Main entry point called by the GUI.
    Returns (success: bool, summary: str).
    """
    validator = DataCleaningValidator()

    # Ensure data directory exists
    try:
        DATA_DIR.mkdir(parents=True, exist_ok=True)
    except Exception as e:
        return False, f"Failed to create data directory {DATA_DIR}: {e}"

    # Check input file
    if not INPUT_PATH.exists():
        validator.log_issue("ERROR", f"Input file not found: {INPUT_PATH}")
        return False, "File not found — please run 'Convert Data' first"

    # Backup
    validator.create_backup(INPUT_PATH)

    # Load
    validator.log_issue("INFO", f"Loading data from: {INPUT_PATH}")
    try:
        df = pd.read_csv(INPUT_PATH, low_memory=False)
        validator.statistics["initial_rows"] = len(df)
        validator.log_issue("INFO", f"Loaded {len(df)} rows, {len(df.columns)} columns")
    except Exception as e:
        validator.log_issue("ERROR", f"Failed to load CSV: {e}")
        return False, f"Failed to load CSV: {e}"

    # Standardise column names
    df.columns = df.columns.str.lower().str.strip()

    # Validate columns
    try:
        validator.validate_columns(df)
    except ValueError as e:
        validator.log_issue("ERROR", str(e))
        validator.save_validation_log()
        validator.generate_report()
        return False, str(e)

    # Clean scores
    df = validator.clean_score_columns(df)

    # Completeness check
    df = validator.validate_record_completeness(df)

    if df.empty:
        validator.log_issue("ERROR", "No valid records after cleaning!")
        validator.statistics["final_rows"] = 0
        validator.statistics["removed_rows"] = validator.statistics["initial_rows"]
        validator.save_validation_log()
        validator.generate_report()
        return (
            False,
            "All records were removed during validation — no valid data remaining",
        )

    # Remove duplicates
    df = validator.remove_duplicates(df)

    # Update statistics
    validator.statistics["final_rows"] = len(df)
    validator.statistics["removed_rows"] = (
        validator.statistics["initial_rows"] - validator.statistics["final_rows"]
    )

    # Ensure reportsent column exists
    if "reportsent" not in df.columns:
        df["reportsent"] = False
        validator.log_issue("INFO", "Added 'reportsent' column (default: False)")

    # Save
    try:
        DATA_DIR.mkdir(parents=True, exist_ok=True)
        df.to_csv(INPUT_PATH, index=False)
        validator.log_issue("INFO", f"Saved cleaned data: {INPUT_PATH}")
    except Exception as e:
        validator.log_issue("ERROR", f"Failed to save CSV: {e}")
        return False, f"Failed to save cleaned data: {e}"

    # Artifacts
    validator.save_validation_log()
    validator.generate_report()

    print("\n" + "=" * 70)
    print("[OK] CLEANING COMPLETED SUCCESSFULLY")
    print("=" * 70)
    print(f"[INFO] Final dataset: {validator.statistics['final_rows']} records")
    print(f"[INFO] Removed: {validator.statistics['removed_rows']} records")
    print(f"[INFO] {CLEANING_REPORT}")
    print(f"[INFO] {VALIDATION_LOG}")
    print("=" * 70)

    # Summary for GUI
    summary_parts = []
    if validator.statistics["removed_rows"] > 0:
        summary_parts.append(
            f"Removed {validator.statistics['removed_rows']} invalid/incomplete record(s)"
        )
    if validator.statistics["duplicates_removed"] > 0:
        summary_parts.append(
            f"Removed {validator.statistics['duplicates_removed']} duplicate(s)"
        )
    if validator.statistics["records_with_insufficient_data"] > 0:
        summary_parts.append(
            f"Excluded {validator.statistics['records_with_insufficient_data']} record(s) with insufficient data"
        )

    summary = (
        "\n".join(summary_parts)
        if summary_parts
        else "All records passed validation — no changes needed!"
    )
    summary += f"\n\nFinal dataset: {validator.statistics['final_rows']} valid records ready for reports"
    summary += (
        f"\n\nDetailed reports saved to:\n- {CLEANING_REPORT}\n- {VALIDATION_LOG}"
    )

    return True, summary


if __name__ == "__main__":
    if sys.platform == "win32":
        import io

        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8")

    success, summary = clean_and_fix()
    level = "[OK]" if success else "[ERROR]"
    print(f"\n{level} {summary}")
    sys.exit(0 if success else 1)
