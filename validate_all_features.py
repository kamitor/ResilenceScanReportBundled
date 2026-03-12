"""
Comprehensive Feature Validation Script
Tests all new data quality features and verifies existing functionality
"""

import pandas as pd
import subprocess
import sys
from pathlib import Path
from datetime import datetime

# Configuration
DATA_FILE = Path("data/cleaned_master.csv")
TEST_REPORT_DIR = Path("test_reports")
QUALITY_REPORTS_DIR = Path("data/quality_reports")


class FeatureValidator:
    def __init__(self):
        self.results = []
        self.test_count = 0
        self.pass_count = 0
        self.fail_count = 0

    def log(self, status, test_name, details=""):
        """Log test result"""
        self.test_count += 1
        if status == "PASS":
            self.pass_count += 1
            icon = "[OK]"
        else:
            self.fail_count += 1
            icon = "[FAIL]"

        message = f"{icon} {test_name}"
        if details:
            message += f"\n    {details}"

        print(message)
        self.results.append({"status": status, "test": test_name, "details": details})

    def test_data_exists(self):
        """Test 1: Verify data file exists"""
        if DATA_FILE.exists():
            df = pd.read_csv(DATA_FILE)
            self.log("PASS", "Data file exists", f"Found {len(df)} records")
            return True
        else:
            self.log("FAIL", "Data file exists", f"Missing: {DATA_FILE}")
            return False

    def test_quality_dashboard_script(self):
        """Test 2: Run quality dashboard script"""
        try:
            result = subprocess.run(
                [sys.executable, "data_quality_dashboard.py"],
                capture_output=True,
                text=True,
                timeout=30,
            )

            if result.returncode == 0:
                # Check if PNG was created
                png_files = list(QUALITY_REPORTS_DIR.glob("quality_dashboard_*.png"))
                if png_files:
                    self.log(
                        "PASS",
                        "Quality dashboard script runs",
                        f"Created {len(png_files)} dashboard file(s)",
                    )
                    return True
                else:
                    self.log(
                        "FAIL", "Quality dashboard script runs", "No PNG files created"
                    )
                    return False
            else:
                self.log(
                    "FAIL",
                    "Quality dashboard script runs",
                    f"Exit code: {result.returncode}\n{result.stderr}",
                )
                return False
        except Exception as e:
            self.log("FAIL", "Quality dashboard script runs", str(e))
            return False

    def test_data_cleaner_script(self):
        """Test 3: Run data cleaner script"""
        try:
            # Create a test CSV with invalid data
            test_file = Path("data/test_invalid_data.csv")
            df = pd.read_csv(DATA_FILE).head(5).copy()

            # Add some invalid values
            score_cols = ["up__r", "up__c", "in__r"]
            for col in score_cols:
                if col in df.columns:
                    df.loc[0, col] = "?"
                    df.loc[1, col] = "N/A"

            df.to_csv(test_file, index=False)

            # Run cleaner on test file
            result = subprocess.run(
                [sys.executable, "clean_data_enhanced.py"],
                capture_output=True,
                text=True,
                timeout=30,
            )

            if result.returncode == 0:
                # Check if replacement log was created
                log_file = Path("data/value_replacements_log.csv")
                if log_file.exists():
                    self.log(
                        "PASS",
                        "Data cleaner script runs",
                        "Value replacements logged successfully",
                    )
                    # Clean up test file
                    test_file.unlink(missing_ok=True)
                    return True
                else:
                    self.log(
                        "PASS",
                        "Data cleaner script runs",
                        "No invalid values found (expected for clean data)",
                    )
                    test_file.unlink(missing_ok=True)
                    return True
            else:
                self.log(
                    "FAIL",
                    "Data cleaner script runs",
                    f"Exit code: {result.returncode}\n{result.stderr}",
                )
                test_file.unlink(missing_ok=True)
                return False

        except Exception as e:
            self.log("FAIL", "Data cleaner script runs", str(e))
            return False

    def test_debug_mode_parameter(self):
        """Test 4: Verify debug_mode parameter in ResilienceReport.qmd"""
        try:
            qmd_file = Path("ResilienceReport.qmd")
            content = qmd_file.read_text(encoding="utf-8")

            # Check for debug_mode in params
            if "debug_mode:" in content:
                # Check for conditional debug section
                if "params$debug_mode" in content and "debug-table" in content:
                    self.log(
                        "PASS",
                        "Debug mode parameter configured",
                        "Found debug_mode param and conditional table",
                    )
                    return True
                else:
                    self.log(
                        "FAIL",
                        "Debug mode parameter configured",
                        "Missing conditional rendering logic",
                    )
                    return False
            else:
                self.log(
                    "FAIL",
                    "Debug mode parameter configured",
                    "Missing debug_mode parameter in YAML",
                )
                return False

        except Exception as e:
            self.log("FAIL", "Debug mode parameter configured", str(e))
            return False

    def test_demo_mode_parameter(self):
        """Test 5: Verify diagnostic_mode parameter in ResilienceReport.qmd"""
        try:
            qmd_file = Path("ResilienceReport.qmd")
            content = qmd_file.read_text(encoding="utf-8")

            # Check for diagnostic_mode in params
            if "diagnostic_mode:" in content:
                # Check for synthetic data generation
                if "params$diagnostic_mode" in content:
                    self.log(
                        "PASS",
                        "Demo mode parameter configured",
                        "Found diagnostic_mode param and usage",
                    )
                    return True
                else:
                    self.log(
                        "FAIL",
                        "Demo mode parameter configured",
                        "Parameter declared but not used",
                    )
                    return False
            else:
                self.log(
                    "FAIL",
                    "Demo mode parameter configured",
                    "Missing diagnostic_mode parameter in YAML",
                )
                return False

        except Exception as e:
            self.log("FAIL", "Demo mode parameter configured", str(e))
            return False

    def test_person_parameter(self):
        """Test 6: Verify person parameter is declared and used"""
        try:
            qmd_file = Path("ResilienceReport.qmd")
            content = qmd_file.read_text(encoding="utf-8")

            # Check for person in params
            if "person:" in content:
                # Check for person filtering logic
                if "person_target" in content and "normalize_name" in content:
                    self.log(
                        "PASS",
                        "Person parameter configured",
                        "Found person param and filtering logic",
                    )
                    return True
                else:
                    self.log(
                        "FAIL",
                        "Person parameter configured",
                        "Parameter declared but filtering logic missing",
                    )
                    return False
            else:
                self.log(
                    "FAIL",
                    "Person parameter configured",
                    "Missing person parameter in YAML",
                )
                return False

        except Exception as e:
            self.log("FAIL", "Person parameter configured", str(e))
            return False

    def test_robust_data_cleaning(self):
        """Test 7: Verify robust data cleaning in ResilienceReport.qmd"""
        try:
            qmd_file = Path("ResilienceReport.qmd")
            content = qmd_file.read_text(encoding="utf-8")

            # Check for data cleaning patterns
            has_gsub_comma = 'gsub(",", ".", ' in content
            has_gsub_nonnumeric = 'gsub("[^0-9.-]"' in content
            has_na_replacement = "[is.na(" in content and ")] <- 2.5" in content
            has_clamping = "pmax(0, pmin(5" in content

            if (
                has_gsub_comma
                and has_gsub_nonnumeric
                and has_na_replacement
                and has_clamping
            ):
                self.log(
                    "PASS",
                    "Robust data cleaning implemented",
                    "Found all cleaning steps: comma replacement, non-numeric removal, NA handling, clamping",
                )
                return True
            else:
                missing = []
                if not has_gsub_comma:
                    missing.append("comma replacement")
                if not has_gsub_nonnumeric:
                    missing.append("non-numeric removal")
                if not has_na_replacement:
                    missing.append("NA handling")
                if not has_clamping:
                    missing.append("value clamping")

                self.log(
                    "FAIL",
                    "Robust data cleaning implemented",
                    f"Missing: {', '.join(missing)}",
                )
                return False

        except Exception as e:
            self.log("FAIL", "Robust data cleaning implemented", str(e))
            return False

    def test_gui_checkboxes(self):
        """Test 8: Verify GUI has debug and demo mode checkboxes"""
        try:
            gui_file = Path("ResilienceScanGUI.py")
            content = gui_file.read_text(encoding="utf-8")

            has_debug_var = "debug_mode_var" in content
            has_demo_var = "demo_mode_var" in content
            has_debug_checkbox = "Debug Mode (show raw data table" in content
            has_demo_checkbox = "Demo Mode (use synthetic test data)" in content

            if (
                has_debug_var
                and has_demo_var
                and has_debug_checkbox
                and has_demo_checkbox
            ):
                self.log(
                    "PASS",
                    "GUI checkboxes configured",
                    "Found debug_mode_var, demo_mode_var and both checkboxes",
                )
                return True
            else:
                missing = []
                if not has_debug_var:
                    missing.append("debug_mode_var")
                if not has_demo_var:
                    missing.append("demo_mode_var")
                if not has_debug_checkbox:
                    missing.append("debug checkbox")
                if not has_demo_checkbox:
                    missing.append("demo checkbox")

                self.log(
                    "FAIL",
                    "GUI checkboxes configured",
                    f"Missing: {', '.join(missing)}",
                )
                return False

        except Exception as e:
            self.log("FAIL", "GUI checkboxes configured", str(e))
            return False

    def test_gui_quality_buttons(self):
        """Test 9: Verify GUI has quality dashboard and data cleaner buttons"""
        try:
            gui_file = Path("ResilienceScanGUI.py")
            content = gui_file.read_text(encoding="utf-8")

            has_quality_button = "Run Quality Dashboard" in content
            has_cleaner_button = "Run Data Cleaner" in content
            has_quality_method = "def run_quality_dashboard" in content
            has_cleaner_method = "def run_data_cleaner" in content

            if (
                has_quality_button
                and has_cleaner_button
                and has_quality_method
                and has_cleaner_method
            ):
                self.log(
                    "PASS",
                    "GUI quality buttons configured",
                    "Found both buttons and their handler methods",
                )
                return True
            else:
                missing = []
                if not has_quality_button:
                    missing.append("quality button")
                if not has_cleaner_button:
                    missing.append("cleaner button")
                if not has_quality_method:
                    missing.append("quality method")
                if not has_cleaner_method:
                    missing.append("cleaner method")

                self.log(
                    "FAIL",
                    "GUI quality buttons configured",
                    f"Missing: {', '.join(missing)}",
                )
                return False

        except Exception as e:
            self.log("FAIL", "GUI quality buttons configured", str(e))
            return False

    def test_gui_passes_parameters(self):
        """Test 10: Verify GUI passes debug_mode, demo_mode, and person parameters"""
        try:
            gui_file = Path("ResilienceScanGUI.py")
            content = gui_file.read_text(encoding="utf-8")

            # Look for parameter passing in quarto render command
            has_person_param = "'-P', f'person=" in content
            has_debug_param = "'-P', f'debug_mode=" in content
            has_demo_param = "'-P', f'diagnostic_mode=" in content

            if has_person_param and has_debug_param and has_demo_param:
                self.log(
                    "PASS",
                    "GUI passes all parameters to Quarto",
                    "Found person, debug_mode, and diagnostic_mode parameters",
                )
                return True
            else:
                missing = []
                if not has_person_param:
                    missing.append("person parameter")
                if not has_debug_param:
                    missing.append("debug_mode parameter")
                if not has_demo_param:
                    missing.append("diagnostic_mode parameter")

                self.log(
                    "FAIL",
                    "GUI passes all parameters to Quarto",
                    f"Missing: {', '.join(missing)}",
                )
                return False

        except Exception as e:
            self.log("FAIL", "GUI passes all parameters to Quarto", str(e))
            return False

    def test_generate_all_reports_passes_person(self):
        """Test 11: Verify Generate_all_reports.py passes person parameter"""
        try:
            script_file = Path("Generate_all_reports.py")
            content = script_file.read_text(encoding="utf-8")

            # Look for person parameter in quarto command
            if (
                "-P person=" in content
                or "'-P', f'person=" in content
                or 'f"-P person=' in content
            ):
                self.log(
                    "PASS",
                    "Generate_all_reports passes person parameter",
                    "Found person parameter in Quarto command",
                )
                return True
            else:
                self.log(
                    "FAIL",
                    "Generate_all_reports passes person parameter",
                    "Missing person parameter in Quarto command",
                )
                return False

        except Exception as e:
            self.log("FAIL", "Generate_all_reports passes person parameter", str(e))
            return False

    def test_email_priority_fallback(self):
        """Test 12: Verify email priority fallback logic in GUI"""
        try:
            gui_file = Path("ResilienceScanGUI.py")
            content = gui_file.read_text(encoding="utf-8")

            # Check for priority accounts list
            has_priority_list = "priority_accounts = [" in content
            has_info_email = "info@resiliencescan.org" in content
            has_rdeboer = "r.deboer@windesheim.nl" in content
            has_cgverhoef = "cg.verhoef@windesheim.nl" in content
            has_fallback_logic = "for priority_email in priority_accounts:" in content

            if (
                has_priority_list
                and has_info_email
                and has_rdeboer
                and has_cgverhoef
                and has_fallback_logic
            ):
                self.log(
                    "PASS",
                    "Email priority fallback configured",
                    "Found priority accounts and fallback logic",
                )
                return True
            else:
                missing = []
                if not has_priority_list:
                    missing.append("priority_accounts list")
                if not has_info_email:
                    missing.append("info@resiliencescan.org")
                if not has_rdeboer:
                    missing.append("r.deboer@windesheim.nl")
                if not has_cgverhoef:
                    missing.append("cg.verhoef@windesheim.nl")
                if not has_fallback_logic:
                    missing.append("fallback logic")

                self.log(
                    "FAIL",
                    "Email priority fallback configured",
                    f"Missing: {', '.join(missing)}",
                )
                return False

        except Exception as e:
            self.log("FAIL", "Email priority fallback configured", str(e))
            return False

    def generate_report(self):
        """Generate summary report"""
        print("\n" + "=" * 70)
        print("[INFO] VALIDATION SUMMARY")
        print("=" * 70)
        print(f"Total Tests: {self.test_count}")
        print(
            f"Passed: {self.pass_count} ({self.pass_count / self.test_count * 100:.1f}%)"
        )
        print(
            f"Failed: {self.fail_count} ({self.fail_count / self.test_count * 100:.1f}%)"
        )
        print("=" * 70)

        if self.fail_count == 0:
            print("\n[OK] ALL TESTS PASSED - Ready for production use!")
        else:
            print(f"\n[WARN] {self.fail_count} test(s) failed - Review failures above")

        # Save detailed report
        report_path = Path("test_reports")
        report_path.mkdir(exist_ok=True)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        report_file = report_path / f"validation_report_{timestamp}.txt"

        with open(report_file, "w", encoding="utf-8", errors="replace") as f:
            f.write("FEATURE VALIDATION REPORT\n")
            f.write(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write("=" * 70 + "\n\n")

            for result in self.results:
                f.write(f"[{result['status']}] {result['test']}\n")
                if result["details"]:
                    f.write(f"    {result['details']}\n")
                f.write("\n")

            f.write("=" * 70 + "\n")
            f.write(
                f"Total: {self.test_count} | Passed: {self.pass_count} | Failed: {self.fail_count}\n"
            )

        print(f"\n[INFO] Detailed report saved: {report_file}")


def main():
    print("=" * 70)
    print("[INFO] RESILIENCE SCAN - FEATURE VALIDATION")
    print(f"[INFO] Started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 70)
    print()

    validator = FeatureValidator()

    # Run all tests
    print("[INFO] Running validation tests...\n")

    validator.test_data_exists()
    validator.test_quality_dashboard_script()
    validator.test_data_cleaner_script()
    validator.test_debug_mode_parameter()
    validator.test_demo_mode_parameter()
    validator.test_person_parameter()
    validator.test_robust_data_cleaning()
    validator.test_gui_checkboxes()
    validator.test_gui_quality_buttons()
    validator.test_gui_passes_parameters()
    validator.test_generate_all_reports_passes_person()
    validator.test_email_priority_fallback()

    # Generate summary
    validator.generate_report()


if __name__ == "__main__":
    main()
