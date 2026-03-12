"""
GenerationMixin — PDF generation tab and all report-generation methods.
"""

import os
import shutil
import subprocess
import sys
import threading
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk

import pandas as pd

from app.app_paths import (
    DATA_FILE,
    DEFAULT_OUTPUT_DIR,
    _DATA_ROOT,
    _check_r_packages_ready,
    _r_library_path,
)
from utils.filename_utils import safe_display_name, safe_filename


class GenerationMixin:
    """Mixin providing the Generation tab and all PDF-generation methods."""

    # ------------------------------------------------------------------
    # Tab creation
    # ------------------------------------------------------------------

    def create_generation_tab(self):
        """Create PDF generation tab"""
        gen_tab = ttk.Frame(self.notebook)
        self.notebook.add(gen_tab, text=" Generation")

        # Controls
        controls_frame = ttk.LabelFrame(gen_tab, text="Generation Controls", padding=10)
        controls_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=10, pady=10)

        # Options
        ttk.Label(controls_frame, text="Template:").grid(row=0, column=0, sticky=tk.W)
        self.template_var = tk.StringVar(value="ResilienceReport.qmd")
        template_combo = ttk.Combobox(
            controls_frame,
            textvariable=self.template_var,
            values=[
                "ResilienceReport.qmd",
                "SCROLReport.qmd",
            ],
            width=45,
        )
        template_combo.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=10)

        ttk.Label(controls_frame, text="Output Folder:").grid(
            row=1, column=0, sticky=tk.W, pady=5
        )
        self.output_folder_var = tk.StringVar(value=str(DEFAULT_OUTPUT_DIR))
        ttk.Entry(controls_frame, textvariable=self.output_folder_var, width=50).grid(
            row=1, column=1, sticky=(tk.W, tk.E), padx=10
        )

        ttk.Button(
            controls_frame, text="Browse...", command=self.browse_output_folder
        ).grid(row=1, column=2)

        # Debug and Demo Mode Checkboxes
        modes_frame = ttk.Frame(controls_frame)
        modes_frame.grid(row=2, column=0, columnspan=3, sticky=tk.W, pady=10)

        self.debug_mode_var = tk.BooleanVar(value=False)
        self.demo_mode_var = tk.BooleanVar(value=False)

        ttk.Checkbutton(
            modes_frame,
            text="Debug Mode (show raw data table at end of report)",
            variable=self.debug_mode_var,
        ).pack(side=tk.LEFT, padx=(0, 20))

        ttk.Checkbutton(
            modes_frame,
            text="Demo Mode (use synthetic test data)",
            variable=self.demo_mode_var,
        ).pack(side=tk.LEFT)

        controls_frame.columnconfigure(1, weight=1)

        # Action buttons
        button_frame = ttk.Frame(controls_frame)
        button_frame.grid(row=3, column=0, columnspan=3, pady=10)

        self.gen_single_btn = ttk.Button(
            button_frame,
            text=" Generate Single",
            command=self.generate_single_report,
            width=20,
        )
        self.gen_single_btn.grid(row=0, column=0, padx=5)

        self.gen_start_btn = ttk.Button(
            button_frame,
            text="\u25b6 Start All",
            command=self.start_generation_all,
            width=20,
        )
        self.gen_start_btn.grid(row=0, column=1, padx=5)

        self.gen_cancel_btn = ttk.Button(
            button_frame,
            text="\u23f9 Cancel",
            command=self.cancel_generation,
            state=tk.DISABLED,
            width=15,
        )
        self.gen_cancel_btn.grid(row=0, column=2, padx=5)

        # Progress
        progress_frame = ttk.LabelFrame(gen_tab, text="Generation Progress", padding=10)
        progress_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), padx=10, pady=10)

        self.gen_progress_label = ttk.Label(progress_frame, text="Ready")
        self.gen_progress_label.grid(row=0, column=0, sticky=tk.W)

        self.gen_progress = ttk.Progressbar(
            progress_frame, orient=tk.HORIZONTAL, mode="determinate"
        )
        self.gen_progress.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=5)

        self.gen_current_label = ttk.Label(
            progress_frame,
            text="No active generation",
            font=("Arial", 9),
            foreground="gray",
        )
        self.gen_current_label.grid(row=2, column=0, sticky=tk.W)

        progress_frame.columnconfigure(0, weight=1)

        # Generation log
        log_frame = ttk.LabelFrame(gen_tab, text="Generation Log", padding=10)
        log_frame.grid(
            row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=10, pady=10
        )

        self.gen_log = scrolledtext.ScrolledText(
            log_frame, wrap=tk.WORD, width=80, height=15, font=("Courier", 9)
        )
        self.gen_log.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        gen_tab.columnconfigure(0, weight=1)
        gen_tab.rowconfigure(2, weight=1)

    # ------------------------------------------------------------------
    # Single-report generation
    # ------------------------------------------------------------------

    def generate_single_report(self):
        """Generate a single report for selected company/person"""
        if self.df is None:
            messagebox.showwarning("Warning", "Please load data first")
            return

        # Create dialog for company/person selection
        dialog = tk.Toplevel(self.root)
        dialog.title("Generate Single Report")
        dialog.geometry("600x450")
        dialog.transient(self.root)
        dialog.grab_set()

        # Company selection
        ttk.Label(dialog, text="Company:").grid(
            row=0, column=0, sticky=tk.W, padx=10, pady=5
        )

        company_var = tk.StringVar()
        company_list = sorted(self.df["company_name"].unique())
        company_combo = ttk.Combobox(
            dialog, textvariable=company_var, values=company_list, width=50
        )
        company_combo.grid(row=0, column=1, padx=10, pady=5, sticky=(tk.W, tk.E))

        # Person selection
        ttk.Label(dialog, text="Person:").grid(
            row=1, column=0, sticky=tk.W, padx=10, pady=5
        )

        person_var = tk.StringVar()
        person_combo = ttk.Combobox(dialog, textvariable=person_var, width=50)
        person_combo.grid(row=1, column=1, padx=10, pady=5, sticky=(tk.W, tk.E))

        # Update person list when company changes
        def update_person_list(event=None):
            company = company_var.get()
            if company:
                persons = self.df[self.df["company_name"] == company]["name"].tolist()
                person_combo["values"] = persons
                if persons:
                    person_combo.current(0)

        company_combo.bind("<<ComboboxSelected>>", update_person_list)

        # Initialize with first company
        if company_list:
            company_combo.current(0)
            update_person_list()

        # Generate button
        def do_generate():
            company = company_var.get()
            person = person_var.get()

            if not company or not person:
                messagebox.showwarning(
                    "Warning", "Please select both company and person"
                )
                return

            dialog.destroy()

            # Find the row
            row_data = self.df[
                (self.df["company_name"] == company) & (self.df["name"] == person)
            ]
            if len(row_data) == 0:
                messagebox.showerror("Error", "Selected record not found in data")
                return

            row = row_data.iloc[0]

            # Validate output folder before closing the dialog
            if not self._validate_output_folder():
                return

            # Validate record
            validation_result = self.validate_record_for_report(row)
            if not validation_result["is_valid"]:
                messagebox.showerror(
                    "Invalid Record",
                    f"Cannot generate report:\n{validation_result['reason']}",
                )
                return

            # Generate the report
            self.generate_single_report_worker(row, company, person)

        ttk.Button(dialog, text="Generate", command=do_generate, width=15).grid(
            row=2, column=0, columnspan=2, pady=20
        )

        dialog.columnconfigure(1, weight=1)

    def generate_single_report_worker(self, row, company, person):
        """Worker function to generate a single report"""
        self.log_gen(f"\n[START] Generating single report for {company} - {person}")

        # Pre-flight: verify R packages before trying quarto render
        r_pkg_err = _check_r_packages_ready()
        if r_pkg_err:
            log_hint = (
                r"C:\ProgramData\ResilienceScan\setup.log"
                if sys.platform == "win32"
                else "~/.local/share/resiliencescan/setup.log"
            )
            self.log_gen(f"[ERROR] R packages not ready: {r_pkg_err}")
            self.log_gen(f"[ERROR] Check setup log: {log_hint}")
            self.root.after(
                0,
                lambda: messagebox.showerror(
                    "R Packages Missing",
                    f"Required R packages are not installed.\n\n{r_pkg_err}\n\n"
                    "The background setup may still be running or may have failed.\n\n"
                    f"Check the setup log:\n{log_hint}\n\n"
                    "Use the System Check button for details.",
                ),
            )
            return

        try:
            safe_company = safe_filename(company)
            safe_person = safe_filename(person)
            display_company = safe_display_name(company)
            display_person = safe_display_name(person)

            # Output filename
            date_str = datetime.now().strftime("%Y%m%d")
            template_name = Path(self.template_var.get()).stem
            if template_name.startswith("Report"):
                report_name = template_name
            else:
                report_name = template_name

            output_filename = (
                f"{date_str} {report_name} ({display_company} - {display_person}).pdf"
            )
            out_dir = Path(self.output_folder_var.get())
            output_file = out_dir / output_filename

            # Check if already exists
            if output_file.exists():
                response = messagebox.askyesnocancel(
                    "File Exists",
                    f"Report already exists:\n{output_filename}\n\nOverwrite?",
                )
                if not response:
                    self.log_gen("[INFO] Generation cancelled - file already exists")
                    return

            # Build quarto command
            # --output must be a bare filename (no path separators) — Quarto
            # 1.6.x rejects any path component in --output.  Use --output-dir
            # to redirect the PDF to the writable out_dir.
            # Use _DATA_ROOT for template path: quarto creates .quarto/ next to
            # the QMD and _internal/ (ROOT_DIR when frozen) is read-only under
            # Program Files.  _sync_template() copied the QMD there at startup.
            selected_template = _DATA_ROOT / self.template_var.get()
            out_dir.mkdir(parents=True, exist_ok=True)
            temp_name = f"temp_{safe_company}_{safe_person}.pdf"
            temp_path = out_dir / temp_name
            cmd = [
                "quarto",
                "render",
                str(selected_template),
                "-P",
                f"company={company}",
                "-P",
                f"person={person}",
                "-P",
                f"debug_mode={str(self.debug_mode_var.get()).lower()}",
                "-P",
                f"diagnostic_mode={str(self.demo_mode_var.get()).lower()}",
                "--to",
                "pdf",
                "--output",
                temp_name,
                "--output-dir",
                str(out_dir),
            ]

            self.log_gen(
                f"[INFO] Rendering PDF with template: {self.template_var.get()}"
            )
            self.status_label.config(text=f"Generating: {company} - {person}")

            # Build env — inject R_LIBS so R finds packages in the bundled
            # r-library/ dir installed by the setup script.
            single_env = os.environ.copy()
            r_lib = _r_library_path()
            if r_lib is not None and r_lib.exists():
                existing = single_env.get("R_LIBS", "")
                single_env["R_LIBS"] = (
                    f"{r_lib}{os.pathsep}{existing}" if existing else str(r_lib)
                )

            # Execute quarto render — cwd=_DATA_ROOT so quarto writes .quarto/
            # there (writable) and R finds data/cleaned_master.csv correctly.
            # The finally block cleans up temp_path if shutil.move() was never reached.
            try:
                result = subprocess.run(
                    cmd,
                    cwd=str(_DATA_ROOT),
                    capture_output=True,
                    text=True,
                    timeout=300,
                    env=single_env,
                )

                if result.returncode == 0:
                    if temp_path.exists():
                        shutil.move(str(temp_path), str(output_file))
                        self.log_gen(f"[OK] Saved: {output_file}")

                        # Validate the generated report
                        try:
                            from validate_single_report import validate_report

                            validation_result = validate_report(
                                pdf_path=str(output_file),
                                csv_path=str(DATA_FILE),
                                company_name=company,
                                person_name=person,
                            )

                            if validation_result["success"]:
                                self.log_gen(
                                    "[OK] Validation passed: All values match CSV"
                                )
                                messagebox.showinfo(
                                    "Success",
                                    f"Report generated and validated!\n\n{output_filename}\n\nAll values match CSV data.",
                                )
                            else:
                                self.log_gen(
                                    f"[WARNING] Validation: {validation_result['message']}"
                                )
                                # Log details
                                for _key, info in validation_result.get(
                                    "details", {}
                                ).items():
                                    if not info["matches"]:
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
                                        self.log_gen(
                                            f"    {info['label']}: Expected={exp}, Actual={act}"
                                        )

                                messagebox.showwarning(
                                    "Report Generated with Warnings",
                                    f"Report generated:\n{output_filename}\n\nBut validation found issues:\n{validation_result['message']}\n\nCheck logs for details.",
                                )
                        except Exception as ve:
                            self.log_gen(f"[INFO] Validation skipped: {ve}")
                            messagebox.showinfo(
                                "Success", f"Report generated!\n\n{output_filename}"
                            )

                        self.status_label.config(text="Report generated successfully")
                    else:
                        self.log_gen("[ERROR] Output file not found after rendering")
                        messagebox.showerror(
                            "Error", "Report generation failed: Output file not found"
                        )
                        self.status_label.config(text="Error")
                else:
                    self.log_gen(
                        f"[ERROR] Quarto render failed with exit code {result.returncode}"
                    )
                    if result.stderr:
                        self.log_gen(f"stderr: {result.stderr[:2000]}")
                    messagebox.showerror(
                        "Generation Failed",
                        "Report generation failed.\n\nCheck logs for details.",
                    )
                    self.status_label.config(text="Error")

            finally:
                temp_path.unlink(missing_ok=True)

        except FileNotFoundError:
            self.log_gen(
                "[ERROR] Quarto not found - please install from https://quarto.org"
            )
            messagebox.showerror(
                "Quarto Not Found",
                "Quarto is not installed.\n\nPlease install from https://quarto.org",
            )
            self.status_label.config(text="Error")
        except subprocess.TimeoutExpired:
            self.log_gen("[ERROR] Generation timeout (>5 minutes)")
            messagebox.showerror("Timeout", "Report generation timed out (>5 minutes)")
            self.status_label.config(text="Error")
        except Exception as e:
            self.log_gen(f"[ERROR] Error: {e}")
            messagebox.showerror("Error", f"Report generation failed:\n{e}")
            self.status_label.config(text="Error")

    # ------------------------------------------------------------------
    # Batch generation
    # ------------------------------------------------------------------

    def start_generation_all(self):
        """Start generating all reports"""
        if self.df is None:
            messagebox.showwarning("Warning", "Please load data first")
            return

        if self.is_generating:
            messagebox.showwarning("Warning", "Generation already in progress")
            return

        if not self._validate_output_folder():
            return

        # Confirm
        response = messagebox.askyesno(
            "Confirm Generation",
            f"Generate reports for all {len(self.df)} respondents?\n\nThis may take several hours.",
        )

        if not response:
            return

        self.is_generating = True
        self._stop_gen.clear()
        self.gen_start_btn.config(state=tk.DISABLED)
        self.gen_cancel_btn.config(state=tk.NORMAL)

        # Start generation in background thread
        thread = threading.Thread(target=self.generate_reports_thread, daemon=True)
        thread.start()

    def validate_record_for_report(self, row):
        """
        Validate if a record has sufficient data to generate a report.
        Returns dict with 'is_valid' (bool) and 'reason' (str).
        Uses same logic as clean_data_enhanced.py
        """
        # Check company name
        company = row.get("company_name")
        if pd.isna(company) or str(company).strip() in ["", "-", "Unknown"]:
            return {"is_valid": False, "reason": "No valid company name"}

        # Check person name
        person = row.get("name")
        if pd.isna(person) or str(person).strip() == "":
            return {"is_valid": False, "reason": "No person name"}

        # Check email
        email = row.get("email_address")
        if pd.isna(email) or "@" not in str(email):
            return {"is_valid": False, "reason": "Invalid/missing email"}

        # Check score availability - need at least 5 valid scores out of 15
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

        available_scores = 0
        for col in score_columns:
            if col in row.index:
                val = row[col]
                if pd.notna(val) and val not in ["?", "", " "]:
                    try:
                        float_val = float(str(val).replace(",", "."))
                        if 0 <= float_val <= 5:
                            available_scores += 1
                    except Exception:
                        pass

        min_scores_required = 5
        if available_scores < min_scores_required:
            return {
                "is_valid": False,
                "reason": f"Insufficient data ({available_scores}/15 scores, need {min_scores_required})",
            }

        # All checks passed
        return {"is_valid": True, "reason": "Valid"}

    def generate_reports_thread(self):
        """Background thread for report generation"""
        self.log_gen("[START] Starting batch report generation...")
        self.log_gen(f"[INFO] Output folder: {self.output_folder_var.get()}")

        # Pre-flight: verify R packages are available before wasting time on 519 renders
        r_pkg_err = _check_r_packages_ready()
        if r_pkg_err:
            log_hint = (
                r"C:\ProgramData\ResilienceScan\setup.log"
                if sys.platform == "win32"
                else "~/.local/share/resiliencescan/setup.log"
            )
            self.log_gen("[ERROR] R packages not ready \u2014 aborting batch.")
            self.log_gen(f"[ERROR] {r_pkg_err}")
            self.log_gen(f"[ERROR] Check setup log: {log_hint}")
            self.root.after(
                0,
                lambda: messagebox.showerror(
                    "R Packages Missing",
                    f"Required R packages are not installed.\n\n{r_pkg_err}\n\n"
                    "The background setup may still be running or may have failed.\n\n"
                    f"Check the setup log:\n{log_hint}\n\n"
                    "Use the System Check button for details.",
                ),
            )

            def _reset_ui():
                self.is_generating = False
                self.gen_start_btn.config(state=tk.NORMAL)
                self.gen_cancel_btn.config(state=tk.DISABLED)
                self.gen_current_label.config(text="Aborted \u2014 R packages missing")

            self.root.after(0, _reset_ui)
            return

        total = len(self.df)
        success = 0
        failed = 0
        skipped = 0

        self.root.after(
            0, lambda t=total: self.gen_progress.configure(maximum=t, value=0)
        )

        for idx, row in self.df.iterrows():
            try:
                if self._stop_gen.is_set():
                    self.log_gen("Generation cancelled by user")
                    break

                company = row.get("company_name", "Unknown")
                person = row.get("name", "Unknown")

                # Update label via main thread (Tkinter is not thread-safe)
                try:
                    display_text = f"Generating: {company} - {person}"
                except (UnicodeDecodeError, UnicodeEncodeError):
                    safe_c = company.encode("ascii", "replace").decode("ascii")
                    safe_p = person.encode("ascii", "replace").decode("ascii")
                    display_text = f"Generating: {safe_c} - {safe_p}"
                self.root.after(
                    0, lambda t=display_text: self.gen_current_label.config(text=t)
                )
                # Pre-generation validation: Check if record has sufficient data
                validation_result = self.validate_record_for_report(row)

                if not validation_result["is_valid"]:
                    self.log_gen(
                        f"[{idx + 1}/{total}] [SKIP] Skipping {company}: {validation_result['reason']}"
                    )
                    skipped += 1
                    continue

                self.log_gen(f"[{idx + 1}/{total}] Generating: {company} - {person}")

                safe_company = safe_filename(company)
                safe_person = safe_filename(person)
                display_company = safe_display_name(company)
                display_person = safe_display_name(person)

                # Output filename with template name
                date_str = datetime.now().strftime("%Y%m%d")

                # Extract report name from template path
                template_name = Path(
                    self.template_var.get()
                ).stem  # Gets filename without extension
                if template_name.startswith("Report"):
                    # For report variations, use the full name (e.g., "Report1_CircularBarplot")
                    report_name = template_name
                else:
                    # For standard reports, use the template name as is
                    report_name = template_name

                output_filename = f"{date_str} {report_name} ({display_company} - {display_person}).pdf"
                out_dir = Path(self.output_folder_var.get())
                output_file = out_dir / output_filename

                # Check if already exists
                if output_file.exists():
                    self.log_gen("  [SKIP] Already exists, skipping")
                    success += 1
                    continue

                # Build quarto command using selected template with both company and person
                # --output must be a bare filename (no path separators) — Quarto
                # 1.6.x rejects any path component in --output.  Use --output-dir
                # to redirect the PDF to the user-selected out_dir.
                selected_template = _DATA_ROOT / self.template_var.get()
                out_dir.mkdir(parents=True, exist_ok=True)
                temp_name = f"temp_{safe_company}_{safe_person}.pdf"
                temp_path = out_dir / temp_name
                cmd = [
                    "quarto",
                    "render",
                    str(selected_template),
                    "-P",
                    f"company={company}",
                    "-P",
                    f"person={person}",
                    "-P",
                    f"debug_mode={str(self.debug_mode_var.get()).lower()}",
                    "-P",
                    f"diagnostic_mode={str(self.demo_mode_var.get()).lower()}",
                    "--to",
                    "pdf",
                    "--output",
                    temp_name,
                    "--output-dir",
                    str(out_dir),
                ]

                # Build subprocess environment — inject R_LIBS if frozen so
                # Quarto finds the R packages bundled by the installer.
                gen_env = os.environ.copy()
                r_lib = _r_library_path()
                if r_lib is not None and r_lib.exists():
                    existing = gen_env.get("R_LIBS", "")
                    gen_env["R_LIBS"] = (
                        f"{r_lib}{os.pathsep}{existing}" if existing else str(r_lib)
                    )

                # Execute quarto render — cwd=_DATA_ROOT so quarto writes
                # .quarto/ there (writable) and R finds data/ correctly.
                proc = subprocess.Popen(
                    cmd,
                    cwd=str(_DATA_ROOT),
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    text=True,
                    env=gen_env,
                )
                with self._gen_proc_lock:
                    self._gen_proc = proc

                stdout_lines = []
                for line in proc.stdout:
                    line = line.rstrip()
                    if line:
                        self.log_gen(f"    {line}")
                        stdout_lines.append(line)
                    # Check for cancel between lines
                    if self._stop_gen.is_set():
                        try:
                            proc.kill()
                            proc.wait()
                        except (OSError, AttributeError):
                            pass
                        if temp_path.exists():
                            temp_path.unlink()
                        with self._gen_proc_lock:
                            self._gen_proc = None
                        break
                else:
                    proc.wait()
                returncode = proc.returncode
                with self._gen_proc_lock:
                    self._gen_proc = None

                # If cancelled mid-render, break outer loop
                if self._stop_gen.is_set():
                    self.log_gen("Generation cancelled by user")
                    break

                if returncode == 0:
                    if temp_path.exists():
                        shutil.move(str(temp_path), str(output_file))
                        self.log_gen(f"  [OK] Saved: {output_file}")

                        # Validate the generated report
                        try:
                            from validate_single_report import validate_report

                            validation_result = validate_report(
                                pdf_path=str(output_file),
                                csv_path=str(DATA_FILE),
                                company_name=company,
                                person_name=person,
                            )

                            if validation_result["success"]:
                                self.log_gen(
                                    "  [OK] Validation passed: All values match CSV"
                                )
                            else:
                                self.log_gen(
                                    f"  [WARNING] Validation: {validation_result['message']}"
                                )
                                # Log details
                                for _key, info in validation_result.get(
                                    "details", {}
                                ).items():
                                    if not info["matches"]:
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
                                        self.log_gen(
                                            f"      {info['label']}: Expected={exp}, Actual={act}"
                                        )
                        except Exception as ve:
                            self.log_gen(f"  [INFO] Validation skipped: {ve}")

                        success += 1
                    else:
                        self.log_gen(
                            "  [ERROR] Error: Output file not found after render"
                        )
                        failed += 1
                else:
                    if temp_path.exists():
                        temp_path.unlink()
                    self.log_gen(
                        f"  [ERROR] Error: Exit code {returncode} (output logged above)"
                    )
                    failed += 1

            except FileNotFoundError:
                failed += 1
                self.log_gen(
                    "  [ERROR] Error: Quarto not found - please install from https://quarto.org"
                )
            except subprocess.TimeoutExpired:
                failed += 1
                self.log_gen("  [ERROR] Error: Generation timeout (>5 minutes)")
            except Exception as e:
                failed += 1
                self.log_gen(f"  [ERROR] Error: {e}")
                _i, _s, _f, _sk = idx + 1, success, failed, skipped
                self.root.after(
                    0,
                    lambda i=_i, t=total, s=_s, f=_f, sk=_sk: (
                        self.gen_progress.configure(value=i),
                        self.gen_progress_label.config(
                            text=f"Progress: {i}/{t} | Success: {s} | Failed: {f} | Skipped: {sk}"
                        ),
                    ),
                )

        def _reset_gen():
            self.is_generating = False
            self.gen_start_btn.config(state=tk.NORMAL)
            self.gen_cancel_btn.config(state=tk.DISABLED)
            self.gen_current_label.config(text="Generation complete")

        self.root.after(0, _reset_gen)

        # Comprehensive summary
        self.log_gen("\n" + "=" * 60)

        # Check if generation completed all records
        processed = success + failed + skipped
        if processed < total:
            self.log_gen("WARNING: GENERATION INCOMPLETE")
            self.log_gen("=" * 60)
            self.log_gen(f"Total records: {total}")
            self.log_gen(f"Processed: {processed}/{total}")
            self.log_gen(f"Successfully generated: {success}")
            self.log_gen(f"Failed: {failed}")
            self.log_gen(f"Skipped (insufficient data): {skipped}")
            self.log_gen(f"NOT PROCESSED: {total - processed}")
            self.log_gen("=" * 60)
            self.log_gen(f"\nCRITICAL: Generation stopped early at record {processed}.")
            self.log_gen("Check error messages above for details.")
        else:
            self.log_gen("GENERATION COMPLETE")
            self.log_gen("=" * 60)
            self.log_gen(f"Total records: {total}")
            self.log_gen(f"Successfully generated: {success}")
            self.log_gen(f"Failed: {failed}")
            self.log_gen(f"Skipped (insufficient data): {skipped}")
            self.log_gen("=" * 60)

        if skipped > 0:
            self.log_gen(
                f"\nNote: {skipped} record(s) were skipped due to insufficient data."
            )
            self.log_gen(
                "   These records don't have enough scores to generate a valid report."
            )
            self.log_gen(
                "   Run 'Clean Data' to see details about removed/insufficient records."
            )

    def cancel_generation(self):
        """Cancel generation and kill any running quarto subprocess."""
        if messagebox.askyesno("Confirm", "Cancel report generation?"):
            self._stop_gen.set()  # signals the generation thread to stop
            with self._gen_proc_lock:
                proc = self._gen_proc
            if proc is not None:
                try:
                    proc.kill()
                    proc.wait()
                except (OSError, AttributeError):
                    pass

    def _validate_output_folder(self) -> bool:
        """Check that the output folder exists (or can be created) and is writable.

        Shows an error dialog and returns False if the check fails.
        """
        folder_path = Path(self.output_folder_var.get())
        try:
            folder_path.mkdir(parents=True, exist_ok=True)
            probe = folder_path / ".write_test"
            probe.write_text("", encoding="utf-8")
            probe.unlink()
            return True
        except OSError as e:
            messagebox.showerror(
                "Output Folder Not Writable",
                f"Cannot write to the output folder:\n\n{folder_path}\n\n{e}\n\n"
                "Please choose a different folder using the Browse button.",
            )
            return False

    def browse_output_folder(self):
        """Browse for output folder and validate it is writable."""
        folder = filedialog.askdirectory(
            title="Select Output Folder", initialdir=self.output_folder_var.get()
        )
        if not folder:
            return
        self.output_folder_var.set(folder)
        self._validate_output_folder()
