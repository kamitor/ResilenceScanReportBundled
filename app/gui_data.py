"""
DataMixin — data tab, data loading, quality analysis, and data-management methods.
"""

import json
import subprocess
import sys
import threading
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk

import pandas as pd

from app.app_paths import DATA_FILE, ROOT_DIR, _DATA_ROOT


class DataMixin:
    """Mixin providing the Data tab and all data-management methods."""

    # ------------------------------------------------------------------
    # Tab creation
    # ------------------------------------------------------------------

    def create_dashboard_tab(self):
        """Create overview dashboard tab"""
        dashboard = ttk.Frame(self.notebook)
        self.notebook.add(dashboard, text="[INFO] Dashboard")

        # Quick actions
        actions_frame = ttk.LabelFrame(dashboard, text="Quick Actions", padding=10)
        actions_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N), padx=10, pady=10)

        ttk.Button(
            actions_frame,
            text="\U0001f504 Reload Data",
            command=self.load_initial_data,
            width=20,
        ).grid(row=0, column=0, padx=5, pady=5)

        ttk.Button(
            actions_frame,
            text=" Generate All Reports",
            command=self.start_generation_all,
            width=20,
        ).grid(row=0, column=1, padx=5, pady=5)

        ttk.Button(
            actions_frame,
            text="\U0001f4e7 Send All Emails",
            command=self.start_email_all,
            width=20,
        ).grid(row=0, column=2, padx=5, pady=5)

        ttk.Button(
            actions_frame,
            text="\U0001f527 Check System",
            command=self.run_system_check,
            width=20,
        ).grid(row=1, column=0, padx=5, pady=5)

        ttk.Button(
            actions_frame,
            text="\U0001f4e6 Repair R Packages",
            command=self._install_r_packages_now,
            width=22,
        ).grid(row=1, column=1, padx=5, pady=5)

        ttk.Button(
            actions_frame,
            text="\U0001fa9f Install Dependencies (Windows)",
            command=self.install_windows_dependencies,
            width=25,
        ).grid(row=2, column=0, columnspan=2, padx=5, pady=5)

        ttk.Button(
            actions_frame,
            text="\U0001f427 Install Dependencies (Linux)",
            command=self.install_linux_dependencies,
            width=25,
        ).grid(row=2, column=2, columnspan=2, padx=5, pady=5)

        # Statistics overview
        stats_frame = ttk.LabelFrame(dashboard, text="Statistics Overview", padding=10)
        stats_frame.grid(
            row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=10, pady=10
        )

        self.stats_text = scrolledtext.ScrolledText(
            stats_frame, wrap=tk.WORD, width=80, height=20, font=("Courier", 10)
        )
        self.stats_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        stats_frame.columnconfigure(0, weight=1)
        stats_frame.rowconfigure(0, weight=1)

        dashboard.columnconfigure(0, weight=1)
        dashboard.rowconfigure(1, weight=1)

    def create_data_tab(self):
        """Create data viewing and processing tab with analysis features"""
        data_tab = ttk.Frame(self.notebook)
        self.notebook.add(data_tab, text="\U0001f4c1 Data")

        # Top controls
        controls_frame = ttk.Frame(data_tab)
        controls_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=10, pady=10)

        ttk.Label(controls_frame, text="Data File:").grid(row=0, column=0, sticky=tk.W)
        self.data_file_label = ttk.Label(
            controls_frame, text=str(DATA_FILE), font=("Arial", 9)
        )
        self.data_file_label.grid(row=0, column=1, sticky=tk.W, padx=10)

        ttk.Button(controls_frame, text="Browse...", command=self.load_data_file).grid(
            row=0, column=2, padx=5
        )

        ttk.Button(
            controls_frame,
            text="\U0001f4e5 Convert Data",
            command=self.run_convert_data,
        ).grid(row=0, column=3, padx=5)

        ttk.Button(
            controls_frame, text="\U0001f9f9 Clean Data", command=self.run_clean_data
        ).grid(row=0, column=4, padx=5)

        ttk.Button(
            controls_frame,
            text="[INFO] View Cleaning Report",
            command=self.view_cleaning_report,
        ).grid(row=0, column=5, padx=5)

        ttk.Button(
            controls_frame,
            text=" Validate Integrity",
            command=self.run_integrity_validation,
        ).grid(row=0, column=6, padx=5)

        ttk.Button(controls_frame, text="Refresh", command=self.load_initial_data).grid(
            row=0, column=7, padx=5
        )

        controls_frame.columnconfigure(1, weight=1)

        # Search and Filter Frame
        filter_frame = ttk.LabelFrame(data_tab, text="Search & Filter", padding=10)
        filter_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), padx=10, pady=10)

        # Search box
        ttk.Label(filter_frame, text="Search:").grid(
            row=0, column=0, sticky=tk.W, padx=5
        )
        self.data_search_var = tk.StringVar()
        self.data_search_var.trace("w", lambda *args: self.filter_data())
        search_entry = ttk.Entry(
            filter_frame, textvariable=self.data_search_var, width=40
        )
        search_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5)

        ttk.Button(
            filter_frame,
            text="Clear",
            command=lambda: self.data_search_var.set(""),
            width=8,
        ).grid(row=0, column=2, padx=5)

        # Filter options
        ttk.Label(filter_frame, text="Show:").grid(
            row=0, column=3, sticky=tk.W, padx=(20, 5)
        )

        self.show_all_var = tk.BooleanVar(value=True)
        self.show_no_email_var = tk.BooleanVar(value=False)
        self.show_duplicates_var = tk.BooleanVar(value=False)

        ttk.Checkbutton(
            filter_frame,
            text="All",
            variable=self.show_all_var,
            command=self.filter_data,
        ).grid(row=0, column=4, padx=5)
        ttk.Checkbutton(
            filter_frame,
            text="Missing Email",
            variable=self.show_no_email_var,
            command=self.filter_data,
        ).grid(row=0, column=5, padx=5)
        ttk.Checkbutton(
            filter_frame,
            text="Duplicates",
            variable=self.show_duplicates_var,
            command=self.filter_data,
        ).grid(row=0, column=6, padx=5)

        # Column selector
        ttk.Label(filter_frame, text="Columns:").grid(
            row=1, column=0, sticky=tk.W, padx=5, pady=(10, 0)
        )

        columns_btn_frame = ttk.Frame(filter_frame)
        columns_btn_frame.grid(row=1, column=1, columnspan=6, sticky=tk.W, pady=(10, 0))

        ttk.Button(
            columns_btn_frame,
            text="Select Columns...",
            command=self.show_column_selector,
            width=15,
        ).pack(side=tk.LEFT, padx=5)

        ttk.Button(
            columns_btn_frame,
            text="Reset View",
            command=self.reset_column_selection,
            width=12,
        ).pack(side=tk.LEFT, padx=5)

        self.selected_columns_label = ttk.Label(
            columns_btn_frame,
            text="Showing: company_name, name, email_address, submitdate",
            font=("Arial", 8),
            foreground="gray",
        )
        self.selected_columns_label.pack(side=tk.LEFT, padx=10)

        filter_frame.columnconfigure(1, weight=1)

        # Data quality frame
        quality_frame = ttk.LabelFrame(
            data_tab, text="Data Quality Analysis", padding=10
        )
        quality_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), padx=10, pady=10)

        self.quality_text = tk.Text(
            quality_frame, height=4, font=("Courier", 9), wrap=tk.WORD
        )
        self.quality_text.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E))

        ttk.Button(
            quality_frame,
            text="\U0001f50d Run Quality Dashboard",
            command=self.run_quality_dashboard,
        ).grid(row=1, column=0, sticky=tk.W, pady=(5, 0))

        ttk.Button(
            quality_frame,
            text="\U0001f9f9 Run Data Cleaner",
            command=self.run_data_cleaner,
        ).grid(row=1, column=1, sticky=tk.W, pady=(5, 0), padx=(10, 0))

        quality_frame.columnconfigure(0, weight=1)
        quality_frame.columnconfigure(1, weight=1)

        # Data preview
        preview_frame = ttk.LabelFrame(data_tab, text="Data Preview", padding=10)
        preview_frame.grid(
            row=3, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=10, pady=10
        )

        # Scrollbars
        tree_scroll_y = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL)
        tree_scroll_y.grid(row=0, column=1, sticky=(tk.N, tk.S))

        tree_scroll_x = ttk.Scrollbar(preview_frame, orient=tk.HORIZONTAL)
        tree_scroll_x.grid(row=1, column=0, sticky=(tk.W, tk.E))

        # Treeview for data
        self.data_tree = ttk.Treeview(
            preview_frame,
            yscrollcommand=tree_scroll_y.set,
            xscrollcommand=tree_scroll_x.set,
            height=15,
        )
        self.data_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        tree_scroll_y.config(command=self.data_tree.yview)
        tree_scroll_x.config(command=self.data_tree.xview)

        # Bind double-click to show row details
        self.data_tree.bind("<Double-Button-1>", self.show_row_details)

        preview_frame.columnconfigure(0, weight=1)
        preview_frame.rowconfigure(0, weight=1)

        # Data info and actions
        info_frame = ttk.Frame(data_tab)
        info_frame.grid(row=4, column=0, sticky=(tk.W, tk.E), padx=10, pady=10)

        self.data_info_label = ttk.Label(
            info_frame, text="No data loaded", font=("Arial", 9)
        )
        self.data_info_label.pack(side=tk.LEFT)

        ttk.Button(
            info_frame, text="Export Filtered Data", command=self.export_filtered_data
        ).pack(side=tk.RIGHT, padx=5)

        ttk.Button(
            info_frame, text="Find Duplicates", command=self.analyze_duplicates
        ).pack(side=tk.RIGHT, padx=5)

        data_tab.columnconfigure(0, weight=1)
        data_tab.rowconfigure(3, weight=1)

        # Store for filtering
        self.filtered_df = None
        self.visible_columns = ["company_name", "name", "email_address", "submitdate"]

    # ------------------------------------------------------------------
    # Data loading
    # ------------------------------------------------------------------

    def load_initial_data(self):
        """Load data on startup"""
        self.log("Loading data from: " + str(DATA_FILE))
        try:
            if DATA_FILE.exists():
                self.df = pd.read_csv(DATA_FILE, encoding="utf-8")
                self.df.columns = self.df.columns.str.lower().str.strip()

                # Update statistics
                self.stats["total_respondents"] = len(self.df)
                self.stats["total_companies"] = self.df["company_name"].nunique()

                # Import data into email tracker
                self.log("Importing email tracking data...")
                imported, skipped = self.email_tracker.import_from_csv(str(DATA_FILE))
                self.log(f"[OK] Email tracker: {imported} imported, {skipped} skipped")

                # Update email statistics
                email_stats = self.email_tracker.get_statistics()
                self.stats["emails_sent"] = email_stats.get("sent", 0)

                self.update_stats_display()
                self.update_data_preview()
                self.update_stats_text()
                self.update_email_status_display()
                self.analyze_data_quality()

                self.log(
                    f"[OK] Data loaded: {len(self.df)} respondents, {self.stats['total_companies']} companies"
                )
                self.status_label.config(text=f"Data loaded: {len(self.df)} records")
            else:
                self.log("[INFO] No data loaded - cleaned_master.csv not found")
                self.log(
                    "[INFO] First time setup: Use 'Data' tab to import and clean your data"
                )
                self.status_label.config(
                    text="No data loaded - use Data tab to import and clean data"
                )
                # Initialize with empty stats
                self.stats["total_respondents"] = 0
                self.stats["total_companies"] = 0
                self.stats["emails_sent"] = 0
                self.update_stats_display()
        except Exception as e:
            self.log(f"[ERROR] Error loading data: {e}")
            messagebox.showerror("Error", f"Failed to load data:\n{e}")

    def load_data_file(self):
        """Browse and load a data file (xlsx, xls, ods, xml, csv, tsv).

        For any file other than cleaned_master.csv itself, the file is copied
        into the data directory and converted via convert_and_save() before
        loading.
        """
        filename = filedialog.askopenfilename(
            title="Select Data File",
            filetypes=[
                (
                    "All supported formats",
                    "*.xlsx *.xlsm *.xls *.ods *.xml *.json *.jsonl *.csv *.tsv",
                ),
                ("Excel files", "*.xlsx *.xlsm *.xls"),
                ("OpenDocument Spreadsheet", "*.ods"),
                ("XML files", "*.xml"),
                ("JSON files", "*.json *.jsonl"),
                ("CSV / TSV files", "*.csv *.tsv"),
                ("All files", "*.*"),
            ],
            initialdir=_DATA_ROOT / "data",
        )

        if not filename:
            return

        path = Path(filename)
        try:
            import convert_data as _cd
            import shutil as _shutil

            dest_dir = _DATA_ROOT / "data"
            dest_dir.mkdir(parents=True, exist_ok=True)
            dest = dest_dir / path.name

            if path.resolve() == DATA_FILE.resolve():
                # User selected cleaned_master.csv itself — load directly
                csv_path = path
            else:
                if dest.resolve() != path.resolve():
                    _shutil.copy2(str(path), str(dest))
                    self.log(f"[INFO] File copied to data dir: {dest.name}")
                fmt = path.suffix.upper().lstrip(".")
                self.log(f"[INFO] Converting {fmt} \u2192 CSV \u2026")
                ok = _cd.convert_and_save(dest)
                if not ok:
                    messagebox.showerror(
                        "Conversion Failed",
                        f"Could not convert {dest.name}.\nCheck the log for details.",
                    )
                    return
                self.log("[OK] Conversion complete \u2014 loading cleaned_master.csv")
                csv_path = DATA_FILE

            self.df = pd.read_csv(csv_path, encoding="utf-8")
            self.df.columns = self.df.columns.str.lower().str.strip()

            self.data_file_label.config(text=str(csv_path))
            self.stats["total_respondents"] = len(self.df)
            self.stats["total_companies"] = self.df["company_name"].nunique()

            self.update_stats_display()
            self.update_data_preview()
            self.update_stats_text()
            self.analyze_data_quality()

            self.log(f"[OK] Data loaded: {len(self.df)} records from {csv_path.name}")
            messagebox.showinfo(
                "Success", f"Data loaded successfully!\n{len(self.df)} records"
            )
        except Exception as e:
            self.log(f"[ERROR] Error loading file: {e}")
            messagebox.showerror("Error", f"Failed to load file:\n{e}")

    def run_convert_data(self):
        """Run the data conversion script to convert Excel files to CSV format"""
        self.log("[START] Starting data conversion process...")
        self.status_label.config(text="Converting data...")

        try:
            # Import and run the convert_data module
            import convert_data

            # Run the conversion function
            self.log("Looking for Excel files in /data folder...")
            success = convert_data.convert_and_save()

            if success:
                self.log("[OK] Data conversion completed!")

                # Automatically load the converted data
                try:
                    self.df = pd.read_csv(DATA_FILE, encoding="utf-8")
                    self.df.columns = self.df.columns.str.lower().str.strip()
                    self.update_stats_display()
                    self.update_data_preview()
                    self.update_stats_text()
                    self.log(f"[OK] Data automatically loaded: {len(self.df)} records")
                except Exception as e:
                    self.log(f"[WARNING] Could not auto-load data: {e}")

                messagebox.showinfo(
                    "Success",
                    "Excel file converted to CSV!\n\n"
                    "The cleaned_master.csv file has been created/updated.\n"
                    "Email tracking status (reportsent) has been preserved.\n\n"
                    "Next step: Click 'Clean Data' to fix any data quality issues.",
                )
                self.status_label.config(
                    text="Data converted and loaded - run Clean Data next"
                )
            else:
                self.log("[ERROR] Data conversion failed - check logs for details")
                messagebox.showerror(
                    "Conversion Failed",
                    "Data conversion failed.\n\n"
                    "Please ensure:\n"
                    "1. Your Excel file (.xlsx or .xls) is in the 'data' folder\n"
                    "2. The file is not open in another program\n"
                    "3. The file contains valid data\n"
                    "4. Check the logs for more details",
                )
                self.status_label.config(text="Data conversion failed")

        except Exception as e:
            self.log(f"[ERROR] Error during data conversion: {e}")
            messagebox.showerror("Error", f"Failed to run data conversion:\n{e}")
            self.status_label.config(text="Error")

    def run_clean_data(self):
        """Run the enhanced data cleaning script with comprehensive validation"""
        self.log("[START] Starting enhanced data cleaning with validation...")
        self.status_label.config(text="Cleaning data...")

        try:
            # Import and run the clean_data module
            import clean_data

            # Run the cleaning function
            self.log("Loading cleaned_master.csv for enhanced cleaning...")
            success, summary = clean_data.clean_and_fix()

            if success:
                self.log("[OK] Data cleaning completed!")
                self.log(f"Summary: {summary}")

                # Automatically reload the cleaned data
                try:
                    self.df = pd.read_csv(DATA_FILE, encoding="utf-8")
                    self.df.columns = self.df.columns.str.lower().str.strip()
                    self.update_stats_display()
                    self.update_data_preview()
                    self.update_stats_text()
                    self.analyze_data_quality()
                    self.log(
                        f"[OK] Cleaned data automatically reloaded: {len(self.df)} records"
                    )
                except Exception as e:
                    self.log(f"[WARNING] Could not auto-reload data: {e}")

                messagebox.showinfo(
                    "Data Cleaned Successfully",
                    f"Enhanced data cleaning completed!\n\n"
                    f"Results:\n{summary}\n\n"
                    f"Data is now ready for report generation.",
                )
                self.status_label.config(
                    text="Data cleaned and loaded - ready for reports"
                )
            else:
                self.log("[ERROR] Data cleaning failed - check logs for details")
                messagebox.showerror(
                    "Cleaning Failed",
                    f"Data cleaning failed.\n\n"
                    f"Reason: {summary}\n\n"
                    "Please ensure:\n"
                    "1. You have run 'Convert Data' first\n"
                    "2. The cleaned_master.csv file exists\n"
                    "3. The data contains required columns (company_name, name, email)\n"
                    "4. Check the cleaning report for detailed feedback",
                )
                self.status_label.config(text="Data cleaning failed")

        except Exception as e:
            self.log(f"[ERROR] Error during data cleaning: {e}")
            messagebox.showerror("Error", f"Failed to run data cleaning:\n{e}")
            self.status_label.config(text="Error")

    def view_cleaning_report(self):
        """Display the detailed cleaning report in a new window"""
        report_path = _DATA_ROOT / "data" / "cleaning_report.txt"
        log_path = _DATA_ROOT / "data" / "cleaning_validation_log.json"

        if not report_path.exists():
            messagebox.showwarning(
                "No Report Available",
                "No cleaning report found.\n\n"
                "Please run 'Clean Data' first to generate a detailed report.",
            )
            return

        # Create new window for report
        report_window = tk.Toplevel(self.root)
        report_window.title("Data Cleaning Report")
        report_window.geometry("900x700")

        # Create frame with scrollbar
        frame = ttk.Frame(report_window, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)

        # Text widget with scrollbar
        text_widget = tk.Text(frame, wrap=tk.WORD, font=("Courier", 10))
        scrollbar = ttk.Scrollbar(frame, command=text_widget.yview)
        text_widget.config(yscrollcommand=scrollbar.set)

        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Load and display report
        try:
            with open(report_path, "r", encoding="utf-8") as f:
                report_content = f.read()
            text_widget.insert("1.0", report_content)
            text_widget.config(state=tk.DISABLED)

            # Button frame
            button_frame = ttk.Frame(report_window, padding=10)
            button_frame.pack(fill=tk.X)

            # Add button to view detailed validation log
            if log_path.exists():
                ttk.Button(
                    button_frame,
                    text="\U0001f4cb View Detailed Validation Log",
                    command=lambda: self.view_validation_log(log_path),
                ).pack(side=tk.LEFT, padx=5)

            ttk.Button(button_frame, text="Close", command=report_window.destroy).pack(
                side=tk.RIGHT, padx=5
            )

        except Exception as e:
            messagebox.showerror("Error", f"Failed to load cleaning report:\n{e}")
            report_window.destroy()

    def view_validation_log(self, log_path):
        """Display the detailed JSON validation log"""
        log_window = tk.Toplevel(self.root)
        log_window.title("Detailed Validation Log")
        log_window.geometry("1000x800")

        # Create frame with scrollbar
        frame = ttk.Frame(log_window, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)

        # Text widget with scrollbar
        text_widget = tk.Text(frame, wrap=tk.WORD, font=("Courier", 9))
        scrollbar = ttk.Scrollbar(frame, command=text_widget.yview)
        text_widget.config(yscrollcommand=scrollbar.set)

        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Load and display log
        try:
            with open(log_path, "r", encoding="utf-8") as f:
                log_data = json.load(f)

            # Format JSON with pretty printing
            formatted_log = json.dumps(log_data, indent=2)
            text_widget.insert("1.0", formatted_log)
            text_widget.config(state=tk.DISABLED)

            # Button frame
            button_frame = ttk.Frame(log_window, padding=10)
            button_frame.pack(fill=tk.X)

            ttk.Button(button_frame, text="Close", command=log_window.destroy).pack(
                side=tk.RIGHT, padx=5
            )

        except Exception as e:
            messagebox.showerror("Error", f"Failed to load validation log:\n{e}")
            log_window.destroy()

    def run_integrity_validation(self):
        """Run data integrity validator to compare Excel vs CSV"""
        self.log("[START] Starting data integrity validation...")
        self.status_label.config(text="Validating data integrity...")

        try:
            # Import and run the integrity validator
            import validate_data_integrity

            # Run validation with 15 samples
            self.log("Comparing Excel source with cleaned CSV (15 random samples)...")
            success = validate_data_integrity.main(num_samples=15)

            if success:
                self.log("[OK] Data integrity validation completed!")

                report_path = _DATA_ROOT / "data" / "integrity_validation_report.txt"
                json_path = _DATA_ROOT / "data" / "integrity_validation_report.json"

                if json_path.exists():
                    with open(json_path, encoding="utf-8") as f:
                        results = json.load(f)

                    stats = results.get("statistics", {})
                    accuracy = 0
                    if stats.get("samples_validated", 0) > 0:
                        accuracy = (
                            (
                                stats.get("perfect_matches", 0)
                                + stats.get("acceptable_matches", 0)
                            )
                            / stats.get("samples_validated", 1)
                            * 100
                        )

                    # Show summary
                    summary = (
                        f"Data Integrity Validation Complete!\n\n"
                        f"Excel records: {stats.get('total_records_excel', 0)}\n"
                        f"CSV records: {stats.get('total_records_csv', 0)}\n"
                        f"Records removed during cleaning: {stats.get('total_records_excel', 0) - stats.get('total_records_csv', 0)}\n\n"
                        f"Samples validated: {stats.get('samples_validated', 0)}\n"
                        f"Perfect matches: {stats.get('perfect_matches', 0)}\n"
                        f"Acceptable matches: {stats.get('acceptable_matches', 0)}\n"
                        f"Mismatches: {stats.get('mismatches', 0)}\n\n"
                        f"Overall accuracy: {accuracy:.1f}%\n\n"
                    )

                    if accuracy >= 95:
                        summary += "[OK] Data integrity verified!\n[OK] Cleaning process preserves data accurately"
                        messagebox.showinfo("Validation Successful", summary)
                    elif accuracy >= 80:
                        summary += "[WARNING] Minor discrepancies detected\n[INFO] Review detailed report for more information"
                        messagebox.showwarning("Validation - Minor Issues", summary)
                    else:
                        summary += "[ERROR] Significant discrepancies detected\n[INFO] Review detailed report immediately"
                        messagebox.showerror("Validation Failed", summary)

                    # Offer to view detailed report
                    if messagebox.askyesno(
                        "View Report?",
                        "Would you like to view the detailed validation report?",
                    ):
                        self.view_integrity_report(report_path)

                self.status_label.config(text="Integrity validation completed")
            else:
                self.log("[ERROR] Data integrity validation failed - check logs")
                messagebox.showerror(
                    "Validation Failed",
                    "Data integrity validation failed.\n\n"
                    "Please ensure:\n"
                    "1. Excel source file exists in data/ folder\n"
                    "2. cleaned_master.csv has been created\n"
                    "3. Check the console logs for details",
                )
                self.status_label.config(text="Integrity validation failed")

        except Exception as e:
            self.log(f"[ERROR] Error during integrity validation: {e}")
            messagebox.showerror("Error", f"Failed to run integrity validation:\n{e}")
            self.status_label.config(text="Error")

    def view_integrity_report(self, report_path):
        """Display the integrity validation report"""
        report_window = tk.Toplevel(self.root)
        report_window.title("Data Integrity Validation Report")
        report_window.geometry("1000x800")

        # Create frame with scrollbar
        frame = ttk.Frame(report_window, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)

        # Text widget with scrollbar
        text_widget = tk.Text(frame, wrap=tk.WORD, font=("Courier", 10))
        scrollbar = ttk.Scrollbar(frame, command=text_widget.yview)
        text_widget.config(yscrollcommand=scrollbar.set)

        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Load and display report
        try:
            with open(report_path, "r", encoding="utf-8") as f:
                report_content = f.read()
            text_widget.insert("1.0", report_content)
            text_widget.config(state=tk.DISABLED)

            # Button frame
            button_frame = ttk.Frame(report_window, padding=10)
            button_frame.pack(fill=tk.X)

            ttk.Button(button_frame, text="Close", command=report_window.destroy).pack(
                side=tk.RIGHT, padx=5
            )

        except Exception as e:
            messagebox.showerror("Error", f"Failed to load report:\n{e}")
            report_window.destroy()

    def update_data_preview(self):
        """Update data preview treeview with current filter"""
        if self.df is None:
            return

        # Apply current filter
        self.filter_data()

        # Run data quality analysis
        self.analyze_data_quality()

    def update_stats_text(self):
        """Update statistics overview text"""
        if self.df is None:
            return

        self.stats_text.delete("1.0", tk.END)

        stats_info = f"""
\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550
RESILIENCESCAN DATA OVERVIEW
\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550

DATASET STATISTICS:
  Total Respondents:       {len(self.df):>6}
  Unique Companies:        {self.df["company_name"].nunique():>6}

ENGAGEMENT METRICS:
  Companies with 1 resp:   {sum(self.df.groupby("company_name").size() == 1):>6}
  Companies with 2-5:      {sum((self.df.groupby("company_name").size() >= 2) & (self.df.groupby("company_name").size() <= 5)):>6}
  Companies with 6-10:     {sum((self.df.groupby("company_name").size() >= 6) & (self.df.groupby("company_name").size() <= 10)):>6}
  Companies with 10+:      {sum(self.df.groupby("company_name").size() > 10):>6}

TOP 10 MOST ENGAGED COMPANIES:
"""

        top_companies = self.df["company_name"].value_counts().head(10)
        for idx, (company, count) in enumerate(top_companies.items(), 1):
            stats_info += f"  {idx:2}. {company:<40} {count:>3} respondents\n"

        # Count existing reports
        _out_dir = Path(self.output_folder_var.get())
        if _out_dir.exists():
            reports = list(_out_dir.glob("*.pdf"))
            stats_info += f"\n\nREPORTS GENERATED:\n  Total PDF files:         {len(reports):>6}\n"
            stats_info += f"  Output folder:           {_out_dir}\n"

        self.stats_text.insert("1.0", stats_info)

    # ------------------------------------------------------------------
    # Filtering and tree view
    # ------------------------------------------------------------------

    def filter_data(self):
        """Filter data based on search and filter options"""
        if self.df is None:
            return

        df_filtered = self.df.copy()

        # Apply search filter
        search_term = self.data_search_var.get().lower()
        if search_term:
            mask = df_filtered.astype(str).apply(
                lambda row: row.str.lower().str.contains(search_term, na=False).any(),
                axis=1,
            )
            df_filtered = df_filtered[mask]

        # Apply checkbox filters
        if self.show_no_email_var.get() and not self.show_all_var.get():
            # Show only missing emails
            df_filtered = df_filtered[
                df_filtered["email_address"].isna()
                | ~df_filtered["email_address"].str.contains("@", na=False)
            ]
        elif self.show_duplicates_var.get() and not self.show_all_var.get():
            # Show only duplicates
            df_filtered = df_filtered[
                df_filtered.duplicated(
                    subset=["company_name", "name", "email_address"], keep=False
                )
            ]

        self.filtered_df = df_filtered

        # Update treeview
        self.refresh_data_tree()

    def refresh_data_tree(self):
        """Refresh treeview with filtered data"""
        if self.filtered_df is None:
            return

        # Clear existing
        for item in self.data_tree.get_children():
            self.data_tree.delete(item)

        # Setup columns
        display_columns = [
            col for col in self.visible_columns if col in self.filtered_df.columns
        ]

        self.data_tree["columns"] = display_columns
        self.data_tree["show"] = "headings"

        for col in display_columns:
            self.data_tree.heading(
                col,
                text=col.replace("_", " ").title(),
                command=lambda c=col: self.sort_by_column(c),
            )
            self.data_tree.column(col, width=150)

        # Add data (show ALL rows, not just 100)
        for idx, row in self.filtered_df.iterrows():
            values = [str(row.get(col, "")) for col in display_columns]

            # Tag duplicates and missing emails for highlighting
            tags = []
            if pd.isna(row.get("email_address")) or "@" not in str(
                row.get("email_address", "")
            ):
                tags.append("no_email")
            if self.filtered_df.duplicated(
                subset=["company_name", "name", "email_address"], keep=False
            ).iloc[self.filtered_df.index.get_loc(idx)]:
                tags.append("duplicate")

            self.data_tree.insert("", tk.END, values=values, tags=tuple(tags))

        # Configure tag colors
        self.data_tree.tag_configure("no_email", background="#ffebee")  # Light red
        self.data_tree.tag_configure("duplicate", background="#fff3e0")  # Light orange

        # Update info
        total_rows = len(self.filtered_df)
        all_rows = len(self.df) if self.df is not None else 0
        info_text = f"Showing {total_rows} of {all_rows} total records"
        if total_rows < all_rows:
            info_text += " (filtered)"
        self.data_info_label.config(text=info_text)

    def sort_by_column(self, col):
        """Sort data by column"""
        if self.filtered_df is None:
            return

        # Toggle sort order
        if not hasattr(self, "sort_column") or self.sort_column != col:
            self.sort_ascending = True
        else:
            self.sort_ascending = not self.sort_ascending

        self.sort_column = col

        # Sort filtered data
        self.filtered_df = self.filtered_df.sort_values(
            by=col, ascending=self.sort_ascending, na_position="last"
        )

        # Refresh display
        self.refresh_data_tree()

    def show_column_selector(self):
        """Show dialog to select visible columns"""
        if self.df is None:
            messagebox.showwarning("No Data", "Please load data first")
            return

        # Create dialog
        dialog = tk.Toplevel(self.root)
        dialog.title("Select Columns")
        dialog.geometry("400x500")
        dialog.transient(self.root)
        dialog.grab_set()

        # Instructions
        ttk.Label(
            dialog,
            text="Select columns to display in the data view:",
            font=("Arial", 10, "bold"),
        ).pack(padx=10, pady=10)

        # Scrollable frame for checkboxes
        canvas = tk.Canvas(dialog)
        scrollbar = ttk.Scrollbar(dialog, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        # Column checkboxes
        column_vars = {}
        for col in self.df.columns:
            var = tk.BooleanVar(value=col in self.visible_columns)
            column_vars[col] = var
            ttk.Checkbutton(scrollable_frame, text=col, variable=var).pack(
                anchor=tk.W, padx=20, pady=2
            )

        canvas.pack(side="left", fill="both", expand=True, padx=10, pady=10)
        scrollbar.pack(side="right", fill="y")

        # Buttons
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(fill=tk.X, padx=10, pady=10)

        def apply_selection():
            self.visible_columns = [
                col for col, var in column_vars.items() if var.get()
            ]
            if not self.visible_columns:
                messagebox.showwarning(
                    "No Columns", "Please select at least one column"
                )
                return
            self.selected_columns_label.config(
                text=f"Showing: {', '.join(self.visible_columns[:3])}{'...' if len(self.visible_columns) > 3 else ''}"
            )
            self.refresh_data_tree()
            dialog.destroy()

        def select_all():
            for var in column_vars.values():
                var.set(True)

        def select_none():
            for var in column_vars.values():
                var.set(False)

        ttk.Button(btn_frame, text="Select All", command=select_all).pack(
            side=tk.LEFT, padx=5
        )
        ttk.Button(btn_frame, text="Select None", command=select_none).pack(
            side=tk.LEFT, padx=5
        )
        ttk.Button(btn_frame, text="Apply", command=apply_selection).pack(
            side=tk.RIGHT, padx=5
        )
        ttk.Button(btn_frame, text="Cancel", command=dialog.destroy).pack(
            side=tk.RIGHT, padx=5
        )

    def reset_column_selection(self):
        """Reset to default columns"""
        self.visible_columns = ["company_name", "name", "email_address", "submitdate"]
        self.selected_columns_label.config(
            text=f"Showing: {', '.join(self.visible_columns)}"
        )
        self.refresh_data_tree()

    def analyze_duplicates(self):
        """Show detailed duplicate analysis"""
        if self.df is None:
            messagebox.showwarning("No Data", "Please load data first")
            return

        # Find duplicates
        duplicates = self.df[
            self.df.duplicated(
                subset=["company_name", "name", "email_address"], keep=False
            )
        ]

        if len(duplicates) == 0:
            messagebox.showinfo("No Duplicates", "No duplicate records found!")
            return

        # Create dialog
        dialog = tk.Toplevel(self.root)
        dialog.title("Duplicate Records Analysis")
        dialog.geometry("800x600")

        # Summary
        summary = f"Found {len(duplicates)} duplicate records ({len(duplicates) // 2} pairs)\n\n"
        summary += "Duplicates based on: company_name + name + email_address"

        ttk.Label(dialog, text=summary, font=("Arial", 10)).pack(padx=10, pady=10)

        # Treeview for duplicates
        tree_frame = ttk.Frame(dialog)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        tree_scroll = ttk.Scrollbar(tree_frame)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        dup_tree = ttk.Treeview(
            tree_frame,
            columns=("Company", "Name", "Email", "Submit Date"),
            show="headings",
            yscrollcommand=tree_scroll.set,
        )
        tree_scroll.config(command=dup_tree.yview)

        dup_tree.heading("Company", text="Company")
        dup_tree.heading("Name", text="Name")
        dup_tree.heading("Email", text="Email")
        dup_tree.heading("Submit Date", text="Submit Date")

        for col in ("Company", "Name", "Email", "Submit Date"):
            dup_tree.column(col, width=150)

        # Add duplicates
        for _idx, row in duplicates.iterrows():
            dup_tree.insert(
                "",
                tk.END,
                values=(
                    row.get("company_name", ""),
                    row.get("name", ""),
                    row.get("email_address", ""),
                    row.get("submitdate", ""),
                ),
            )

        dup_tree.pack(fill=tk.BOTH, expand=True)

        # Close button
        ttk.Button(dialog, text="Close", command=dialog.destroy).pack(pady=10)

    def export_filtered_data(self):
        """Export currently filtered data to CSV"""
        if self.filtered_df is None or len(self.filtered_df) == 0:
            messagebox.showwarning("No Data", "No data to export")
            return

        # Ask for filename
        filename = filedialog.asksaveasfilename(
            title="Export Filtered Data",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            initialfile=f"filtered_data_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
        )

        if filename:
            try:
                self.filtered_df.to_csv(filename, index=False, encoding="utf-8")
                messagebox.showinfo(
                    "Success",
                    f"Exported {len(self.filtered_df)} records to:\n{filename}",
                )
                self.log(f"[OK] Exported {len(self.filtered_df)} records to {filename}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export data:\n{e}")
                self.log(f"[ERROR] Export failed: {e}")

    def show_row_details(self, event):
        """Show detailed view of selected row on double-click"""
        selection = self.data_tree.selection()
        if not selection or self.filtered_df is None:
            return

        # Get row index
        item = selection[0]
        row_values = self.data_tree.item(item)["values"]

        # Find matching row in dataframe
        for _idx, row in self.filtered_df.iterrows():
            if all(
                str(row.get(col, "")) == str(val)
                for col, val in zip(self.visible_columns, row_values)
            ):
                # Create detail dialog
                dialog = tk.Toplevel(self.root)
                dialog.title("Record Details")
                dialog.geometry("600x700")

                # Scrollable text
                text_frame = ttk.Frame(dialog)
                text_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

                scroll = ttk.Scrollbar(text_frame)
                scroll.pack(side=tk.RIGHT, fill=tk.Y)

                detail_text = tk.Text(
                    text_frame,
                    wrap=tk.WORD,
                    yscrollcommand=scroll.set,
                    font=("Courier", 10),
                )
                detail_text.pack(fill=tk.BOTH, expand=True)
                scroll.config(command=detail_text.yview)

                # Add all fields
                details = "=" * 60 + "\n"
                details += "RECORD DETAILS\n"
                details += "=" * 60 + "\n\n"

                for col in self.df.columns:
                    value = row.get(col, "")
                    details += f"{col}:\n  {value}\n\n"

                detail_text.insert("1.0", details)
                detail_text.config(state=tk.DISABLED)

                # Close button
                ttk.Button(dialog, text="Close", command=dialog.destroy).pack(pady=10)
                break

    # ------------------------------------------------------------------
    # Data quality analysis (v2 — live version)
    # ------------------------------------------------------------------

    def analyze_data_quality(self):
        """Automatically analyze and display basic data quality metrics"""
        if self.df is None or len(self.df) == 0:
            self.quality_text.delete("1.0", tk.END)
            self.quality_text.insert("1.0", "No data loaded.")
            return

        try:
            # Score columns
            score_cols = [
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

            available_score_cols = [col for col in score_cols if col in self.df.columns]

            # Calculate metrics
            total_records = len(self.df)
            total_companies = (
                self.df["company_name"].nunique()
                if "company_name" in self.df.columns
                else 0
            )

            # Missing values
            if available_score_cols:
                missing_count = self.df[available_score_cols].isna().sum().sum()
                total_cells = len(self.df) * len(available_score_cols)
                missing_pct = (
                    (missing_count / total_cells) * 100 if total_cells > 0 else 0
                )
            else:
                missing_count = 0
                missing_pct = 0

            # Email completeness
            has_email = 0
            if "email_address" in self.df.columns:
                has_email = self.df["email_address"].notna().sum()
            email_pct = (has_email / total_records) * 100 if total_records > 0 else 0

            # Out of range values (for score columns)
            out_of_range = 0
            if available_score_cols:
                for col in available_score_cols:
                    numeric_col = pd.to_numeric(self.df[col], errors="coerce")
                    out_of_range += ((numeric_col < 0) | (numeric_col > 5)).sum()

            # Build quality summary
            quality_summary = f"""DATA QUALITY ANALYSIS
Total Records: {total_records} | Companies: {total_companies} | Emails: {has_email} ({email_pct:.1f}%)
Missing Values: {missing_count} ({missing_pct:.1f}%) | Out of Range: {out_of_range}
Quality Status: {"[OK] Good" if missing_pct < 5 and out_of_range == 0 else "[WARNING] Issues detected"}

Click 'Run Quality Dashboard' for detailed analysis with visualizations."""

            self.quality_text.delete("1.0", tk.END)
            self.quality_text.insert("1.0", quality_summary)

        except Exception as e:
            self.quality_text.delete("1.0", tk.END)
            self.quality_text.insert("1.0", f"Error analyzing data: {str(e)}")

    def run_quality_dashboard(self):
        """Run data quality monitoring dashboard"""
        if self.df is None:
            messagebox.showwarning("Warning", "Please load data first")
            return

        self.quality_text.delete("1.0", tk.END)
        self.quality_text.insert("1.0", "Running quality dashboard...\n")
        self.root.update()

        def run_in_thread():
            try:
                result = subprocess.run(
                    [sys.executable, "data_quality_dashboard.py"],
                    cwd=ROOT_DIR,
                    capture_output=True,
                    text=True,
                    timeout=60,
                )

                if result.returncode == 0:
                    self.quality_text.delete("1.0", tk.END)
                    self.quality_text.insert("1.0", result.stdout)

                    # Find and show the generated PNG
                    quality_dir = ROOT_DIR / "data" / "quality_reports"
                    if quality_dir.exists():
                        png_files = sorted(quality_dir.glob("quality_dashboard_*.png"))
                        if png_files:
                            latest_png = png_files[-1]
                            messagebox.showinfo(
                                "Quality Dashboard Complete",
                                f"Dashboard generated successfully!\n\nSaved to:\n{latest_png}\n\nCheck the Data tab for details.",
                            )
                else:
                    self.quality_text.delete("1.0", tk.END)
                    self.quality_text.insert("1.0", f"Error:\n{result.stderr}")

            except Exception as e:
                self.quality_text.delete("1.0", tk.END)
                self.quality_text.insert("1.0", f"Error: {str(e)}")

        threading.Thread(target=run_in_thread, daemon=True).start()

    def run_data_cleaner(self):
        """Run enhanced data cleaner"""
        response = messagebox.askyesno(
            "Run Data Cleaner",
            "This will run the enhanced data cleaner and create a backup.\n\nContinue?",
        )

        if not response:
            return

        self.quality_text.delete("1.0", tk.END)
        self.quality_text.insert("1.0", "Running data cleaner...\n")
        self.root.update()

        def run_in_thread():
            try:
                result = subprocess.run(
                    [sys.executable, "clean_data_enhanced.py"],
                    cwd=ROOT_DIR,
                    capture_output=True,
                    text=True,
                    timeout=120,
                )

                if result.returncode == 0:
                    self.quality_text.delete("1.0", tk.END)
                    self.quality_text.insert("1.0", result.stdout)

                    # Check for replacement log
                    replacement_log = ROOT_DIR / "data" / "value_replacements_log.csv"
                    if replacement_log.exists():
                        messagebox.showinfo(
                            "Data Cleaning Complete",
                            f"Data cleaned successfully!\n\nCheck logs:\n- {ROOT_DIR / 'data' / 'cleaning_report.txt'}\n- {replacement_log}",
                        )
                    else:
                        messagebox.showinfo(
                            "Data Cleaning Complete",
                            "Data cleaned successfully!\nNo invalid values found.",
                        )

                    # Reload data
                    self.load_initial_data()
                else:
                    self.quality_text.delete("1.0", tk.END)
                    self.quality_text.insert("1.0", f"Error:\n{result.stderr}")

            except Exception as e:
                self.quality_text.delete("1.0", tk.END)
                self.quality_text.insert("1.0", f"Error: {str(e)}")

        threading.Thread(target=run_in_thread, daemon=True).start()

    # ------------------------------------------------------------------
    # Utility / stats display
    # ------------------------------------------------------------------

    def update_stats_display(self):
        """Update statistics in header"""
        # Count actual reports in output directory
        _out_dir = Path(self.output_folder_var.get())
        if _out_dir.exists():
            reports = list(_out_dir.glob("*.pdf"))
            self.stats["reports_generated"] = len(reports)

        self.stats_labels["respondents"].config(
            text=str(self.stats["total_respondents"])
        )
        self.stats_labels["companies"].config(text=str(self.stats["total_companies"]))
        self.stats_labels["reports"].config(text=str(self.stats["reports_generated"]))
        self.stats_labels["emails"].config(text=str(self.stats["emails_sent"]))
