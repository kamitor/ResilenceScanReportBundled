"""
EmailSendMixin — email sending tab, send thread, and stop control.
"""

import glob
import smtplib
import threading
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import messagebox, scrolledtext, ttk

import pandas as pd

from app.app_paths import DATA_FILE
from utils.constants import SMTP_TIMEOUT_SECONDS, TEST_MODE_LABEL


def _find_row(df: "pd.DataFrame | None", company: str, person: str):
    """Return the first matching row Series, or None if not found."""
    if df is None:
        return None
    matches = df[
        (df["company_name"].str.strip() == company.strip())
        & (df["name"].str.strip() == person.strip())
    ]
    return matches.iloc[0] if not matches.empty else None


class EmailSendMixin:
    """Mixin providing the Email Sending tab and all email-dispatch methods."""

    # ------------------------------------------------------------------
    # Tab creation
    # ------------------------------------------------------------------

    def create_email_sending_tab(self, parent):
        """Create email sending tab (original email tab content)"""

        # Email Status Section
        status_frame = ttk.LabelFrame(parent, text="Email Status Overview", padding=10)
        status_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N), padx=10, pady=10)

        # Statistics labels
        stats_row = ttk.Frame(status_frame)
        stats_row.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=5)

        self.email_stats_label = ttk.Label(
            stats_row,
            text="Total: 0 | Pending: 0 | Sent: 0 | Failed: 0",
            font=("Segoe UI", 10, "bold"),
        )
        self.email_stats_label.pack(side=tk.LEFT, padx=5)

        # Filter buttons
        filter_frame = ttk.Frame(status_frame)
        filter_frame.grid(row=1, column=0, sticky=tk.W, pady=5)

        ttk.Label(filter_frame, text="Filter:").pack(side=tk.LEFT, padx=5)

        self.email_filter_var = tk.StringVar(value="all")
        ttk.Radiobutton(
            filter_frame,
            text="All",
            variable=self.email_filter_var,
            value="all",
            command=self.update_email_status_display,
        ).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(
            filter_frame,
            text="Pending",
            variable=self.email_filter_var,
            value="pending",
            command=self.update_email_status_display,
        ).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(
            filter_frame,
            text="Sent",
            variable=self.email_filter_var,
            value="sent",
            command=self.update_email_status_display,
        ).pack(side=tk.LEFT, padx=5)

        # Email status treeview
        tree_frame = ttk.Frame(status_frame)
        tree_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)

        tree_scroll = ttk.Scrollbar(tree_frame)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self.email_status_tree = ttk.Treeview(
            tree_frame,
            columns=("Company", "Person", "Email", "Status", "Date", "Mode"),
            show="headings",
            height=8,
            yscrollcommand=tree_scroll.set,
        )
        tree_scroll.config(command=self.email_status_tree.yview)

        self.email_status_tree.heading("Company", text="Company")
        self.email_status_tree.heading("Person", text="Person")
        self.email_status_tree.heading("Email", text="Email")
        self.email_status_tree.heading("Status", text="Status")
        self.email_status_tree.heading("Date", text="Date Sent")
        self.email_status_tree.heading("Mode", text="Mode")

        self.email_status_tree.column("Company", width=150)
        self.email_status_tree.column("Person", width=120)
        self.email_status_tree.column("Email", width=180)
        self.email_status_tree.column("Status", width=80)
        self.email_status_tree.column("Date", width=130)
        self.email_status_tree.column("Mode", width=60)

        self.email_status_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Manual update buttons
        update_btn_frame = ttk.Frame(status_frame)
        update_btn_frame.grid(row=3, column=0, sticky=tk.W, pady=5)

        ttk.Button(
            update_btn_frame,
            text="Mark as Sent",
            command=self.mark_selected_as_sent,
            width=15,
        ).pack(side=tk.LEFT, padx=5)
        ttk.Button(
            update_btn_frame,
            text="Reset to Pending",
            command=self.mark_selected_as_pending,
            width=15,
        ).pack(side=tk.LEFT, padx=5)
        ttk.Button(
            update_btn_frame,
            text="Refresh",
            command=self.update_email_status_display,
            width=12,
        ).pack(side=tk.LEFT, padx=5)

        status_frame.columnconfigure(0, weight=1)
        status_frame.rowconfigure(2, weight=1)

        # Controls
        controls_frame = ttk.LabelFrame(parent, text="Email Controls", padding=10)
        controls_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), padx=10, pady=10)

        # Sender identity (reflects current SMTP profile)
        ttk.Label(controls_frame, text="Sending from:").grid(
            row=0, column=0, sticky=tk.W, pady=(5, 2)
        )
        sender_row = ttk.Frame(controls_frame)
        sender_row.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=10, pady=(5, 2))
        # smtp_from_var is defined in EmailTemplateMixin — display it as readonly
        ttk.Entry(
            sender_row,
            textvariable=self.smtp_from_var,
            state="readonly",
            width=38,
        ).pack(side=tk.LEFT)
        ttk.Label(
            sender_row, text="(change in Email Settings tab)", foreground="gray"
        ).pack(side=tk.LEFT, padx=6)

        ttk.Separator(controls_frame, orient=tk.HORIZONTAL).grid(
            row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=4
        )

        # Test mode
        self.test_mode_var = tk.BooleanVar(value=True)
        test_check = ttk.Checkbutton(
            controls_frame,
            text="Test Mode (send to test email only)",
            variable=self.test_mode_var,
            command=self.toggle_test_mode,
        )
        test_check.grid(row=2, column=0, columnspan=2, sticky=tk.W, pady=5)

        ttk.Label(controls_frame, text="Test Email:").grid(row=3, column=0, sticky=tk.W)
        self.test_email_var = tk.StringVar(value="")
        ttk.Entry(controls_frame, textvariable=self.test_email_var, width=40).grid(
            row=3, column=1, sticky=(tk.W, tk.E), padx=10
        )

        controls_frame.columnconfigure(1, weight=1)

        # Action buttons
        button_frame = ttk.Frame(controls_frame)
        button_frame.grid(row=2, column=0, columnspan=2, pady=10)

        self.email_start_btn = ttk.Button(
            button_frame,
            text="\u25b6 Start Sending",
            command=self.start_email_all,
            width=20,
        )
        self.email_start_btn.grid(row=0, column=0, padx=5)

        self.email_stop_btn = ttk.Button(
            button_frame,
            text="\u23f9 Stop",
            command=self.stop_email,
            state=tk.DISABLED,
            width=15,
        )
        self.email_stop_btn.grid(row=0, column=1, padx=5)

        # Progress
        progress_frame = ttk.LabelFrame(parent, text="Email Progress", padding=10)
        progress_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), padx=10, pady=10)

        self.email_progress_label = ttk.Label(progress_frame, text="Ready")
        self.email_progress_label.grid(row=0, column=0, sticky=tk.W)

        self.email_progress = ttk.Progressbar(
            progress_frame, orient=tk.HORIZONTAL, mode="determinate"
        )
        self.email_progress.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=5)

        self.email_current_label = ttk.Label(
            progress_frame,
            text="No active sending",
            font=("Arial", 9),
            foreground="gray",
        )
        self.email_current_label.grid(row=2, column=0, sticky=tk.W)

        progress_frame.columnconfigure(0, weight=1)

        # Email log
        log_frame = ttk.LabelFrame(parent, text="Email Log", padding=10)
        log_frame.grid(
            row=3, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=10, pady=10
        )

        self.email_log = scrolledtext.ScrolledText(
            log_frame, wrap=tk.WORD, width=80, height=10, font=("Courier", 9)
        )
        self.email_log.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(3, weight=1)

        # Load initial email status display
        self.root.after(100, self.update_email_status_display)

    # ------------------------------------------------------------------
    # Email sending
    # ------------------------------------------------------------------

    def toggle_test_mode(self):
        """Toggle test mode for emails"""
        if self.test_mode_var.get():
            self.log_email(
                "[INFO] Test mode enabled - emails will only go to test address"
            )
        else:
            self.log_email(
                "[WARNING] Test mode disabled - emails will go to real recipients!"
            )

    def start_email_all(self):
        """Start sending all emails with prerequisite checks"""
        if self.df is None:
            messagebox.showwarning("Warning", "Please load data first")
            return

        if self.is_sending_emails:
            messagebox.showwarning("Warning", "Email sending already in progress")
            return

        # ===== PREREQUISITE CHECKS (Issue #51 Fix) =====

        # CHECK 1: SMTP Configuration available?
        # No longer using Outlook COM - using SMTP instead
        smtp_server = getattr(self, "smtp_server_var", None)
        if not smtp_server or not smtp_server.get():
            messagebox.showerror(
                "SMTP Not Configured",
                "Email server (SMTP) is not configured.\n\n"
                "Please configure SMTP settings in the Email tab:\n"
                "- SMTP Server (e.g., smtp.gmail.com)\n"
                "- SMTP Port (e.g., 587 for TLS)\n"
                "- Username (your email address)\n"
                "- Password (app password)\n\n"
                "For Gmail: Use app-specific password\n"
                "For Office365: smtp.office365.com port 587",
            )
            return  # STOP - don't proceed

        # CHECK 2: Any reports exist?
        _out_dir = Path(self.output_folder_var.get())
        pdf_files = list(_out_dir.glob("*.pdf")) if _out_dir.exists() else []
        if len(pdf_files) == 0:
            messagebox.showerror(
                "No Reports Found",
                f"No PDF reports found in:\n{_out_dir}\n\n"
                "Please generate reports first (Generation tab).",
            )
            return  # STOP - don't proceed

        # CHECK 3: Any pending emails?
        stats = self.email_tracker.get_statistics()
        pending = stats.get("pending", 0)
        already_sent = stats.get("sent", 0)

        if pending == 0:
            messagebox.showinfo(
                "No Pending Emails",
                "All emails have already been sent or failed.\n\n"
                "Check email status table for details.",
            )
            return  # STOP - nothing to do

        # CHECK 4: If test mode is enabled, validate test email address
        if self.test_mode_var.get():
            test_email = self.test_email_var.get().strip()
            _parts = test_email.split("@")
            if (
                len(_parts) != 2
                or not _parts[0]
                or not _parts[1]
                or "." not in _parts[1]
            ):
                messagebox.showerror(
                    "Invalid Test Email",
                    "Test mode is enabled but the test email address is invalid.\n\n"
                    "Please enter a valid email address in the 'Test Email' field.",
                )
                return  # STOP - invalid test email

        # ===== ALL CHECKS PASSED - CONFIRM WITH USER =====

        # Warn if test mode is off
        if not self.test_mode_var.get():
            response = messagebox.askyesno(
                "Confirm Live Sending",
                f"[WARNING] TEST MODE IS OFF!\n\n"
                f"Emails will be sent to REAL recipients.\n\n"
                f"Pending: {pending}\n"
                f"Already sent: {already_sent}\n"
                f"Reports available: {len(pdf_files)}\n"
                f"SMTP Server: {self.smtp_server_var.get()}:{self.smtp_port_var.get()}\n\n"
                f"Are you sure?",
            )
            if not response:
                return
        else:
            response = messagebox.askyesno(
                "Confirm Email Sending",
                f"Ready to send {pending} pending emails.\n\n"
                f"Reports available: {len(pdf_files)}\n"
                f"SMTP Server: {self.smtp_server_var.get()}:{self.smtp_port_var.get()}\n"
                f"Test mode: YES\n"
                f"Test email: {self.test_email_var.get().strip()}\n\n"
                f"Already sent: {already_sent}\n\n"
                f"Continue?",
            )
            if not response:
                return

        try:
            smtp_port_val = int(self.smtp_port_var.get() or "587")
        except ValueError:
            messagebox.showerror(
                "Invalid Port", "SMTP port must be a number (e.g. 587)."
            )
            return

        self.is_sending_emails = True
        self.email_start_btn.config(state=tk.DISABLED)
        self.email_stop_btn.config(state=tk.NORMAL)

        # Capture all widget values and the current DataFrame on the main thread
        # before the background thread starts — Tkinter widgets are not thread-safe,
        # and self.df can be replaced by the main thread while the email thread runs.
        send_config = {
            "smtp_server": self.smtp_server_var.get(),
            "smtp_port": smtp_port_val,
            "smtp_username": self.smtp_username_var.get(),
            "smtp_password": self.smtp_password_var.get(),
            "smtp_from": self.smtp_from_var.get(),
            "out_dir": Path(self.output_folder_var.get()),
            "test_mode": self.test_mode_var.get(),
            "test_email": self.test_email_var.get().strip(),
            "subject_template": self.email_subject_var.get(),
            "body_template": self.email_body_text.get("1.0", tk.END).strip(),
            "df": self.df.copy() if self.df is not None else None,
            "outlook_accounts": list(getattr(self, "outlook_accounts", [])),
        }

        # Start email sending in background thread
        thread = threading.Thread(
            target=self.send_emails_thread, args=(send_config,), daemon=True
        )
        thread.start()

    def send_emails_thread(self, send_config):
        """Background thread for sending emails - works directly from PDF reports"""
        # Initialize COM for this thread (Windows only — no-op on Linux/macOS)
        _com_initialized = False
        try:
            import pythoncom

            pythoncom.CoInitialize()
            _com_initialized = True
        except ImportError:
            pass

        try:
            self.log_email("[START] Starting email distribution...")

            # Log test mode status at the start
            if send_config["test_mode"]:
                self.log_email(
                    f"[TEST] TEST MODE ENABLED - All emails will be sent to: {send_config['test_email']}"
                )
            else:
                self.log_email(
                    "[LIVE] LIVE MODE - Emails will be sent to actual recipients"
                )

            self._send_emails_impl(send_config)

        except Exception as exc:
            import traceback

            self.log_email(f"[ERROR] Unexpected error in email thread: {exc}")
            self.log_email(f"  {traceback.format_exc()}")

            def _finalize_error(exc=exc):
                self.is_sending_emails = False
                self.email_start_btn.config(state=tk.NORMAL)
                self.email_stop_btn.config(state=tk.DISABLED)
                messagebox.showerror(
                    "Email Error",
                    f"An unexpected error stopped the email thread:\n\n{exc}\n\n"
                    "Check the Email Log tab for the full traceback.",
                )

            self.root.after(0, _finalize_error)

        finally:
            if _com_initialized:
                import pythoncom

                pythoncom.CoUninitialize()

    def _send_emails_impl(self, send_config):
        """Implementation of email sending - separated for COM initialization"""

        # Use the DataFrame snapshot captured on the main thread (thread-safe).
        df_snap = send_config.get("df")

        # Scan output folder for PDF files (same as display logic)
        _out_dir = send_config["out_dir"]
        report_files = glob.glob(str(_out_dir / "*.pdf"))

        if not report_files:
            self.log_email(f"[ERROR] No PDF reports found in {_out_dir}")

            def finalize_empty():
                self.is_sending_emails = False
                self.email_start_btn.config(state=tk.NORMAL)
                self.email_stop_btn.config(state=tk.DISABLED)
                messagebox.showwarning(
                    "No Reports Found",
                    f"No PDF reports found in:\n{_out_dir}\n\n"
                    "Please generate reports first, then try sending emails.",
                )

            self.root.after(0, finalize_empty)
            return

        # Parse PDF filenames and build list of reports to send
        pending_records = []

        for pdf_path in report_files:
            filename = Path(pdf_path).name

            # Extract company and person from filename
            # Format: YYYYMMDD ResilienceScanReport (COMPANY NAME - Firstname Lastname).pdf
            # Also support legacy format: YYYYMMDD ResilienceReport (COMPANY NAME - Firstname Lastname).pdf
            try:
                content = None
                # Try new format first
                if "ResilienceScanReport (" in filename and ").pdf" in filename:
                    content = filename.split("ResilienceScanReport (")[1].split(
                        ").pdf"
                    )[0]
                # Fallback to legacy format
                elif "ResilienceReport (" in filename and ").pdf" in filename:
                    content = filename.split("ResilienceReport (")[1].split(").pdf")[0]

                if content and " - " in content:
                    company, person = content.rsplit(" - ", 1)

                    # Look up email and sent-status from CSV snapshot in one pass
                    row = _find_row(df_snap, company, person)
                    if row is None:
                        self.log_email(
                            f"[WARN] No CSV record found for {company} – {person}; email will be blank"
                        )
                    email = row.get("email_address", "") if row is not None else ""
                    is_sent = (
                        row.get("reportsent", False)
                        if row is not None and "reportsent" in df_snap.columns
                        else False
                    )

                    # Only add if not sent yet
                    if not is_sent:
                        pending_records.append(
                            {
                                "company": company,
                                "person": person,
                                "email": email,
                                "pdf_path": pdf_path,
                            }
                        )
            except Exception as e:
                self.log_email(f"[WARNING] Could not parse filename: {filename} - {e}")
                continue

        total = len(pending_records)
        self.log_email(f"[INFO] Total reports ready to send: {total}")

        if total == 0:
            self.log_email("[INFO] All reports have already been sent!")

            def finalize_empty():
                self.is_sending_emails = False
                self.email_start_btn.config(state=tk.NORMAL)
                self.email_stop_btn.config(state=tk.DISABLED)
                messagebox.showinfo(
                    "No Pending Emails",
                    "All reports have already been sent.\n\n"
                    "Check the email status table for details.",
                )

            self.root.after(0, finalize_empty)
            return

        sent_count = 0
        failed_count = 0

        # Read send config (all values captured on main thread before this thread started)
        smtp_server = send_config["smtp_server"]
        smtp_port = send_config["smtp_port"]
        smtp_username = send_config["smtp_username"]
        smtp_password = send_config["smtp_password"]
        smtp_from = send_config["smtp_from"]
        subject_template = send_config["subject_template"]
        body_template = send_config["body_template"]
        test_mode = send_config["test_mode"]
        test_email = send_config["test_email"] if test_mode else None

        self.log_email(f"[SMTP] Connecting to SMTP server: {smtp_server}:{smtp_port}")

        # Note: SMTP connection will be created per-email to avoid timeout issues

        for idx, record in enumerate(pending_records):
            if not self.is_sending_emails:
                self.log_email("[STOP] Email sending stopped by user")
                break

            company = record["company"]
            person = record["person"]
            email = record["email"]
            attachment_path = Path(record["pdf_path"])

            # Update current label
            def update_current(company=company, person=person):
                self.email_progress.configure(maximum=total)
                self.email_current_label.config(text=f"Sending: {company} - {person}")

            self.root.after(0, update_current)

            try:
                self.log_email(
                    f"[{idx + 1}/{total}] [SEND] Sending to: {company} - {person}"
                )
                self.log_email(
                    f"  Email address: {email if email else 'NO EMAIL FOUND'}"
                )
                self.log_email(f"  PDF: {attachment_path.name}")

                # Check if email exists - handle NaN/None/empty values
                if (
                    pd.isna(email)
                    or not email
                    or (
                        isinstance(email, str)
                        and (email.strip() == "" or email == "NO EMAIL")
                    )
                ):
                    raise ValueError(f"No email address found for {company} - {person}")

                # Format subject and body with template
                self.log_email("  Formatting email...")
                subject = subject_template.format(
                    company=company,
                    name=person,
                    date=datetime.now().strftime("%Y-%m-%d"),
                )

                body = body_template.format(
                    company=company,
                    name=person,
                    date=datetime.now().strftime("%Y-%m-%d"),
                )

                # Determine recipient
                recipient = test_email if test_mode else email
                if test_mode:
                    self.log_email(
                        f"  [TEST MODE] Sending to: {recipient} (original: {email})"
                    )
                    body = f"[TEST MODE]\nOriginal recipient: {email}\n\n" + body
                else:
                    self.log_email(f"  [LIVE MODE] Sending to: {recipient}")

                # Validate recipient before sending
                if not recipient or "@" not in recipient:
                    raise ValueError(f"Invalid recipient email address: {recipient}")

                # Try Outlook first with priority account fallback, then SMTP

                try:
                    self.log_email("  [OUTLOOK] Attempting to send via Outlook...")
                    import win32com.client

                    # Create Outlook instance
                    outlook = win32com.client.Dispatch("Outlook.Application")

                    # Define account priority list
                    # Priority accounts loaded from config.yml (outlook_accounts key)
                    priority_accounts = send_config.get("outlook_accounts", [])

                    # Get all available accounts
                    available_accounts = []
                    try:
                        for i in range(1, outlook.Session.Accounts.Count + 1):
                            account = outlook.Session.Accounts.Item(i)
                            available_accounts.append((account.SmtpAddress, account))
                        self.log_email(
                            f"  Found {len(available_accounts)} Outlook account(s)"
                        )
                    except Exception as e:
                        self.log_email(
                            f"  [WARNING] Could not enumerate Outlook accounts: {e}"
                        )

                    # Select account based on priority
                    selected_account = None
                    selected_address = None

                    # Try priority accounts first
                    for priority_email in priority_accounts:
                        for smtp_address, account in available_accounts:
                            if smtp_address.lower() == priority_email.lower():
                                selected_account = account
                                selected_address = smtp_address
                                self.log_email(
                                    f"  [OK] Using priority account: {selected_address}"
                                )
                                break
                        if selected_account:
                            break

                    # If no priority account found, use any available account
                    if not selected_account and available_accounts:
                        selected_address, selected_account = available_accounts[0]
                        self.log_email(
                            f"  [INFO] No priority account available, using: {selected_address}"
                        )

                    # Create mail item
                    mail = outlook.CreateItem(0)  # 0 = MailItem

                    # Set email properties
                    self.log_email(f"  Setting recipient to: {recipient}")
                    mail.To = recipient
                    mail.Subject = subject
                    mail.Body = body

                    # Set the sending account if we found one
                    if selected_account:
                        mail.SendUsingAccount = selected_account
                        self.log_email(
                            f"  [OK] Configured to send from: {selected_address}"
                        )
                    else:
                        self.log_email(
                            "  [WARNING] No specific account found, using Outlook default"
                        )

                    # Add attachment
                    self.log_email(f"  Attaching PDF: {attachment_path}...")
                    mail.Attachments.Add(str(attachment_path.absolute()))

                    # VERIFICATION: Log the actual recipient before sending
                    self.log_email("  [OK] Email configured:")
                    self.log_email(f"     To: {mail.To}")
                    self.log_email(
                        f"     From: {selected_address if selected_address else '(default account)'}"
                    )
                    self.log_email(f"     Subject: {mail.Subject[:50]}...")

                    # Send the email
                    self.log_email("  [OUTLOOK] Sending via Outlook...")
                    mail.Send()
                    self.log_email(
                        f"  [OK] Email sent successfully via Outlook from {selected_address if selected_address else 'default account'}!"
                    )

                except Exception as outlook_ex:
                    outlook_error = str(outlook_ex)
                    self.log_email(f"  [WARNING] Outlook failed: {outlook_error}")
                    self.log_email("  [FALLBACK] Attempting SMTP as fallback...")

                    # Fallback to SMTP
                    from email import encoders
                    from email.mime.base import MIMEBase
                    from email.mime.multipart import MIMEMultipart
                    from email.mime.text import MIMEText

                    # Create message
                    self.log_email(
                        "  [SMTP] Creating SMTP message as final fallback..."
                    )
                    self.log_email(f"  Setting recipient to: {recipient}")
                    msg = MIMEMultipart()
                    msg["From"] = smtp_from
                    msg["To"] = recipient
                    msg["Subject"] = subject

                    # Add body
                    msg.attach(MIMEText(body, "plain"))

                    # Add attachment
                    self.log_email("  Attaching PDF...")
                    with open(attachment_path, "rb") as f:
                        part = MIMEBase("application", "octet-stream")
                        part.set_payload(f.read())
                        encoders.encode_base64(part)
                        part.add_header(
                            "Content-Disposition",
                            f"attachment; filename={attachment_path.name}",
                        )
                        msg.attach(part)

                    # Connect and send
                    self.log_email(
                        f"  [SMTP] Connecting to SMTP: {smtp_server}:{smtp_port}..."
                    )
                    server = smtplib.SMTP(
                        smtp_server, smtp_port, timeout=SMTP_TIMEOUT_SECONDS
                    )
                    try:
                        self.log_email("  [SMTP] Starting TLS...")
                        server.starttls()

                        self.log_email(f"  [SMTP] Logging in as: {smtp_username}...")
                        server.login(smtp_username, smtp_password)

                        self.log_email(f"  [SMTP] Sending message from {smtp_from}...")
                        server.send_message(msg)
                    finally:
                        self.log_email("  [SMTP] Closing connection...")
                        try:
                            server.quit()
                        except Exception:
                            server.close()

                    self.log_email(
                        f"  [OK] Email sent successfully via SMTP from {smtp_from} (used as fallback after Outlook failed)!"
                    )
                # Mark as sent in CSV (ONLY if not in test mode)
                if not test_mode:
                    self.mark_as_sent_in_csv(company, person)
                    self.log_email("  Updated CSV: marked as sent")
                else:
                    self.log_email("  Test mode: NOT updating CSV")

                # Always update email tracker
                self.email_tracker.mark_sent(company, person)

                sent_count += 1
                self.log_email("  [OK] SUCCESS: Email sent!")

            except smtplib.SMTPAuthenticationError as e:
                failed_count += 1
                self.log_email(
                    f"  [ERROR] Authentication error — check username/password: {e}"
                )
                self.email_tracker.mark_failed(company, person)
            except smtplib.SMTPException as e:
                failed_count += 1
                self.log_email(f"  [ERROR] SMTP error: {e}")
                self.email_tracker.mark_failed(company, person)
            except OSError as e:
                failed_count += 1
                self.log_email(
                    f"  [ERROR] Network error connecting to SMTP server: {e}"
                )
                self.email_tracker.mark_failed(company, person)
            except Exception as e:
                failed_count += 1
                import traceback

                self.log_email(f"  [ERROR] FAILED ({type(e).__name__}): {e}")
                self.log_email(f"  Full error:\n{traceback.format_exc()}")
                self.email_tracker.mark_failed(company, person)

            # Update progress
            current_idx = idx + 1

            def update_progress(
                current_idx=current_idx,
                sent_count=sent_count,
                failed_count=failed_count,
            ):
                self.email_progress.configure(value=current_idx)
                self.email_progress_label.config(
                    text=f"Progress: {current_idx}/{total} | Sent: {sent_count} | Failed: {failed_count}"
                )

            self.root.after(0, update_progress)

            # Update email status display after every email
            self.root.after(0, self.update_email_status_display)

        # Final updates
        def finalize():
            self.is_sending_emails = False
            self.email_start_btn.config(state=tk.NORMAL)
            self.email_stop_btn.config(state=tk.DISABLED)
            self.email_current_label.config(text="Email distribution complete")

            # Reload data to get latest sent status
            try:
                self.df = pd.read_csv(DATA_FILE, encoding="utf-8")
                self.df.columns = self.df.columns.str.lower().str.strip()
            except Exception as e:
                self.log_email(f"[WARNING] Could not reload CSV: {e}")

            # Update email status display
            self.update_email_status_display()

            # Update statistics in header
            if self.df is not None and "reportsent" in self.df.columns:
                self.stats["emails_sent"] = self.df["reportsent"].sum()
            self.update_stats_display()

            # Show final summary dialog
            test_mode_str = f" {TEST_MODE_LABEL}" if test_mode else ""
            summary_message = (
                f"Email Distribution Complete{test_mode_str}\n\n"
                f"[OK] Successfully sent: {sent_count}\n"
                f"[ERROR] Failed: {failed_count}\n"
                f"[INFO] Total processed: {total}\n\n"
            )

            if test_mode:
                summary_message += (
                    "[WARNING] Test mode was enabled - CSV not updated.\n"
                )
                summary_message += "All emails were sent to test address.\n\n"

            if failed_count > 0:
                summary_message += "Check Email Log tab for error details."
                messagebox.showwarning(
                    "Email Sending Complete (with failures)", summary_message
                )
            else:
                messagebox.showinfo("Email Sending Complete", summary_message)

        self.root.after(0, finalize)

        self.log_email(
            f"\n[OK] Email distribution complete! Sent: {sent_count}, Failed: {failed_count}"
        )

    def stop_email(self):
        """Stop email sending"""
        if messagebox.askyesno("Confirm", "Stop email sending?"):
            self.is_sending_emails = False
