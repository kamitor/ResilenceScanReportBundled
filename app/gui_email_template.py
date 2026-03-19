"""
EmailTemplateMixin — email template editor and SMTP configuration tab.
"""

import glob
import json
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import messagebox, scrolledtext, ttk

import pandas as pd

from app.app_paths import CONFIG_FILE, _DATA_ROOT

try:
    import yaml
except ImportError:
    yaml = None  # type: ignore[assignment]

try:
    import keyring
except ImportError:
    keyring = None  # type: ignore[assignment]

_KEYRING_SERVICE = "ResilienceScan"


class EmailTemplateMixin:
    """Mixin providing the Email Template tab (editor, SMTP config, preview)."""

    # ------------------------------------------------------------------
    # Tab creation
    # ------------------------------------------------------------------

    def create_email_template_tab(self, parent):
        """Create email template editing tab"""
        # Template editor frame
        editor_frame = ttk.LabelFrame(parent, text="Email Template Editor", padding=10)
        editor_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N), padx=10, pady=10)

        # Subject line
        ttk.Label(editor_frame, text="Subject:").grid(
            row=0, column=0, sticky=tk.W, pady=5
        )
        self.email_subject_var = tk.StringVar(
            value="Your Resilience Scan Report \u2013 {company}"
        )
        ttk.Entry(editor_frame, textvariable=self.email_subject_var, width=60).grid(
            row=0, column=1, sticky=(tk.W, tk.E), padx=10, pady=5
        )

        # Template help
        help_text = "Available placeholders: {name}, {company}, {date}"
        ttk.Label(
            editor_frame, text=help_text, font=("Arial", 8), foreground="gray"
        ).grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=5)

        # Body editor
        ttk.Label(editor_frame, text="Email Body:").grid(
            row=2, column=0, sticky=(tk.W, tk.N), pady=5
        )

        body_scroll = ttk.Scrollbar(editor_frame)
        body_scroll.grid(row=2, column=2, sticky=(tk.N, tk.S), pady=5)

        self.email_body_text = scrolledtext.ScrolledText(
            editor_frame,
            wrap=tk.WORD,
            width=70,
            height=12,
            font=("Arial", 10),
            yscrollcommand=body_scroll.set,
        )
        self.email_body_text.grid(row=2, column=1, sticky=(tk.W, tk.E), padx=10, pady=5)
        body_scroll.config(command=self.email_body_text.yview)

        # Default template
        default_body = (
            "Dear {name},\n\n"
            "Please find attached your resilience scan report for {company}.\n\n"
            "If you have any questions, feel free to reach out.\n\n"
            "Best regards,\n\n"
            "[Your Name]\n"
            "[Your Organization]"
        )
        self.email_body_text.insert("1.0", default_body)

        # Buttons
        btn_frame = ttk.Frame(editor_frame)
        btn_frame.grid(row=3, column=0, columnspan=3, pady=10)

        ttk.Button(
            btn_frame,
            text="\U0001f4be Save Template",
            command=self.save_email_template,
            width=15,
        ).grid(row=0, column=0, padx=5)

        ttk.Button(
            btn_frame,
            text="[RESET] Reset to Default",
            command=self.reset_email_template,
            width=18,
        ).grid(row=0, column=1, padx=5)

        ttk.Button(
            btn_frame,
            text="\U0001f441\ufe0f Preview Email",
            command=self.preview_email,
            width=15,
        ).grid(row=0, column=2, padx=5)

        # SMTP Configuration Section
        smtp_frame = ttk.LabelFrame(
            parent, text="SMTP Server Configuration", padding=10
        )
        smtp_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), padx=10, pady=10)

        # ── Sender profile selector ──────────────────────────────────────
        ttk.Label(smtp_frame, text="Sender Profile:").grid(
            row=0, column=0, sticky=tk.W, pady=5
        )
        profile_row = ttk.Frame(smtp_frame)
        profile_row.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=10, pady=5)

        self._smtp_profiles: list[dict] = []  # in-memory list from config.yml
        self.smtp_profile_var = tk.StringVar()
        self._profile_combo = ttk.Combobox(
            profile_row,
            textvariable=self.smtp_profile_var,
            state="readonly",
            width=30,
        )
        self._profile_combo.pack(side=tk.LEFT)
        self._profile_combo.bind("<<ComboboxSelected>>", self._on_profile_selected)

        ttk.Button(
            profile_row, text="Load", command=self._load_selected_profile, width=7
        ).pack(side=tk.LEFT, padx=(6, 2))
        ttk.Button(
            profile_row, text="Save as…", command=self._save_as_profile, width=9
        ).pack(side=tk.LEFT, padx=2)
        ttk.Button(
            profile_row, text="Delete", command=self._delete_profile, width=7
        ).pack(side=tk.LEFT, padx=2)

        ttk.Separator(smtp_frame, orient=tk.HORIZONTAL).grid(
            row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=6
        )

        # ── Individual SMTP fields ───────────────────────────────────────
        # SMTP Server
        ttk.Label(smtp_frame, text="SMTP Server:").grid(
            row=2, column=0, sticky=tk.W, pady=5
        )
        self.smtp_server_var = tk.StringVar(value="smtp.office365.com")
        ttk.Entry(smtp_frame, textvariable=self.smtp_server_var, width=40).grid(
            row=2, column=1, sticky=(tk.W, tk.E), padx=10, pady=5
        )

        # SMTP Port
        ttk.Label(smtp_frame, text="SMTP Port:").grid(
            row=3, column=0, sticky=tk.W, pady=5
        )
        self.smtp_port_var = tk.StringVar(value="587")
        ttk.Entry(smtp_frame, textvariable=self.smtp_port_var, width=10).grid(
            row=3, column=1, sticky=tk.W, padx=10, pady=5
        )

        # From Email
        ttk.Label(smtp_frame, text="From Email:").grid(
            row=4, column=0, sticky=tk.W, pady=5
        )
        self.smtp_from_var = tk.StringVar(value="info@resiliencescan.org")
        ttk.Entry(smtp_frame, textvariable=self.smtp_from_var, width=40).grid(
            row=4, column=1, sticky=(tk.W, tk.E), padx=10, pady=5
        )

        # SMTP Username
        ttk.Label(smtp_frame, text="SMTP Username:").grid(
            row=5, column=0, sticky=tk.W, pady=5
        )
        self.smtp_username_var = tk.StringVar(value="")
        ttk.Entry(smtp_frame, textvariable=self.smtp_username_var, width=40).grid(
            row=5, column=1, sticky=(tk.W, tk.E), padx=10, pady=5
        )

        # SMTP Password
        ttk.Label(smtp_frame, text="SMTP Password:").grid(
            row=6, column=0, sticky=tk.W, pady=5
        )
        self.smtp_password_var = tk.StringVar(value="")
        ttk.Entry(
            smtp_frame, textvariable=self.smtp_password_var, width=40, show="*"
        ).grid(row=6, column=1, sticky=(tk.W, tk.E), padx=10, pady=5)

        # Help text
        help_text = (
            "Gmail: smtp.gmail.com:587 (use app-specific password)\n"
            "Office365: smtp.office365.com:587\n"
            "Outlook.com: smtp-mail.outlook.com:587"
        )
        ttk.Label(
            smtp_frame, text=help_text, font=("Segoe UI", 8), foreground="gray"
        ).grid(row=7, column=0, columnspan=2, sticky=tk.W, pady=5)

        ttk.Button(
            smtp_frame, text="Save Configuration", command=self.save_config
        ).grid(row=8, column=0, columnspan=2, sticky=tk.W, pady=5)

        smtp_frame.columnconfigure(1, weight=1)

        editor_frame.columnconfigure(1, weight=1)

        # Preview frame
        preview_frame = ttk.LabelFrame(parent, text="Email Preview", padding=10)
        preview_frame.grid(
            row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=10, pady=10
        )

        self.email_preview_text = scrolledtext.ScrolledText(
            preview_frame,
            wrap=tk.WORD,
            width=80,
            height=15,
            font=("Courier", 9),
            state=tk.DISABLED,
        )
        self.email_preview_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        preview_frame.columnconfigure(0, weight=1)
        preview_frame.rowconfigure(0, weight=1)

        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(1, weight=1)

        # Load saved template if exists
        self.load_email_template()

    # ------------------------------------------------------------------
    # Config
    # ------------------------------------------------------------------

    # ------------------------------------------------------------------
    # Profile helpers
    # ------------------------------------------------------------------

    def _profile_keyring_key(self, profile_name: str) -> str:
        """Return the keyring service name for a profile."""
        return f"{_KEYRING_SERVICE}__{profile_name}"

    def _store_profile_password(
        self, profile_name: str, username: str, password: str
    ) -> None:
        if keyring is not None and username and password:
            try:
                keyring.set_password(
                    self._profile_keyring_key(profile_name), username, password
                )
            except Exception as kr_err:
                self.log(
                    f"[WARNING] keyring unavailable: {kr_err} — password not saved"
                )

    def _load_profile_password(self, profile_name: str, username: str) -> str:
        if keyring is not None and username:
            try:
                return (
                    keyring.get_password(
                        self._profile_keyring_key(profile_name), username
                    )
                    or ""
                )
            except Exception:
                pass
        return ""

    def _refresh_profile_combo(self) -> None:
        names = [p["name"] for p in self._smtp_profiles]
        self._profile_combo["values"] = names
        if names and not self.smtp_profile_var.get():
            self.smtp_profile_var.set(names[0])

    def _on_profile_selected(self, _event=None) -> None:
        self._load_selected_profile()

    def _load_selected_profile(self) -> None:
        name = self.smtp_profile_var.get()
        profile = next((p for p in self._smtp_profiles if p["name"] == name), None)
        if profile is None:
            return
        self.smtp_server_var.set(profile.get("server", ""))
        self.smtp_port_var.set(str(profile.get("port", 587)))
        self.smtp_from_var.set(profile.get("from_address", ""))
        username = profile.get("username", "")
        self.smtp_username_var.set(username)
        password = self._load_profile_password(name, username)
        self.smtp_password_var.set(password)
        self.log(f"[INFO] Loaded sender profile: {name}")

    def _save_as_profile(self) -> None:
        from tkinter.simpledialog import askstring

        name = askstring(
            "Save Profile",
            "Profile name:",
            initialvalue=self.smtp_profile_var.get() or "Default",
            parent=self.root,
        )
        if not name:
            return
        try:
            port = int(self.smtp_port_var.get() or 587)
        except ValueError:
            messagebox.showerror("Invalid Port", "SMTP port must be a number.")
            return
        username = self.smtp_username_var.get()
        password = self.smtp_password_var.get()
        profile = {
            "name": name,
            "server": self.smtp_server_var.get(),
            "port": port,
            "from_address": self.smtp_from_var.get(),
            "username": username,
        }
        # Replace existing profile with same name or append
        existing = next(
            (i for i, p in enumerate(self._smtp_profiles) if p["name"] == name), None
        )
        if existing is not None:
            self._smtp_profiles[existing] = profile
        else:
            self._smtp_profiles.append(profile)
        self._store_profile_password(name, username, password)
        self._refresh_profile_combo()
        self.smtp_profile_var.set(name)
        self._write_config()
        messagebox.showinfo("Saved", f"Profile '{name}' saved.")

    def _delete_profile(self) -> None:
        name = self.smtp_profile_var.get()
        if not name:
            return
        if not messagebox.askyesno("Delete Profile", f"Delete profile '{name}'?"):
            return
        self._smtp_profiles = [p for p in self._smtp_profiles if p["name"] != name]
        self._refresh_profile_combo()
        if self._smtp_profiles:
            self.smtp_profile_var.set(self._smtp_profiles[0]["name"])
        else:
            self.smtp_profile_var.set("")
        self._write_config()

    def _write_config(self) -> None:
        """Persist smtp_profiles (and legacy smtp key) to config.yml."""
        if yaml is None:
            return
        try:
            data: dict = {}
            if CONFIG_FILE.exists():
                data = yaml.safe_load(CONFIG_FILE.read_text(encoding="utf-8")) or {}
            data["smtp_profiles"] = self._smtp_profiles
            # Keep legacy smtp key in sync with current active profile fields
            # so that old code that reads smtp: still works.
            data["smtp"] = {
                "server": self.smtp_server_var.get(),
                "port": int(self.smtp_port_var.get() or 587),
                "from_address": self.smtp_from_var.get(),
                "username": self.smtp_username_var.get(),
            }
            CONFIG_FILE.parent.mkdir(parents=True, exist_ok=True)
            CONFIG_FILE.write_text(
                yaml.dump(data, default_flow_style=False, allow_unicode=True),
                encoding="utf-8",
            )
        except Exception as e:
            self.log(f"[WARNING] Could not write config.yml: {e}")

    # ------------------------------------------------------------------
    # Config save / load
    # ------------------------------------------------------------------

    def save_config(self):
        """Save current SMTP fields to config.yml and keyring."""
        if yaml is None:
            messagebox.showerror(
                "Error", "PyYAML is not installed — cannot save configuration."
            )
            return
        try:
            int(self.smtp_port_var.get() or 587)
        except ValueError:
            messagebox.showerror(
                "Invalid Port", "SMTP port must be a number (e.g. 587)."
            )
            return
        username = self.smtp_username_var.get()
        password = self.smtp_password_var.get()
        # Persist password to keyring using the active profile name (or "Default")
        profile_name = self.smtp_profile_var.get() or "Default"
        self._store_profile_password(profile_name, username, password)
        self._write_config()
        messagebox.showinfo("Saved", f"Configuration saved to:\n{CONFIG_FILE}")

    def load_config(self):
        """Load SMTP settings and profiles from config.yml into GUI fields."""
        if yaml is None:
            self.log("[WARNING] PyYAML not installed — cannot load config.yml")
            return
        if not CONFIG_FILE.exists():
            return
        try:
            data = yaml.safe_load(CONFIG_FILE.read_text(encoding="utf-8")) or {}

            # ── Load profiles list ───────────────────────────────────────
            self._smtp_profiles = data.get("smtp_profiles", [])

            # Migrate legacy smtp: block into a "Default" profile if no profiles yet
            smtp = data.get("smtp", {})
            if not self._smtp_profiles and smtp:
                self._smtp_profiles = [
                    {
                        "name": "Default",
                        "server": smtp.get("server", ""),
                        "port": smtp.get("port", 587),
                        "from_address": smtp.get("from_address", ""),
                        "username": smtp.get("username", ""),
                    }
                ]

            self._refresh_profile_combo()

            # ── Populate fields from first profile / legacy smtp block ───
            if self._smtp_profiles:
                first = self._smtp_profiles[0]
                self.smtp_profile_var.set(first["name"])
                self.smtp_server_var.set(first.get("server", ""))
                self.smtp_port_var.set(str(first.get("port", 587)))
                self.smtp_from_var.set(first.get("from_address", ""))
                username = first.get("username", "")
                self.smtp_username_var.set(username)
                password = self._load_profile_password(first["name"], username)
                # Fall back to legacy plaintext and migrate
                if not password and smtp.get("password"):
                    password = smtp["password"]
                    self._store_profile_password(first["name"], username, password)
                    smtp_clean = {k: v for k, v in smtp.items() if k != "password"}
                    data["smtp"] = smtp_clean
                    CONFIG_FILE.write_text(
                        yaml.dump(data, default_flow_style=False, allow_unicode=True),
                        encoding="utf-8",
                    )
                    self.log("[INFO] SMTP password migrated to OS credential store")
                if password:
                    self.smtp_password_var.set(password)
            elif smtp:
                # Legacy path — no profiles, just smtp block
                if smtp.get("server"):
                    self.smtp_server_var.set(smtp["server"])
                if smtp.get("port"):
                    self.smtp_port_var.set(str(smtp["port"]))
                if smtp.get("from_address"):
                    self.smtp_from_var.set(smtp["from_address"])
                username = smtp.get("username", "")
                if username:
                    self.smtp_username_var.set(username)

            # Load Outlook account priority list (empty list = use default account)
            self.outlook_accounts = data.get("outlook_accounts", [])
        except Exception as e:
            self.log(f"[WARNING] Could not load config.yml: {e}")

    # ------------------------------------------------------------------
    # Email template methods
    # ------------------------------------------------------------------

    def save_email_template(self):
        """Save email template to file"""
        try:
            template_data = {
                "subject": self.email_subject_var.get(),
                "body": self.email_body_text.get("1.0", tk.END).strip(),
            }

            template_file = _DATA_ROOT / "email_template.json"
            with open(template_file, "w", encoding="utf-8") as f:
                json.dump(template_data, f, indent=2)

            self.log("[OK] Email template saved")
            messagebox.showinfo("Success", "Email template saved successfully!")

        except Exception as e:
            self.log(f"[ERROR] Error saving template: {e}")
            messagebox.showerror("Error", f"Failed to save template:\n{e}")

    def load_email_template(self):
        """Load email template from file"""
        try:
            template_file = _DATA_ROOT / "email_template.json"
            if template_file.exists():
                with open(template_file, encoding="utf-8") as f:
                    template_data = json.load(f)

                self.email_subject_var.set(
                    template_data.get(
                        "subject", "Your Resilience Scan Report \u2013 {company}"
                    )
                )
                self.email_body_text.delete("1.0", tk.END)
                self.email_body_text.insert("1.0", template_data.get("body", ""))

                self.log("[OK] Email template loaded")
        except Exception as e:
            self.log(f"[WARNING] Could not load template: {e}")

    def reset_email_template(self):
        """Reset to default template"""
        default_subject = "Your Resilience Scan Report \u2013 {company}"
        default_body = (
            "Dear {name},\n\n"
            "Please find attached your resilience scan report for {company}.\n\n"
            "If you have any questions, feel free to reach out.\n\n"
            "Best regards,\n\n"
            "[Your Name]\n"
            "[Your Organization]"
        )

        self.email_subject_var.set(default_subject)
        self.email_body_text.delete("1.0", tk.END)
        self.email_body_text.insert("1.0", default_body)

        self.log("[RESET] Email template reset to default")
        messagebox.showinfo("Reset", "Template reset to default!")

    def preview_email(self):
        """Preview email with sample data"""
        if self.df is None or len(self.df) == 0:
            messagebox.showwarning(
                "No Data", "Please load data first to preview emails."
            )
            return

        # Get first row as sample
        sample_row = self.df.iloc[0]
        sample_company = sample_row.get("company_name", "Example Company")
        sample_name = sample_row.get("name", "John Doe")
        sample_email = sample_row.get("email_address", "john.doe@example.com")

        # Get template
        subject_template = self.email_subject_var.get()
        body_template = self.email_body_text.get("1.0", tk.END).strip()

        # Replace placeholders
        sample_date = datetime.now().strftime("%Y-%m-%d")

        subject = subject_template.format(
            company=sample_company, name=sample_name, date=sample_date
        )

        body = body_template.format(
            company=sample_company, name=sample_name, date=sample_date
        )

        # Find report file
        def safe_display_name(name):
            if pd.isna(name) or name == "":
                return "Unknown"
            name_str = str(name).strip()
            name_str = name_str.replace("/", "-")
            name_str = name_str.replace("\\", "-")
            name_str = name_str.replace(":", "-")
            return name_str

        display_company = safe_display_name(sample_company)
        display_person = safe_display_name(sample_name)

        # Look for report file - try both formats
        pattern_new = (
            f"*ResilienceScanReport ({display_company} - {display_person}).pdf"
        )
        pattern_legacy = f"*ResilienceReport ({display_company} - {display_person}).pdf"
        _out_dir = Path(self.output_folder_var.get())
        matches = glob.glob(str(_out_dir / pattern_new))
        if not matches:
            matches = glob.glob(str(_out_dir / pattern_legacy))

        attachment_info = ""
        if matches:
            attachment_file = Path(matches[0])
            file_size = attachment_file.stat().st_size / (1024 * 1024)  # MB
            attachment_info = (
                f"\n[ATTACH] Attachment: {attachment_file.name} ({file_size:.2f} MB)"
            )
        else:
            attachment_info = (
                f"\n[WARNING] No report found for {display_company} - {display_person}"
            )

        # Build preview
        preview = "=" * 70 + "\n"
        preview += "EMAIL PREVIEW\n"
        preview += "=" * 70 + "\n\n"
        preview += f"To: {sample_email}\n"
        preview += f"Subject: {subject}\n"
        preview += attachment_info + "\n"
        preview += "\n" + "-" * 70 + "\n"
        preview += "MESSAGE BODY:\n"
        preview += "-" * 70 + "\n\n"
        preview += body
        preview += "\n\n" + "=" * 70 + "\n"
        preview += "This is a preview using the first record from your data.\n"
        preview += f"Sample: {sample_company} - {sample_name}\n"
        preview += "=" * 70

        # Display preview
        self.email_preview_text.config(state=tk.NORMAL)
        self.email_preview_text.delete("1.0", tk.END)
        self.email_preview_text.insert("1.0", preview)
        self.email_preview_text.config(state=tk.DISABLED)

        self.log("[PREVIEW] Email preview generated")
