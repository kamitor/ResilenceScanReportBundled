#!/usr/bin/env python3
"""
ResilienceScan Control Center
A graphical interface for managing ResilienceScan reports and email distribution

Features:
- Data processing and validation
- PDF generation with real-time monitoring
- Email distribution management
- Log viewing and status tracking
"""

import queue
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import ttk

# Ensure repo root is on sys.path so sibling modules can be imported
# (needed when running as `python app/main.py` from the repo root)
_repo_root = Path(__file__).resolve().parents[1]
if str(_repo_root) not in sys.path:
    sys.path.insert(0, str(_repo_root))

from app.gui_data import DataMixin  # noqa: E402
from app.gui_email import EmailMixin  # noqa: E402
from app.gui_generate import GenerationMixin  # noqa: E402
from app.gui_logs import LogsMixin  # noqa: E402
from app.gui_settings import SettingsMixin  # noqa: E402
from email_tracker import EmailTracker  # noqa: E402


class ResilienceScanGUI(
    DataMixin, GenerationMixin, EmailMixin, SettingsMixin, LogsMixin
):
    """Main GUI Application for ResilienceScan Control Center"""

    def __init__(self, root):
        self.root = root
        from update_checker import _current_version

        _APP_VERSION = _current_version()
        self.root.title(f"ResilienceScan Control Center  v{_APP_VERSION}")
        self.root.geometry("1200x800")
        self.root.minsize(1000, 600)

        # Data storage
        self.df = None
        self.generation_queue = queue.Queue()
        self.email_queue = queue.Queue()
        self.is_generating = False
        self.is_sending_emails = False
        self._gen_proc = None  # running quarto subprocess (for cancel/kill)
        self._gen_proc_lock = threading.Lock()  # guards _gen_proc across threads
        self._stop_gen = threading.Event()  # set() to request cancellation

        # Email tracking system
        self.email_tracker = EmailTracker()

        # Statistics
        self.stats = {
            "total_companies": 0,
            "total_respondents": 0,
            "reports_generated": 0,
            "emails_sent": 0,
            "errors": 0,
        }

        # Setup GUI
        self.setup_ui()
        self.load_config()
        self._startup_guard()
        self.load_initial_data()

        # Check for updates in the background — non-blocking, fails silently
        try:
            from update_checker import start_background_check

            start_background_check(self._on_update_available, tk_root=self.root)
        except Exception:
            pass

    def setup_ui(self):
        """Create the main UI layout"""

        # Menu bar
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Load Data File...", command=self.load_data_file)
        file_menu.add_command(label="Reload Data", command=self.load_initial_data)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)

        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="About", command=self.show_about)

        # Main container
        main_container = ttk.Frame(self.root, padding="10")
        main_container.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_container.columnconfigure(0, weight=1)
        main_container.rowconfigure(1, weight=1)

        # Header
        self.create_header(main_container)

        # Tab control
        self.notebook = ttk.Notebook(main_container)
        self.notebook.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)

        # Create tabs
        self.create_dashboard_tab()
        self.create_data_tab()
        self.create_generation_tab()
        self.create_email_tab()
        self.create_logs_tab()

        # Status bar
        self.create_status_bar(main_container)

    def create_header(self, parent):
        """Create application header"""
        header_frame = ttk.Frame(parent)
        header_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))

        # Logo/Title
        title_label = ttk.Label(
            header_frame,
            text=" ResilienceScan Control Center",
            font=("Arial", 20, "bold"),
        )
        title_label.grid(row=0, column=0, sticky=tk.W)

        from update_checker import _current_version

        subtitle_label = ttk.Label(
            header_frame,
            text=f"Supply Chain Resilience Assessment Management System  \u2022  v{_current_version()}",
            font=("Arial", 10),
        )
        subtitle_label.grid(row=1, column=0, sticky=tk.W)

        # Quick stats
        stats_frame = ttk.Frame(header_frame)
        stats_frame.grid(row=0, column=1, rowspan=2, sticky=tk.E, padx=20)

        self.stats_labels = {}
        stats_items = [
            ("respondents", "Respondents", "0"),
            ("companies", "Companies", "0"),
            ("reports", "Reports", "0"),
            ("emails", "Emails", "0"),
        ]

        for idx, (key, label, value) in enumerate(stats_items):
            frame = ttk.Frame(stats_frame)
            frame.grid(row=0, column=idx, padx=10)

            ttk.Label(frame, text=label, font=("Arial", 8)).pack()
            self.stats_labels[key] = ttk.Label(
                frame, text=value, font=("Arial", 14, "bold"), foreground="#0277BD"
            )
            self.stats_labels[key].pack()

        header_frame.columnconfigure(1, weight=1)

    def create_status_bar(self, parent):
        """Create status bar at bottom"""
        status_frame = ttk.Frame(parent, relief=tk.SUNKEN)
        status_frame.grid(row=2, column=0, sticky=(tk.W, tk.E))

        self.status_label = ttk.Label(status_frame, text="Ready", font=("Arial", 9))
        self.status_label.grid(row=0, column=0, sticky=tk.W, padx=5)

        # Update notification label — hidden until an update is found
        self._update_label = tk.Label(
            status_frame,
            text="",
            font=("Arial", 9),
            fg="#0066cc",
            cursor="hand2",
        )
        self._update_label.grid(row=0, column=1, sticky=tk.W, padx=10)

        self.status_time_label = ttk.Label(status_frame, text="", font=("Arial", 9))
        self.status_time_label.grid(row=0, column=2, sticky=tk.E, padx=5)

        status_frame.columnconfigure(2, weight=1)

        # Update time every second
        self.update_time()

    def _on_update_available(self, info):
        """Called (on the main thread) when the update check completes."""
        if not info:
            return
        version = info.get("version", "")
        url = info.get("url", "")
        if not version or not url:
            return
        text = f"\u2b06 Update available: v{version} \u2014 Download"
        self._update_label.config(text=text)
        self._update_label.bind(
            "<Button-1>",
            lambda _e: __import__("webbrowser").open(url),
        )


def main():
    """Main entry point"""
    root = tk.Tk()
    ResilienceScanGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
