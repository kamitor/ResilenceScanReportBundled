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

import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
import pandas as pd
import os
import shutil
import subprocess
import threading
import queue
from pathlib import Path
from datetime import datetime
import json
import sys
import yaml

# Ensure repo root is on sys.path so sibling modules can be imported
# (needed when running as `python app/main.py` from the repo root)
_repo_root = Path(__file__).resolve().parents[1]
if str(_repo_root) not in sys.path:
    sys.path.insert(0, str(_repo_root))

# Import email tracking system
from email_tracker import EmailTracker  # noqa: E402

# Import system checker
from gui_system_check import SystemChecker, setup_status

# Import dependency manager
from dependency_manager import DependencyManager


# ---------------------------------------------------------------------------
# Path resolution — split into asset root (QMD + images, read-only) and
# data root (CSV, reports, logs — must be user-writable).
# ---------------------------------------------------------------------------
def _asset_root() -> Path:
    """Directory that contains ResilienceReport.qmd and companion assets.

    Dev:    repo root (one level up from app/)
    Frozen: sys._MEIPASS == _internal/ where --add-data extracts files
    """
    if getattr(sys, "frozen", False):
        return Path(sys._MEIPASS)
    return Path(__file__).resolve().parents[1]


def _data_root() -> Path:
    """User-writable directory for data/, reports/, and logs.

    Dev:    repo root (same as asset root, data files live alongside scripts)
    Frozen: APPDATA/ResilienceScan (Windows) or ~/.local/share/resiliencescan (Linux)
    """
    if getattr(sys, "frozen", False):
        if sys.platform == "win32":
            return Path(os.environ.get("APPDATA", str(Path.home()))) / "ResilienceScan"
        return Path.home() / ".local" / "share" / "resiliencescan"
    return Path(__file__).resolve().parents[1]


def _default_output_dir() -> Path:
    """User-visible default folder for generated PDF reports.

    In the frozen app we place reports in Documents/ResilienceScanReports so
    they are easy for users to find.  AppData/Roaming is hidden by default on
    Windows and confuses users.  In dev mode we keep the repo reports/ folder.
    """
    if not getattr(sys, "frozen", False):
        return _data_root() / "reports"
    if sys.platform == "win32":
        docs = Path(os.environ.get("USERPROFILE", str(Path.home()))) / "Documents"
    else:
        docs = Path.home() / "Documents"
    if not docs.exists():
        docs = Path.home()
    return docs / "ResilienceScanReports"


def _sync_template() -> None:
    """Copy QMD and companion assets from _asset_root() to _data_root().

    In the frozen app the QMD lives in _internal/ (under Program Files) which
    is read-only for normal users.  Quarto always creates a .quarto/ scratch
    directory *next to the QMD file*, so it needs to be in a writable location.
    Copying to _data_root() (APPDATA/ResilienceScan) fixes this.

    In dev mode _asset_root() == _data_root() so no copy is needed.
    Only re-copies when the source QMD is newer than the destination (i.e. after
    an app update).
    """
    if not getattr(sys, "frozen", False):
        return
    src = _asset_root()
    dst = _data_root()
    dst.mkdir(parents=True, exist_ok=True)

    # Determine whether any QMD needs updating (use ResilienceReport as the sentinel)
    src_qmd = src / "ResilienceReport.qmd"
    dst_qmd = dst / "ResilienceReport.qmd"
    if (
        dst_qmd.exists()
        and src_qmd.exists()
        and src_qmd.stat().st_mtime <= dst_qmd.stat().st_mtime
    ):
        return  # already up-to-date

    # Copy all QMDs and shared assets
    for name in (
        "ResilienceReport.qmd",
        "SCROLReport.qmd",
        "references.bib",
        "QTDublinIrish.otf",
    ):
        s = src / name
        if s.exists():
            shutil.copy2(str(s), str(dst / name))
    for dname in ("img", "tex", "_extensions"):
        s = src / dname
        d = dst / dname
        if s.exists():
            if d.exists():
                shutil.rmtree(str(d))
            shutil.copytree(str(s), str(d))


ROOT_DIR = _asset_root()  # read-only assets (_internal/ when frozen)
_DATA_ROOT = _data_root()  # data/, reports/, logs — always writable
_sync_template()  # copy QMD + assets to _DATA_ROOT so quarto can write .quarto/ next to them
DATA_FILE = _DATA_ROOT / "data" / "cleaned_master.csv"
REPORTS_DIR = _DATA_ROOT / "reports"
DEFAULT_OUTPUT_DIR = _default_output_dir()  # user-visible reports folder
TEMPLATE = (
    _DATA_ROOT / "ResilienceReport.qmd"
)  # must be in writable _DATA_ROOT, not ROOT_DIR
LOG_FILE = _DATA_ROOT / "gui_log.txt"


def _config_path() -> Path:
    """Return path to config.yml in the writable user data directory."""
    if getattr(sys, "frozen", False):
        if sys.platform == "win32":
            _base = Path(os.environ.get("APPDATA", str(Path.home()))) / "ResilienceScan"
        else:
            _base = Path.home() / ".local" / "share" / "resiliencescan"
    else:
        _base = ROOT_DIR
    return _base / "config.yml"


CONFIG_FILE = _config_path()


def _r_library_path() -> Path | None:
    """Return the bundled R library path when frozen, None in dev mode.

    The NSIS / postinst installer places R packages in an ``r-library``
    directory alongside the executable so the app uses them instead of (or
    in addition to) whatever the user has installed system-wide.
    """
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent / "r-library"
    return None


def _check_r_packages_ready() -> str | None:
    """Return None if all required R packages are findable, or an error string.

    Uses the same R_LIBS setup that the render subprocess uses, so this
    check is representative of what will happen during quarto render.
    Returns immediately (< 5 s) and is safe to call from any thread.
    """
    from gui_system_check import _R_PACKAGES, _find_rscript

    rscript = _find_rscript()
    if not rscript:
        return "Rscript not found on PATH"

    env = os.environ.copy()
    r_lib = _r_library_path()
    if r_lib is not None and r_lib.exists():
        existing = env.get("R_LIBS", "")
        env["R_LIBS"] = f"{r_lib}{os.pathsep}{existing}" if existing else str(r_lib)

    pkg_list = ", ".join(f'"{p}"' for p in _R_PACKAGES)
    script = (
        f"pkgs <- c({pkg_list}); "
        "missing <- pkgs[!pkgs %in% rownames(installed.packages())];"
        "if (length(missing) == 0) cat('OK') "
        "else cat('MISSING:', paste(missing, collapse=', '))"
    )
    try:
        result = subprocess.run(
            [rscript, "-e", script],
            capture_output=True,
            text=True,
            timeout=30,
            env=env,
        )
        out = (result.stdout + result.stderr).strip()
    except Exception as e:
        return f"R check error: {e}"

    if out.strip() == "OK":
        return None
    return out


class ResilienceScanGUI:
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
            text=f"Supply Chain Resilience Assessment Management System  •  v{_current_version()}",
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

    def create_dashboard_tab(self):
        """Create overview dashboard tab"""
        dashboard = ttk.Frame(self.notebook)
        self.notebook.add(dashboard, text="[INFO] Dashboard")

        # Quick actions
        actions_frame = ttk.LabelFrame(dashboard, text="Quick Actions", padding=10)
        actions_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N), padx=10, pady=10)

        ttk.Button(
            actions_frame,
            text="🔄 Reload Data",
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
            text="📧 Send All Emails",
            command=self.start_email_all,
            width=20,
        ).grid(row=0, column=2, padx=5, pady=5)

        ttk.Button(
            actions_frame,
            text="🔧 Check System",
            command=self.run_system_check,
            width=20,
        ).grid(row=1, column=0, padx=5, pady=5)

        ttk.Button(
            actions_frame,
            text="🪟 Install Dependencies (Windows)",
            command=self.install_windows_dependencies,
            width=25,
        ).grid(row=1, column=1, columnspan=2, padx=5, pady=5)

        ttk.Button(
            actions_frame,
            text="🐧 Install Dependencies (Linux)",
            command=self.install_linux_dependencies,
            width=25,
        ).grid(row=1, column=3, padx=5, pady=5)

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
        self.notebook.add(data_tab, text="📁 Data")

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
            controls_frame, text="📥 Convert Data", command=self.run_convert_data
        ).grid(row=0, column=3, padx=5)

        ttk.Button(
            controls_frame, text="🧹 Clean Data", command=self.run_clean_data
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
            text="🔍 Run Quality Dashboard",
            command=self.run_quality_dashboard,
        ).grid(row=1, column=0, sticky=tk.W, pady=(5, 0))

        ttk.Button(
            quality_frame, text="🧹 Run Data Cleaner", command=self.run_data_cleaner
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
            text="▶ Start All",
            command=self.start_generation_all,
            width=20,
        )
        self.gen_start_btn.grid(row=0, column=1, padx=5)

        self.gen_cancel_btn = ttk.Button(
            button_frame,
            text="⏹ Cancel",
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

    def create_email_tab(self):
        """Create email distribution tab"""
        email_tab = ttk.Frame(self.notebook)
        self.notebook.add(email_tab, text="📧 Email")

        # Create notebook for email tab sections
        email_notebook = ttk.Notebook(email_tab)
        email_notebook.grid(
            row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=5
        )

        # Email Template Tab
        template_tab = ttk.Frame(email_notebook)
        email_notebook.add(template_tab, text="✉️ Template")

        # Email Sending Tab
        sending_tab = ttk.Frame(email_notebook)
        email_notebook.add(sending_tab, text="📤 Send Emails")

        email_tab.columnconfigure(0, weight=1)
        email_tab.rowconfigure(0, weight=1)

        # Build template tab
        self.create_email_template_tab(template_tab)

        # Build sending tab (move existing content here)
        self.create_email_sending_tab(sending_tab)

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
            value="Your Resilience Scan Report – {company}"
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
            text="💾 Save Template",
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
            btn_frame, text="👁️ Preview Email", command=self.preview_email, width=15
        ).grid(row=0, column=2, padx=5)

        # SMTP Configuration Section
        smtp_frame = ttk.LabelFrame(
            parent, text="SMTP Server Configuration", padding=10
        )
        smtp_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), padx=10, pady=10)

        # SMTP Server
        ttk.Label(smtp_frame, text="SMTP Server:").grid(
            row=0, column=0, sticky=tk.W, pady=5
        )
        self.smtp_server_var = tk.StringVar(value="smtp.office365.com")
        ttk.Entry(smtp_frame, textvariable=self.smtp_server_var, width=40).grid(
            row=0, column=1, sticky=(tk.W, tk.E), padx=10, pady=5
        )

        # SMTP Port
        ttk.Label(smtp_frame, text="SMTP Port:").grid(
            row=1, column=0, sticky=tk.W, pady=5
        )
        self.smtp_port_var = tk.StringVar(value="587")
        ttk.Entry(smtp_frame, textvariable=self.smtp_port_var, width=10).grid(
            row=1, column=1, sticky=tk.W, padx=10, pady=5
        )

        # From Email
        ttk.Label(smtp_frame, text="From Email:").grid(
            row=2, column=0, sticky=tk.W, pady=5
        )
        self.smtp_from_var = tk.StringVar(value="info@resiliencescan.org")
        from_entry = ttk.Entry(
            smtp_frame, textvariable=self.smtp_from_var, width=40, state="readonly"
        )
        from_entry.grid(row=2, column=1, sticky=(tk.W, tk.E), padx=10, pady=5)

        # SMTP Username
        ttk.Label(smtp_frame, text="SMTP Username:").grid(
            row=3, column=0, sticky=tk.W, pady=5
        )
        self.smtp_username_var = tk.StringVar(value="")
        ttk.Entry(smtp_frame, textvariable=self.smtp_username_var, width=40).grid(
            row=3, column=1, sticky=(tk.W, tk.E), padx=10, pady=5
        )

        # SMTP Password
        ttk.Label(smtp_frame, text="SMTP Password:").grid(
            row=4, column=0, sticky=tk.W, pady=5
        )
        self.smtp_password_var = tk.StringVar(value="")
        ttk.Entry(
            smtp_frame, textvariable=self.smtp_password_var, width=40, show="*"
        ).grid(row=4, column=1, sticky=(tk.W, tk.E), padx=10, pady=5)

        # Help text
        help_text = (
            "Gmail: smtp.gmail.com:587 (use app-specific password)\n"
            "Office365: smtp.office365.com:587\n"
            "Outlook.com: smtp-mail.outlook.com:587"
        )
        ttk.Label(
            smtp_frame, text=help_text, font=("Arial", 8), foreground="gray"
        ).grid(row=5, column=0, columnspan=2, sticky=tk.W, pady=5)

        ttk.Button(
            smtp_frame, text="Save Configuration", command=self.save_config
        ).grid(row=6, column=0, columnspan=2, sticky=tk.W, pady=5)

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
            font=("Arial", 10, "bold"),
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

        # Test mode
        self.test_mode_var = tk.BooleanVar(value=True)
        test_check = ttk.Checkbutton(
            controls_frame,
            text="Test Mode (send to test email only)",
            variable=self.test_mode_var,
            command=self.toggle_test_mode,
        )
        test_check.grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=5)

        ttk.Label(controls_frame, text="Test Email:").grid(row=1, column=0, sticky=tk.W)
        self.test_email_var = tk.StringVar(value="")
        ttk.Entry(controls_frame, textvariable=self.test_email_var, width=40).grid(
            row=1, column=1, sticky=(tk.W, tk.E), padx=10
        )

        controls_frame.columnconfigure(1, weight=1)

        # Action buttons
        button_frame = ttk.Frame(controls_frame)
        button_frame.grid(row=2, column=0, columnspan=2, pady=10)

        self.email_start_btn = ttk.Button(
            button_frame, text="▶ Start Sending", command=self.start_email_all, width=20
        )
        self.email_start_btn.grid(row=0, column=0, padx=5)

        self.email_stop_btn = ttk.Button(
            button_frame,
            text="⏹ Stop",
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

    def create_logs_tab(self):
        """Create system logs tab"""
        logs_tab = ttk.Frame(self.notebook)
        self.notebook.add(logs_tab, text="📋 Logs")

        # Controls
        controls_frame = ttk.Frame(logs_tab)
        controls_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=10, pady=10)

        ttk.Button(
            controls_frame, text="🔄 Refresh Logs", command=self.refresh_logs
        ).grid(row=0, column=0, padx=5)

        ttk.Button(controls_frame, text="🗑️ Clear Logs", command=self.clear_logs).grid(
            row=0, column=1, padx=5
        )

        ttk.Button(
            controls_frame, text="💾 Export Logs", command=self.export_logs
        ).grid(row=0, column=2, padx=5)

        # System log
        log_frame = ttk.LabelFrame(logs_tab, text="System Log", padding=10)
        log_frame.grid(
            row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=10, pady=10
        )

        self.system_log = scrolledtext.ScrolledText(
            log_frame, wrap=tk.WORD, width=80, height=25, font=("Courier", 9)
        )
        self.system_log.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        logs_tab.columnconfigure(0, weight=1)
        logs_tab.rowconfigure(1, weight=1)

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
        text = f"⬆ Update available: v{version} — Download"
        self._update_label.config(text=text)
        self._update_label.bind(
            "<Button-1>",
            lambda _e: __import__("webbrowser").open(url),
        )

    # ==================== Data Methods ====================

    def _startup_guard(self):
        """Check that R, Quarto, TinyTeX, and R packages are present; show a
        blocking warning dialog if any critical component is missing."""
        checker = SystemChecker()
        result = checker.check_all()
        install_status = setup_status()

        critical = ["R", "quarto", "tinytex"]
        missing = [k for k in critical if not result.get(k, {}).get("ok")]

        if missing:
            names = {"R": "R", "quarto": "Quarto", "tinytex": "TinyTeX (tlmgr)"}
            missing_str = "\n".join(f"  • {names[k]}" for k in missing)
            if install_status == "running":
                messagebox.showinfo(
                    "Setup In Progress",
                    "Dependency setup is still running in the background.\n\n"
                    "The following components are not ready yet:\n\n"
                    f"{missing_str}\n\n"
                    "This normally takes 5–20 minutes after installation.\n"
                    "The status bar will update when setup completes.\n\n"
                    "You can use the app in the meantime, but generating\n"
                    "PDFs will not work until setup finishes.",
                )
            elif install_status == "complete_fail":
                log_hint = (
                    r"C:\ProgramData\ResilienceScan\setup.log"
                    if sys.platform == "win32"
                    else "/var/log/resilencescan-setup.log"
                )
                messagebox.showwarning(
                    "Setup Failed",
                    "The background dependency setup finished with errors.\n\n"
                    "Missing components:\n\n"
                    f"{missing_str}\n\n"
                    f"Check the setup log for details:\n{log_hint}\n\n"
                    "Re-run setup or contact support.",
                )
            else:
                messagebox.showwarning(
                    "Missing Components",
                    "The following required components were not found on PATH:\n\n"
                    f"{missing_str}\n\n"
                    "The installation may be incomplete.  Report generation will\n"
                    "not work until these are installed.\n\n"
                    "You can continue, but generating PDFs will fail.",
                )

        # Separately warn if R packages are missing (non-blocking, but important).
        if not result.get("r_packages", {}).get("ok"):
            log_hint = (
                r"C:\ProgramData\ResilienceScan\setup.log"
                if sys.platform == "win32"
                else "/var/log/resilencescan-setup.log"
            )
            if install_status == "running":
                # Already covered by the "Setup In Progress" dialog above (or no
                # critical components were missing).  Just update the status bar.
                pass
            elif install_status == "complete_fail":
                messagebox.showwarning(
                    "R Packages Missing",
                    "Required R packages failed to install.\n\n"
                    f"Check the setup log:\n{log_hint}\n\n"
                    "Report generation will fail until packages are available.",
                )
            else:
                messagebox.showwarning(
                    "R Packages Not Ready",
                    "Required R packages are not yet installed.\n\n"
                    "The background setup may still be running (allow 5-20 minutes\n"
                    "after installation) or it may have failed.\n\n"
                    f"Check the setup log:\n{log_hint}\n\n"
                    "Report generation will fail until packages are available.\n"
                    "Use the System Check button to re-check.",
                )

        # If setup is still running, show a status bar indicator and start polling.
        if install_status == "running":
            self.status_label.config(text="Installing dependencies... (5-20 min)")
            self.root.after(30_000, self._poll_setup_completion)

    def _poll_setup_completion(self):
        """Poll every 30 s for background setup completion; update status bar."""
        status = setup_status()
        if status == "complete_pass":
            self.status_label.config(text="Setup complete — all dependencies ready.")
            self.root.after(10_000, lambda: self.status_label.config(text="Ready"))
        elif status == "complete_fail":
            self.status_label.config(
                text="Setup finished with errors — see System Check."
            )
        elif status == "running":
            self.root.after(30_000, self._poll_setup_completion)
        # else 'unknown' (dev mode / flags cleared) — stop polling silently

    def load_config(self):
        """Load SMTP settings from config.yml into GUI fields."""
        if not CONFIG_FILE.exists():
            return
        try:
            data = yaml.safe_load(CONFIG_FILE.read_text(encoding="utf-8")) or {}
            smtp = data.get("smtp", {})
            if smtp.get("server"):
                self.smtp_server_var.set(smtp["server"])
            if smtp.get("port"):
                self.smtp_port_var.set(str(smtp["port"]))
            if smtp.get("from_address"):
                self.smtp_from_var.set(smtp["from_address"])
            if smtp.get("username"):
                self.smtp_username_var.set(smtp["username"])
            if smtp.get("password"):
                self.smtp_password_var.set(smtp["password"])
        except Exception as e:
            self.log(f"[WARNING] Could not load config.yml: {e}")

    def save_config(self):
        """Save SMTP settings from GUI fields to config.yml."""
        data = {
            "smtp": {
                "server": self.smtp_server_var.get(),
                "port": int(self.smtp_port_var.get() or 587),
                "from_address": self.smtp_from_var.get(),
                "username": self.smtp_username_var.get(),
                "password": self.smtp_password_var.get(),
            }
        }
        try:
            CONFIG_FILE.parent.mkdir(parents=True, exist_ok=True)
            CONFIG_FILE.write_text(
                yaml.dump(data, default_flow_style=False, allow_unicode=True),
                encoding="utf-8",
            )
            messagebox.showinfo("Saved", f"Configuration saved to:\n{CONFIG_FILE}")
        except Exception as e:
            messagebox.showerror("Error", f"Could not save configuration:\n{e}")

    def load_initial_data(self):
        """Load data on startup"""
        self.log("Loading data from: " + str(DATA_FILE))
        try:
            if DATA_FILE.exists():
                self.df = pd.read_csv(DATA_FILE)
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
                ("All supported formats", "*.xlsx *.xls *.ods *.xml *.csv *.tsv"),
                ("Excel files", "*.xlsx *.xls"),
                ("OpenDocument Spreadsheet", "*.ods"),
                ("XML files", "*.xml"),
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
                self.log(f"[INFO] Converting {fmt} → CSV …")
                ok = _cd.convert_and_save(dest)
                if not ok:
                    messagebox.showerror(
                        "Conversion Failed",
                        f"Could not convert {dest.name}.\nCheck the log for details.",
                    )
                    return
                self.log("[OK] Conversion complete — loading cleaned_master.csv")
                csv_path = DATA_FILE

            self.df = pd.read_csv(csv_path)
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
                    self.df = pd.read_csv(DATA_FILE)
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
                    self.df = pd.read_csv(DATA_FILE)
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
        from pathlib import Path
        import json

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
                    text="📋 View Detailed Validation Log",
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
        import json

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

                # Load the validation results
                from pathlib import Path
                import json

                report_path = Path("./data/integrity_validation_report.txt")
                json_path = Path("./data/integrity_validation_report.json")

                if json_path.exists():
                    with open(json_path, "r") as f:
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
═══════════════════════════════════════════════════════════════
RESILIENCESCAN DATA OVERVIEW
═══════════════════════════════════════════════════════════════

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

    # ==================== Data Analysis Methods ====================

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
            info_text += f" (filtered)"
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

    def analyze_data_quality(self):
        """Analyze data quality and show summary"""
        if self.df is None:
            return

        # Calculate statistics
        total_records = len(self.df)
        missing_email = self.df["email_address"].isna().sum()
        invalid_email = (
            ~self.df["email_address"].str.contains("@", na=False)
        ).sum() - missing_email
        duplicates = self.df.duplicated(
            subset=["company_name", "name", "email_address"]
        ).sum()
        unique_companies = self.df["company_name"].nunique()

        # Build quality report
        report = f"Total: {total_records} records | "
        report += f"Companies: {unique_companies} | "
        report += f"Missing email: {missing_email} | "
        report += f"Invalid email: {invalid_email} | "
        report += f"Duplicates: {duplicates}"

        # Update quality text
        self.quality_text.delete("1.0", tk.END)
        self.quality_text.insert("1.0", report)

        # Highlight issues
        if missing_email > 0 or invalid_email > 0 or duplicates > 0:
            self.quality_text.config(bg="#fff3cd")  # Warning yellow
        else:
            self.quality_text.config(bg="#d4edda")  # Success green

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
        for idx, row in duplicates.iterrows():
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
                self.filtered_df.to_csv(filename, index=False)
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
        for idx, row in self.filtered_df.iterrows():
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

    # ==================== Generation Methods ====================

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
        from datetime import datetime
        import subprocess
        import shutil

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
            # Create safe filenames
            def safe_filename(name):
                if pd.isna(name) or name == "":
                    return "Unknown"
                return "".join(
                    c if c.isalnum() or c in [" ", "-"] else "_" for c in str(name)
                ).replace(" ", "_")

            def safe_display_name(name):
                if pd.isna(name) or name == "":
                    return "Unknown"
                name_str = str(name).strip()
                name_str = (
                    name_str.replace("/", "-").replace("\\", "-").replace(":", "-")
                )
                name_str = name_str.replace("*", "").replace("?", "").replace('"', "'")
                name_str = (
                    name_str.replace("<", "(").replace(">", ")").replace("|", "-")
                )
                return name_str

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
                                f"[OK] Validation passed: All values match CSV"
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
                            for key, info in validation_result.get(
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
                    self.log_gen(f"[ERROR] Output file not found after rendering")
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
                    f"Report generation failed.\n\nCheck logs for details.",
                )
                self.status_label.config(text="Error")

        except FileNotFoundError:
            self.log_gen(
                f"[ERROR] Quarto not found - please install from https://quarto.org"
            )
            messagebox.showerror(
                "Quarto Not Found",
                "Quarto is not installed.\n\nPlease install from https://quarto.org",
            )
            self.status_label.config(text="Error")
        except subprocess.TimeoutExpired:
            self.log_gen(f"[ERROR] Generation timeout (>5 minutes)")
            messagebox.showerror("Timeout", "Report generation timed out (>5 minutes)")
            self.status_label.config(text="Error")
        except Exception as e:
            self.log_gen(f"[ERROR] Error: {e}")
            messagebox.showerror("Error", f"Report generation failed:\n{e}")
            self.status_label.config(text="Error")

    def start_generation_all(self):
        """Start generating all reports"""
        if self.df is None:
            messagebox.showwarning("Warning", "Please load data first")
            return

        if self.is_generating:
            messagebox.showwarning("Warning", "Generation already in progress")
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
                    except:
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
            self.log_gen(f"[ERROR] R packages not ready — aborting batch.")
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
                self.gen_current_label.config(text="Aborted — R packages missing")

            self.root.after(0, _reset_ui)
            return

        total = len(self.df)
        success = 0
        failed = 0
        skipped = 0

        self.root.after(0, lambda: self.gen_progress.config(maximum=total, value=0))

        for idx, row in self.df.iterrows():
            try:
                if self._stop_gen.is_set():
                    self.log_gen("Generation cancelled by user")
                    break

                company = row.get("company_name", "Unknown")
                person = row.get("name", "Unknown")

                try:
                    display_text = f"Generating: {company} - {person}"
                except (UnicodeDecodeError, UnicodeEncodeError):
                    display_text = "Generating: {safe} - {safe}".format(
                        safe=company.encode("ascii", "replace").decode("ascii")
                    )
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

                # Create safe filenames
                def safe_filename(name):
                    if pd.isna(name) or name == "":
                        return "Unknown"
                    return "".join(
                        c if c.isalnum() or c in [" ", "-"] else "_" for c in str(name)
                    ).replace(" ", "_")

                def safe_display_name(name):
                    if pd.isna(name) or name == "":
                        return "Unknown"
                    name_str = str(name).strip()
                    name_str = (
                        name_str.replace("/", "-").replace("\\", "-").replace(":", "-")
                    )
                    name_str = (
                        name_str.replace("*", "").replace("?", "").replace('"', "'")
                    )
                    name_str = (
                        name_str.replace("<", "(").replace(">", ")").replace("|", "-")
                    )
                    return name_str

                safe_company = safe_filename(company)
                safe_person = safe_filename(person)
                display_company = safe_display_name(company)
                display_person = safe_display_name(person)

                # Output filename with template name
                from datetime import datetime

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
                    self.log_gen(f"  [SKIP] Already exists, skipping")
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
                        import shutil

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
                                    f"  [OK] Validation passed: All values match CSV"
                                )
                            else:
                                self.log_gen(
                                    f"  [WARNING] Validation: {validation_result['message']}"
                                )
                                # Log details
                                for key, info in validation_result.get(
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
                            f"  [ERROR] Error: Output file not found after render"
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

            # Update progress bar (thread-safe)
            _s, _f, _sk, _i, _t = success, failed, skipped, idx + 1, total
            self.root.after(
                0,
                lambda v=_i, s=_s, f=_f, sk=_sk, t=_t: (
                    self.gen_progress.config(value=v),
                    self.gen_progress_label.config(
                        text=f"Progress: {v}/{t} | Success: {s} | Failed: {f} | Skipped: {sk}"
                    ),
                ),
            )

        def _finish_ui():
            self.is_generating = False
            self.gen_start_btn.config(state=tk.NORMAL)
            self.gen_cancel_btn.config(state=tk.DISABLED)
            self.gen_current_label.config(text="Generation complete")

        self.root.after(0, _finish_ui)

        # Comprehensive summary
        self.log_gen(f"\n" + "=" * 60)

        # Check if generation completed all records
        processed = success + failed + skipped
        if processed < total:
            self.log_gen(f"WARNING: GENERATION INCOMPLETE")
            self.log_gen(f"=" * 60)
            self.log_gen(f"Total records: {total}")
            self.log_gen(f"Processed: {processed}/{total}")
            self.log_gen(f"Successfully generated: {success}")
            self.log_gen(f"Failed: {failed}")
            self.log_gen(f"Skipped (insufficient data): {skipped}")
            self.log_gen(f"NOT PROCESSED: {total - processed}")
            self.log_gen(f"=" * 60)
            self.log_gen(f"\nCRITICAL: Generation stopped early at record {processed}.")
            self.log_gen(f"Check error messages above for details.")
        else:
            self.log_gen(f"GENERATION COMPLETE")
            self.log_gen(f"=" * 60)
            self.log_gen(f"Total records: {total}")
            self.log_gen(f"Successfully generated: {success}")
            self.log_gen(f"Failed: {failed}")
            self.log_gen(f"Skipped (insufficient data): {skipped}")
            self.log_gen(f"=" * 60)

        if skipped > 0:
            self.log_gen(
                f"\nNote: {skipped} record(s) were skipped due to insufficient data."
            )
            self.log_gen(
                f"   These records don't have enough scores to generate a valid report."
            )
            self.log_gen(
                f"   Run 'Clean Data' to see details about removed/insufficient records."
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

    def browse_output_folder(self):
        """Browse for output folder"""
        folder = filedialog.askdirectory(
            title="Select Output Folder", initialdir=self.output_folder_var.get()
        )
        if folder:
            self.output_folder_var.set(folder)

    def run_system_check(self):
        """Run system check and display results"""
        self.log("Running system check...")
        self.status_label.config(text="Checking system...")

        try:
            # Run system check
            checker = SystemChecker(ROOT_DIR)
            check_result = checker.check_all()
            all_ok = all(v.get("ok", False) for v in check_result.values())

            # Clear stats text and show results
            self.stats_text.delete("1.0", tk.END)

            # Build report
            report = "=" * 70 + "\n"
            report += "SYSTEM CHECK REPORT\n"
            report += "=" * 70 + "\n"
            report += f"Checked at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
            report += "=" * 70 + "\n\n"

            # Summary
            total_checks = len(checker.checks)
            errors = len(checker.errors)
            warnings = len(checker.warnings)
            successes = total_checks - errors - warnings

            if all_ok:
                report += "[OK] SYSTEM STATUS: ALL CHECKS PASSED\n\n"
            else:
                report += f"[WARNING] SYSTEM STATUS: {errors} ERROR(S), {warnings} WARNING(S)\n\n"

            report += (
                f"Summary: {successes} OK | {warnings} Warnings | {errors} Errors\n"
            )
            report += "=" * 70 + "\n\n"

            # Detailed results
            for check in checker.checks:
                icon = check["item"].split()[0]
                item = " ".join(check["item"].split()[1:])
                report += f"{icon} {item:<40} {check['status']:<20}\n"
                if check["description"]:
                    report += f"   → {check['description']}\n"
                report += "\n"

            # Show errors and warnings
            if checker.errors:
                report += "\n" + "=" * 70 + "\n"
                report += "ERRORS FOUND:\n"
                report += "=" * 70 + "\n"
                for error in checker.errors:
                    report += f"[ERROR] {error}\n"

            if checker.warnings:
                report += "\n" + "=" * 70 + "\n"
                report += "WARNINGS:\n"
                report += "=" * 70 + "\n"
                for warning in checker.warnings:
                    report += f"[WARNING] {warning}\n"

            # Display in stats text
            self.stats_text.insert("1.0", report)

            # Log summary
            self.log(
                f"[OK] System check complete: {successes} OK, {warnings} warnings, {errors} errors"
            )

            # Show summary dialog
            if all_ok:
                messagebox.showinfo(
                    "System Check Complete",
                    f"[OK] All checks passed!\n\n{total_checks} checks completed successfully.",
                )
            else:
                messagebox.showwarning(
                    "System Check Complete",
                    f"[WARNING] Found {errors} error(s) and {warnings} warning(s)\n\nSee Dashboard for details.",
                )

        except Exception as e:
            self.log(f"[ERROR] Error running system check: {e}")
            messagebox.showerror("Error", f"Failed to run system check:\n{e}")
        finally:
            self.status_label.config(text="Ready")

    def install_windows_dependencies(self):
        """Install dependencies on Windows - runs installation/install_dependencies_auto.py"""
        import platform
        import subprocess

        if platform.system() != "Windows":
            messagebox.showwarning(
                "Wrong Platform",
                "This option is for Windows systems.\n\nYou are running on: "
                + platform.system(),
            )
            return

        self.log("Starting Windows dependency installation...")
        self.status_label.config(text="Installing dependencies...")

        # Clear stats text
        self.stats_text.delete("1.0", tk.END)

        report = "=" * 70 + "\n"
        report += "WINDOWS DEPENDENCY INSTALLATION\n"
        report += "=" * 70 + "\n\n"
        report += "Running installation/install_dependencies_auto.py...\n\n"

        self.stats_text.insert("1.0", report)
        self.stats_text.update()

        try:
            # Run the installation script from the installation folder
            install_script = ROOT_DIR / "installation" / "install_dependencies_auto.py"

            if not install_script.exists():
                raise FileNotFoundError(
                    f"Installation script not found: {install_script}"
                )

            # Run the script and capture output
            result = subprocess.run(
                [sys.executable, str(install_script)],
                capture_output=True,
                text=True,
                cwd=ROOT_DIR,
                timeout=300,
            )

            # Show the output
            output = result.stdout if result.stdout else ""
            if result.stderr:
                output += "\n\nErrors:\n" + result.stderr

            self.stats_text.insert(tk.END, output)

            if result.returncode == 0:
                self.log("Installation completed successfully")
                messagebox.showinfo(
                    "Installation Complete",
                    "Python packages installed successfully!\n\n"
                    "For R and Quarto, run PowerShell installer:\n"
                    "installation/Install-ResilienceScan.ps1\n\n"
                    "See installation/INSTALL.md for details.",
                )
            else:
                self.log(
                    f"Installation completed with errors (exit code: {result.returncode})"
                )
                messagebox.showwarning(
                    "Installation Completed with Errors",
                    "Some packages may have failed to install.\n\n"
                    "Check the Dashboard for details.",
                )

        except FileNotFoundError as e:
            self.log(f"Error: Installation script not found")
            self.stats_text.insert(tk.END, f"\nERROR: {e}\n")
            messagebox.showerror(
                "Installation Script Not Found",
                f"Could not find installation script:\n{e}\n\n"
                "Please ensure installation/ folder exists.",
            )
        except subprocess.TimeoutExpired:
            self.log("Error: Installation timed out after 5 minutes")
            messagebox.showerror(
                "Installation Timeout",
                "Installation took longer than 5 minutes and was cancelled.",
            )
        except Exception as e:
            self.log(f"Error during installation: {e}")
            self.stats_text.insert(tk.END, f"\nERROR: {e}\n")
            import traceback

            traceback_str = traceback.format_exc()
            self.stats_text.insert(tk.END, f"\n{traceback_str}\n")
            messagebox.showerror(
                "Installation Error", f"Failed to install dependencies:\n\n{e}"
            )
        finally:
            self.status_label.config(text="Ready")

    def install_linux_dependencies(self):
        """Install dependencies on Linux"""
        import platform

        if platform.system() != "Linux":
            messagebox.showwarning(
                "Wrong Platform",
                "This option is for Linux systems.\n\nYou are running on: "
                + platform.system(),
            )
            return

        self.log("Starting Linux dependency installation...")
        self.status_label.config(text="Installing dependencies...")

        try:
            manager = DependencyManager()
            checks = manager.check_all()

            # Clear stats text and show installation commands
            self.stats_text.delete("1.0", tk.END)

            report = "=" * 70 + "\n"
            report += "LINUX DEPENDENCY INSTALLATION GUIDE\n"
            report += "=" * 70 + "\n\n"

            # Auto-install Python packages
            python_packages_installed = 0
            python_packages_failed = 0

            for check in checks:
                if check["category"] == "Python Packages" and not check["installed"]:
                    package_name = check["name"].replace("Python Package: ", "")
                    self.log(f"Installing Python package: {package_name}")
                    report += f"Installing {package_name}...\n"

                    result = manager.install_package(package_name)
                    if result["success"]:
                        report += f"  [OK] {package_name} installed successfully\n\n"
                        python_packages_installed += 1
                    else:
                        report += f"  [ERROR] Failed to install {package_name}\n"
                        report += f"  Error: {result['error']}\n\n"
                        python_packages_failed += 1

            # Installation commands for R and Quarto
            report += "\n" + "=" * 70 + "\n"
            report += "SYSTEM PACKAGE INSTALLATION COMMANDS\n"
            report += "=" * 70 + "\n\n"

            report += "Copy and run these commands in your terminal:\n\n"

            for check in checks:
                if not check["installed"] and check["category"] in ["R", "Quarto"]:
                    install_cmd = manager.get_install_command(check["name"])
                    report += f"# Install {check['name']}\n"
                    if "command" in install_cmd:
                        report += f"{install_cmd['command']}\n\n"

            # Summary
            report += "\n" + "=" * 70 + "\n"
            report += "INSTALLATION SUMMARY\n"
            report += "=" * 70 + "\n"
            report += f"Python packages installed: {python_packages_installed}\n"
            report += f"Python packages failed: {python_packages_failed}\n"
            report += f"\nRun the commands above to install R and Quarto.\n"
            report += f"Then click 'Check System' to verify installation.\n"

            self.stats_text.insert("1.0", report)
            self.log(
                f"[OK] Linux installation guide displayed: {python_packages_installed} packages installed"
            )

            messagebox.showinfo(
                "Installation Guide",
                f"[OK] Installed {python_packages_installed} Python package(s)\n\n"
                f"See Dashboard for commands to install R and Quarto.",
            )

        except Exception as e:
            self.log(f"[ERROR] Error installing dependencies: {e}")
            messagebox.showerror("Error", f"Failed to install dependencies:\n{e}")
        finally:
            self.status_label.config(text="Ready")

    # ==================== Email Template Methods ====================

    def save_email_template(self):
        """Save email template to file"""
        try:
            template_data = {
                "subject": self.email_subject_var.get(),
                "body": self.email_body_text.get("1.0", tk.END).strip(),
            }

            template_file = ROOT_DIR / "email_template.json"
            with open(template_file, "w") as f:
                json.dump(template_data, f, indent=2)

            self.log("[OK] Email template saved")
            messagebox.showinfo("Success", "Email template saved successfully!")

        except Exception as e:
            self.log(f"[ERROR] Error saving template: {e}")
            messagebox.showerror("Error", f"Failed to save template:\n{e}")

    def load_email_template(self):
        """Load email template from file"""
        try:
            template_file = ROOT_DIR / "email_template.json"
            if template_file.exists():
                with open(template_file, "r") as f:
                    template_data = json.load(f)

                self.email_subject_var.set(
                    template_data.get(
                        "subject", "Your Resilience Scan Report – {company}"
                    )
                )
                self.email_body_text.delete("1.0", tk.END)
                self.email_body_text.insert("1.0", template_data.get("body", ""))

                self.log("[OK] Email template loaded")
        except Exception as e:
            self.log(f"[WARNING] Could not load template: {e}")

    def reset_email_template(self):
        """Reset to default template"""
        default_subject = "Your Resilience Scan Report – {company}"
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
        from datetime import datetime

        sample_date = datetime.now().strftime("%Y-%m-%d")

        subject = subject_template.format(
            company=sample_company, name=sample_name, date=sample_date
        )

        body = body_template.format(
            company=sample_company, name=sample_name, date=sample_date
        )

        # Find report file
        from pathlib import Path
        import glob

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

    # ==================== Email Methods ====================

    def update_email_status_display(self):
        """Update email status treeview - ONLY shows companies with generated PDF reports"""
        # Load CSV data if not already loaded
        if self.df is None and DATA_FILE.exists():
            try:
                self.df = pd.read_csv(DATA_FILE)
                self.df.columns = self.df.columns.str.lower().str.strip()
                self.log_email("[LOAD] Loaded CSV data for email display")
            except Exception as e:
                self.log_email(f"[WARNING] Could not load CSV: {e}")

        # Clear existing items
        for item in self.email_status_tree.get_children():
            self.email_status_tree.delete(item)

        # Scan output folder for PDF files
        import glob

        _out_dir = Path(self.output_folder_var.get())
        report_files = glob.glob(str(_out_dir / "*.pdf"))

        if not report_files:
            self.log_email(f"[INFO] No PDF reports found in {_out_dir}")
            self.email_stats_label.config(
                text="No PDF reports found - generate reports first"
            )
            return

        # Parse PDF filenames to extract company and person info
        # Format: YYYYMMDD ResilienceScanReport (COMPANY - PERSON).pdf
        reports_ready = []

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
                    # Split by " - " to get company and person
                    company, person = content.rsplit(" - ", 1)

                    # Look up email address from CSV data
                    email = ""
                    if self.df is not None:
                        # Find matching record
                        matches = self.df[
                            (self.df["company_name"].str.strip() == company.strip())
                            & (self.df["name"].str.strip() == person.strip())
                        ]
                        if not matches.empty:
                            email = matches.iloc[0].get("email_address", "")

                    # Check status: prefer email_tracker (updated by send thread,
                    # reflects test-mode sends) then fall back to CSV reportsent.
                    tracker_key = f"{company.strip()}|{person.strip()}"
                    tracker_entry = self.email_tracker._recipients.get(tracker_key)
                    if tracker_entry:
                        sent_status = tracker_entry["status"]  # pending/sent/failed
                    else:
                        sent_status = "pending"
                        if self.df is not None:
                            matches = self.df[
                                (self.df["company_name"].str.strip() == company.strip())
                                & (self.df["name"].str.strip() == person.strip())
                            ]
                            if not matches.empty and "reportsent" in self.df.columns:
                                is_sent = matches.iloc[0].get("reportsent", False)
                                if is_sent:
                                    sent_status = "sent"

                    reports_ready.append(
                        {
                            "company": company,
                            "person": person,
                            "email": email,
                            "status": sent_status,
                            "pdf_path": pdf_path,
                        }
                    )
            except Exception as e:
                self.log_email(f"[WARNING] Could not parse filename: {filename} - {e}")
                continue

        # Update statistics
        total = len(reports_ready)
        pending = sum(1 for r in reports_ready if r["status"] == "pending")
        sent = sum(1 for r in reports_ready if r["status"] == "sent")
        failed = sum(1 for r in reports_ready if r["status"] == "failed")

        self.email_stats_label.config(
            text=f"Reports Ready: {total} | Pending: {pending} | Sent: {sent} | Failed: {failed}"
        )

        # Get filter value
        filter_status = self.email_filter_var.get()

        # Display reports
        for report in reports_ready:
            # Apply filter
            if filter_status != "all" and report["status"] != filter_status:
                continue

            # Insert into tree with tag for color coding
            values = (
                report["company"],
                report["person"],
                report["email"] if report["email"] else "NO EMAIL",
                report["status"].upper(),
                "",  # No date for pending
                "",  # No mode needed
            )

            item = self.email_status_tree.insert("", tk.END, values=values)

            # Color code by status
            if report["status"] == "sent":
                self.email_status_tree.item(item, tags=("sent",))
            else:
                self.email_status_tree.item(item, tags=("pending",))

        # Configure tag colors
        self.email_status_tree.tag_configure("sent", foreground="green")
        self.email_status_tree.tag_configure("pending", foreground="orange")

    def mark_as_sent_in_csv(self, company, person):
        """Mark a report as sent in the CSV file"""
        try:
            # Update in-memory dataframe
            if self.df is not None and "reportsent" in self.df.columns:
                mask = (self.df["company_name"].str.strip() == company.strip()) & (
                    self.df["name"].str.strip() == person.strip()
                )
                self.df.loc[mask, "reportsent"] = True

                # Save back to CSV file
                self.df.to_csv(DATA_FILE, index=False)

                # Reload the CSV to ensure we have the latest data
                self.df = pd.read_csv(DATA_FILE)
                self.df.columns = self.df.columns.str.lower().str.strip()

                self.log_email(
                    f"  [UPDATE] Updated CSV: {company} - {person} marked as sent"
                )
        except Exception as e:
            self.log_email(f"  [WARNING] Could not update CSV: {e}")

    def mark_selected_as_sent(self):
        """Mark selected email as sent"""
        selection = self.email_status_tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select an email record first")
            return

        for item in selection:
            values = self.email_status_tree.item(item)["values"]
            company, person, email = values[0], values[1], values[2]

            # Update in CSV
            self.mark_as_sent_in_csv(company, person)

        self.update_email_status_display()
        self.log_email(f"[OK] Marked {len(selection)} record(s) as sent")

    def mark_selected_as_pending(self):
        """Reset selected email to pending"""
        selection = self.email_status_tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select an email record first")
            return

        for item in selection:
            values = self.email_status_tree.item(item)["values"]
            company, person, email = values[0], values[1], values[2]

            # Reset in CSV by setting reportsent to False
            try:
                if self.df is not None and "reportsent" in self.df.columns:
                    mask = (self.df["company_name"].str.strip() == company.strip()) & (
                        self.df["name"].str.strip() == person.strip()
                    )
                    self.df.loc[mask, "reportsent"] = False

                    # Save back to CSV file
                    self.df.to_csv(DATA_FILE, index=False)
            except Exception as e:
                self.log_email(f"[WARNING] Could not update CSV: {e}")

        self.update_email_status_display()
        self.log_email(f"[RESET] Reset {len(selection)} record(s) to pending")

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
            if not test_email or "@" not in test_email:
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
            smtp_port = int(self.smtp_port_var.get() or 587)
            if not 1 <= smtp_port <= 65535:
                raise ValueError(f"Port must be 1–65535, got {smtp_port}")
        except ValueError as port_err:
            messagebox.showerror("Invalid Port", str(port_err))
            return

        self.is_sending_emails = True
        self.email_start_btn.config(state=tk.DISABLED)
        self.email_stop_btn.config(state=tk.NORMAL)

        # Capture all widget values on the main thread before the background
        # thread starts — Tkinter widgets are not thread-safe.
        send_config = {
            "smtp_server": self.smtp_server_var.get(),
            "smtp_port": smtp_port,
            "smtp_username": self.smtp_username_var.get(),
            "smtp_password": self.smtp_password_var.get(),
            "smtp_from": self.smtp_from_var.get(),
            "out_dir": Path(self.output_folder_var.get()),
            "test_mode": self.test_mode_var.get(),
            "test_email": self.test_email_var.get().strip(),
            "subject_template": self.email_subject_var.get(),
            "body_template": self.email_body_text.get("1.0", tk.END).strip(),
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

            def _finalize_error():
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

        # Scan output folder for PDF files (same as display logic)
        import glob

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

                    # Look up email address from CSV data
                    email = ""
                    if self.df is not None:
                        matches = self.df[
                            (self.df["company_name"].str.strip() == company.strip())
                            & (self.df["name"].str.strip() == person.strip())
                        ]
                        if not matches.empty:
                            email = matches.iloc[0].get("email_address", "")

                    # Check if already sent (from CSV reportsent column)
                    is_sent = False
                    if self.df is not None and "reportsent" in self.df.columns:
                        matches = self.df[
                            (self.df["company_name"].str.strip() == company.strip())
                            & (self.df["name"].str.strip() == person.strip())
                        ]
                        if not matches.empty:
                            is_sent = matches.iloc[0].get("reportsent", False)

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
            def update_current():
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
                import pandas as pd

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
                self.log_email(f"  Formatting email...")
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
                use_outlook = True
                outlook_error = None

                if use_outlook:
                    try:
                        self.log_email(f"  [OUTLOOK] Attempting to send via Outlook...")
                        import win32com.client

                        # Create Outlook instance
                        outlook = win32com.client.Dispatch("Outlook.Application")

                        # Define account priority list
                        priority_accounts = [
                            "info@resiliencescan.org",
                            "r.deboer@windesheim.nl",
                            "cg.verhoef@windesheim.nl",
                        ]

                        # Get all available accounts
                        available_accounts = []
                        try:
                            for i in range(1, outlook.Session.Accounts.Count + 1):
                                account = outlook.Session.Accounts.Item(i)
                                available_accounts.append(
                                    (account.SmtpAddress, account)
                                )
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
                                f"  [WARNING] No specific account found, using Outlook default"
                            )

                        # Add attachment
                        self.log_email(f"  Attaching PDF: {attachment_path}...")
                        mail.Attachments.Add(str(attachment_path.absolute()))

                        # VERIFICATION: Log the actual recipient before sending
                        self.log_email(f"  [OK] Email configured:")
                        self.log_email(f"     To: {mail.To}")
                        self.log_email(
                            f"     From: {selected_address if selected_address else '(default account)'}"
                        )
                        self.log_email(f"     Subject: {mail.Subject[:50]}...")

                        # Send the email
                        self.log_email(f"  [OUTLOOK] Sending via Outlook...")
                        mail.Send()
                        self.log_email(
                            f"  [OK] Email sent successfully via Outlook from {selected_address if selected_address else 'default account'}!"
                        )

                    except Exception as outlook_ex:
                        outlook_error = str(outlook_ex)
                        self.log_email(f"  [WARNING] Outlook failed: {outlook_error}")
                        self.log_email(f"  [FALLBACK] Attempting SMTP as fallback...")

                        # Fallback to SMTP
                        import smtplib
                        from email.mime.multipart import MIMEMultipart
                        from email.mime.text import MIMEText
                        from email.mime.base import MIMEBase
                        from email import encoders

                        # Create message
                        self.log_email(
                            f"  [SMTP] Creating SMTP message as final fallback..."
                        )
                        self.log_email(f"  Setting recipient to: {recipient}")
                        msg = MIMEMultipart()
                        msg["From"] = smtp_from
                        msg["To"] = recipient
                        msg["Subject"] = subject

                        # Add body
                        msg.attach(MIMEText(body, "plain"))

                        # Add attachment
                        self.log_email(f"  Attaching PDF...")
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
                        server = smtplib.SMTP(smtp_server, smtp_port)

                        self.log_email(f"  [SMTP] Starting TLS...")
                        server.starttls()

                        self.log_email(f"  [SMTP] Logging in as: {smtp_username}...")
                        server.login(smtp_username, smtp_password)

                        self.log_email(f"  [SMTP] Sending message from {smtp_from}...")
                        server.send_message(msg)

                        self.log_email(f"  [SMTP] Closing connection...")
                        server.quit()

                        self.log_email(
                            f"  [OK] Email sent successfully via SMTP from {smtp_from} (used as fallback after Outlook failed)!"
                        )
                else:
                    # Direct SMTP if Outlook disabled
                    import smtplib
                    from email.mime.multipart import MIMEMultipart
                    from email.mime.text import MIMEText
                    from email.mime.base import MIMEBase
                    from email import encoders

                    msg = MIMEMultipart()
                    msg["From"] = smtp_from
                    msg["To"] = recipient
                    msg["Subject"] = subject
                    msg.attach(MIMEText(body, "plain"))

                    with open(attachment_path, "rb") as f:
                        part = MIMEBase("application", "octet-stream")
                        part.set_payload(f.read())
                        encoders.encode_base64(part)
                        part.add_header(
                            "Content-Disposition",
                            f"attachment; filename={attachment_path.name}",
                        )
                        msg.attach(part)

                    server = smtplib.SMTP(smtp_server, smtp_port)
                    server.starttls()
                    server.login(smtp_username, smtp_password)
                    server.send_message(msg)
                    server.quit()

                # Mark as sent in CSV (ONLY if not in test mode)
                if not test_mode:
                    self.mark_as_sent_in_csv(company, person)
                    self.log_email(f"  Updated CSV: marked as sent")
                else:
                    self.log_email(f"  Test mode: NOT updating CSV")

                # Always update email tracker
                self.email_tracker.mark_sent(company, person)

                sent_count += 1
                self.log_email(f"  [OK] SUCCESS: Email sent!")

            except Exception as e:
                failed_count += 1
                error_msg = str(e)
                self.log_email(f"  [ERROR] FAILED: {error_msg}")
                self.log_email(f"  Error type: {type(e).__name__}")

                # Log more details for common errors
                import traceback

                self.log_email(f"  Full error:\n{traceback.format_exc()}")

                # Always update email tracker on failure
                self.email_tracker.mark_failed(company, person)

            # Update progress
            current_idx = idx + 1

            def update_progress():
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
                self.df = pd.read_csv(DATA_FILE)
                self.df.columns = self.df.columns.str.lower().str.strip()
            except Exception as e:
                self.log_email(f"[WARNING] Could not reload CSV: {e}")

            # Update email status display
            self.update_email_status_display()

            # Update statistics in header
            if "reportsent" in self.df.columns:
                self.stats["emails_sent"] = self.df["reportsent"].sum()
            self.update_stats_display()

            # Show final summary dialog
            test_mode_str = " [TEST MODE]" if test_mode else ""
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

    # ==================== Logging Methods ====================

    def log(self, message):
        """Log to system log — thread-safe."""
        import threading

        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_message = f"[{timestamp}] {message}\n"

        # Write to file immediately (thread-safe)
        try:
            with open(LOG_FILE, "a") as f:
                f.write(log_message)
        except:
            pass

        def _update():
            self.system_log.insert(tk.END, log_message)
            self.system_log.see(tk.END)

        if threading.current_thread() is threading.main_thread():
            _update()
        else:
            self.root.after(0, _update)

    def log_gen(self, message):
        """Log to generation log — thread-safe."""
        import threading

        timestamp = datetime.now().strftime("%H:%M:%S")
        log_message = f"[{timestamp}] {message}\n"

        def _update():
            self.gen_log.insert(tk.END, log_message)
            self.gen_log.see(tk.END)

        if threading.current_thread() is threading.main_thread():
            _update()
        else:
            self.root.after(0, _update)
        self.log(message)

    def log_email(self, message):
        """Log to email log — thread-safe."""
        import threading

        timestamp = datetime.now().strftime("%H:%M:%S")
        log_message = f"[{timestamp}] {message}\n"

        def _update():
            self.email_log.insert(tk.END, log_message)
            self.email_log.see(tk.END)

        if threading.current_thread() is threading.main_thread():
            _update()
        else:
            self.root.after(0, _update)
        self.log(message)

    def refresh_logs(self):
        """Refresh system logs"""
        self.system_log.delete("1.0", tk.END)
        try:
            if LOG_FILE.exists():
                with open(LOG_FILE, "r") as f:
                    self.system_log.insert("1.0", f.read())
                self.system_log.see(tk.END)
        except Exception as e:
            self.log(f"Error loading log file: {e}")

    def clear_logs(self):
        """Clear all logs"""
        if messagebox.askyesno("Confirm", "Clear all logs?"):
            self.system_log.delete("1.0", tk.END)
            self.gen_log.delete("1.0", tk.END)
            self.email_log.delete("1.0", tk.END)

            try:
                if LOG_FILE.exists():
                    LOG_FILE.unlink()
            except:
                pass

            self.log("Logs cleared")

    def export_logs(self):
        """Export logs to file"""
        filename = filedialog.asksaveasfilename(
            title="Export Logs",
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
            initialfile=f"resilience_log_{datetime.now().strftime('%Y%m%d_%H%M')}.txt",
        )

        if filename:
            try:
                with open(filename, "w") as f:
                    f.write(self.system_log.get("1.0", tk.END))
                messagebox.showinfo("Success", f"Logs exported to:\n{filename}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export logs:\n{e}")

    # ==================== Utility Methods ====================

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

    def update_time(self):
        """Update time in status bar"""
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.status_time_label.config(text=current_time)
        self.root.after(1000, self.update_time)

    def show_about(self):
        """Show about dialog"""
        from update_checker import _current_version

        about_text = f"""
ResilienceScan Control Center
Version {_current_version()}

A graphical interface for managing supply chain resilience assessments.

Features:
• Data processing and validation
• PDF report generation
• Email distribution
• Real-time monitoring and logging

© 2025 Supply Chain Finance Lectoraat
Hogeschool Windesheim
"""
        messagebox.showinfo("About", about_text)

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
                import subprocess

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
                import subprocess

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


def main():
    """Main entry point"""
    root = tk.Tk()
    app = ResilienceScanGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
