"""
gui_logs.py — LogsMixin: logs tab, log methods, log controls.

Mixed into ResilienceScanGUI in app/main.py.
"""

import threading
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, messagebox, scrolledtext, ttk

from app.app_paths import LOG_FILE


class LogsMixin:
    """Mixin: system-log tab, log(), log_gen(), log_email(), and log controls."""

    def create_logs_tab(self):
        """Create system logs tab"""
        logs_tab = ttk.Frame(self.notebook)
        self.notebook.add(logs_tab, text="\U0001f4cb Logs")

        # Controls
        controls_frame = ttk.Frame(logs_tab)
        controls_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=10, pady=10)

        ttk.Button(
            controls_frame, text="\U0001f504 Refresh Logs", command=self.refresh_logs
        ).grid(row=0, column=0, padx=5)

        ttk.Button(
            controls_frame, text="\U0001f5d1\ufe0f Clear Logs", command=self.clear_logs
        ).grid(row=0, column=1, padx=5)

        ttk.Button(
            controls_frame, text="\U0001f4be Export Logs", command=self.export_logs
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

    def log(self, message):
        """Log to system log — thread-safe."""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_message = f"[{timestamp}] {message}\n"

        # Write to file immediately (thread-safe)
        try:
            with open(LOG_FILE, "a", encoding="utf-8") as f:
                f.write(log_message)
        except Exception:
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
                with open(LOG_FILE, encoding="utf-8") as f:
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
            except Exception:
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
                with open(filename, "w", encoding="utf-8") as f:
                    f.write(self.system_log.get("1.0", tk.END))
                messagebox.showinfo("Success", f"Logs exported to:\n{filename}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export logs:\n{e}")
