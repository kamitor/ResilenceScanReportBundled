"""
SettingsMixin — startup guard, system check, and dependency-install methods.
"""

import subprocess
import sys
import tkinter as tk
from datetime import datetime
from tkinter import messagebox

from app.app_paths import ROOT_DIR

# gui_system_check is a repo-root module (not inside app/)
from gui_system_check import SystemChecker, setup_status
from dependency_manager import DependencyManager


class SettingsMixin:
    """Mixin providing startup guard, system check, and install helpers."""

    # ------------------------------------------------------------------
    # Startup / polling
    # ------------------------------------------------------------------

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
            missing_str = "\n".join(f"  \u2022 {names[k]}" for k in missing)
            if install_status == "running":
                messagebox.showinfo(
                    "Setup In Progress",
                    "Dependency setup is still running in the background.\n\n"
                    "The following components are not ready yet:\n\n"
                    f"{missing_str}\n\n"
                    "This normally takes 5\u201320 minutes after installation.\n"
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
            self.status_label.config(
                text="Setup complete \u2014 all dependencies ready."
            )
            self.root.after(10_000, lambda: self.status_label.config(text="Ready"))
        elif status == "complete_fail":
            self.status_label.config(
                text="Setup finished with errors \u2014 see System Check."
            )
        elif status == "running":
            self.root.after(30_000, self._poll_setup_completion)
        # else 'unknown' (dev mode / flags cleared) — stop polling silently

    # ------------------------------------------------------------------
    # System check
    # ------------------------------------------------------------------

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
                    report += f"   \u2192 {check['description']}\n"
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

    # ------------------------------------------------------------------
    # Dependency installation
    # ------------------------------------------------------------------

    def install_windows_dependencies(self):
        """Install dependencies on Windows - runs installation/install_dependencies_auto.py"""
        import platform

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
            self.log("Error: Installation script not found")
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
            report += "\nRun the commands above to install R and Quarto.\n"
            report += "Then click 'Check System' to verify installation.\n"

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
