"""
SettingsMixin — startup guard, system check, and dependency-install methods.
"""

import subprocess
import sys
import threading
import tkinter as tk
from datetime import datetime
from tkinter import messagebox

from app.app_paths import ROOT_DIR

# gui_system_check is a repo-root module (not inside app/)
from gui_system_check import SystemChecker
from dependency_manager import DependencyManager


class SettingsMixin:
    """Mixin providing startup guard, system check, and install helpers."""

    # ------------------------------------------------------------------
    # Startup / polling
    # ------------------------------------------------------------------

    def _startup_guard(self):
        """Check that R, Quarto, TinyTeX, and R packages are present; show a
        blocking warning dialog if any critical component is missing.

        Dependencies are bundled inside the application.  If any are missing
        the installation is corrupted — the user should reinstall.
        """
        checker = SystemChecker()
        result = checker.check_all()

        critical = ["R", "quarto", "tinytex"]
        missing = [k for k in critical if not result.get(k, {}).get("ok")]

        if missing:
            names = {"R": "R", "quarto": "Quarto", "tinytex": "TinyTeX (tlmgr)"}
            missing_str = "\n".join(f"  \u2022 {names[k]}" for k in missing)
            messagebox.showwarning(
                "Missing Components",
                "The following bundled components were not found:\n\n"
                f"{missing_str}\n\n"
                "The installation may be corrupted.\n"
                "Please reinstall the application.\n\n"
                "You can continue, but generating PDFs will fail.",
            )

        # R packages missing — attempt automatic repair so a partial bundle
        # can self-heal without requiring a full reinstall.
        if not result.get("r_packages", {}).get("ok"):
            self.log(
                "[INFO] R packages missing at startup — attempting automatic repair..."
            )
            self.status_label.config(
                text="Installing missing R packages... (may take a few minutes)"
            )
            self.root.after(500, lambda: self._install_r_packages_now(silent=True))

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

    def _install_r_packages_now(self, silent: bool = False) -> None:
        """Install missing R packages in a background thread.

        Installs to R's default user library (no explicit lib= path) so this
        works even when the frozen app's r-library is in a read-only location.
        R_LIBS_USER is always included in .libPaths() by R at runtime, so
        packages installed here are found by both the system check and by
        quarto render.

        Args:
            silent: if True, suppress the "already installed" success dialog.
        """
        from gui_system_check import _R_PACKAGES, _find_rscript, _refresh_windows_path

        _refresh_windows_path()
        rscript = _find_rscript()
        if not rscript:
            if not silent:
                messagebox.showerror(
                    "R Not Found",
                    "Cannot find Rscript.\n\n"
                    "R must be installed before packages can be installed.\n"
                    "Re-run the installer or install R from https://cran.r-project.org",
                )
            self.log("[ERROR] R package repair failed: Rscript not found on PATH.")
            self.status_label.config(text="Ready")
            return

        pkg_list = ", ".join(f"'{p}'" for p in _R_PACKAGES)
        # Install to user's writable R library (default — no lib= argument).
        # Binary packages are fastest; fall back to source if binary is
        # unavailable for the installed R version.
        script = (
            f"pkgs <- c({pkg_list}); "
            f"bad <- pkgs[!sapply(pkgs, requireNamespace, quietly=TRUE)]; "
            f"if (length(bad) == 0) {{ cat('ALREADY_OK\\n'); quit(status=0) }}; "
            f"cat('Installing', length(bad), 'package(s):', paste(bad, collapse=', '), '\\n'); "
            f"for (p in bad) {{ "
            f"  cat('  ->', p, '\\n'); "
            f"  ok <- tryCatch({{ "
            f"    install.packages(p, repos='https://cloud.r-project.org', type='binary', quiet=FALSE); "
            f"    requireNamespace(p, quietly=TRUE) "
            f"  }}, error=function(e) FALSE); "
            f"  if (!ok) {{ "
            f"    cat('  binary unavailable, trying source:', p, '\\n'); "
            f"    tryCatch( "
            f"      install.packages(p, repos='https://cloud.r-project.org', quiet=FALSE), "
            f"      error=function(e) cat('  ERROR:', conditionMessage(e), '\\n') "
            f"    ) "
            f"  }} "
            f"}}; "
            f"still_bad <- bad[!sapply(bad, requireNamespace, quietly=TRUE)]; "
            f"if (length(still_bad) == 0) cat('SUCCESS\\n') "
            f"else cat('MISSING:', paste(still_bad, collapse=', '), '\\n')"
        )

        self.log(f"[INFO] Running R package install via: {rscript}")
        self.status_label.config(
            text="Installing R packages... (may take a few minutes)"
        )

        def _run() -> None:
            try:
                proc = subprocess.run(
                    [rscript, "--no-save", "-e", script],
                    capture_output=True,
                    text=True,
                    timeout=600,
                )
                output = (proc.stdout + proc.stderr).strip()
                self.root.after(0, lambda: self._r_install_done(output, silent))
            except subprocess.TimeoutExpired:
                self.root.after(0, lambda: self._r_install_done("TIMEOUT", silent))
            except Exception as exc:  # noqa: BLE001
                msg = f"ERROR: {exc}"
                self.root.after(0, lambda m=msg: self._r_install_done(m, silent))

        threading.Thread(target=_run, daemon=True).start()

    def _r_install_done(self, output: str, silent: bool) -> None:
        """Called on the main thread when R package installation finishes."""
        self.status_label.config(text="Ready")
        for line in output.splitlines():
            if line.strip():
                self.log(f"  [R] {line}")

        if "ALREADY_OK" in output:
            self.log("[OK] All R packages were already installed.")
            if not silent:
                messagebox.showinfo(
                    "Packages Ready",
                    "All required R packages are already installed.",
                )
        elif "SUCCESS" in output:
            self.log("[OK] R packages installed successfully.")
            messagebox.showinfo(
                "R Packages Installed",
                "All required R packages were installed successfully.\n\n"
                "You can now generate reports.",
            )
        elif "MISSING:" in output:
            missing = output.split("MISSING:")[-1].strip()
            self.log(f"[WARN] Some packages could not be installed: {missing}")
            messagebox.showwarning(
                "Some R Packages Failed",
                f"The following packages could not be installed:\n\n{missing}\n\n"
                "Possible causes:\n"
                "  \u2022 No internet connection\n"
                "  \u2022 Package not available as binary for your R version\n"
                "  \u2022 R version too old or too new\n\n"
                "Use the System Check button to retry or check the Logs tab.",
            )
        elif "TIMEOUT" in output:
            self.log("[WARN] R package install timed out after 10 minutes.")
            messagebox.showwarning(
                "Installation Timeout",
                "R package installation timed out after 10 minutes.\n\n"
                "Check your internet connection and use System Check to retry.",
            )
        else:
            self.log(
                f"[WARN] R install finished with unexpected output: {output[:300]}"
            )
            messagebox.showwarning(
                "Installation Result Unclear",
                "R package installation finished but the result is uncertain.\n\n"
                "Use the System Check button to verify the current state.",
            )

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
