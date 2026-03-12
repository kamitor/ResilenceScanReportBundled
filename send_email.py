import os
import glob
import smtplib
from datetime import datetime
from pathlib import Path
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

import pandas as pd

try:
    import yaml

    _HAS_YAML = True
except ImportError:
    _HAS_YAML = False

from utils.filename_utils import safe_display_name
from utils.path_utils import get_user_base_dir

_user_base = get_user_base_dir()

_CONFIG_FILE = _user_base / "config.yml"

# [OK] CONFIGURATION
CSV_PATH = str(_user_base / "data" / "cleaned_master.csv")
REPORTS_FOLDER = str(_user_base / "reports")
TEST_MODE = True
# Default to a safe placeholder; override via env var RESILIENCESCAN_TEST_EMAIL
# or via config.yml smtp.test_email — never hardcode a real address here.
TEST_EMAIL = os.environ.get("RESILIENCESCAN_TEST_EMAIL", "test@example.com")

# SMTP defaults (overridden by config.yml if present)
SMTP_SERVER = "smtp.office365.com"
SMTP_PORT = 587
SMTP_FROM = "info@resiliencescan.org"
SMTP_USERNAME = ""
SMTP_PASSWORD = ""

# Load SMTP settings from config.yml (if present)
if _HAS_YAML and _CONFIG_FILE.exists():
    try:
        _cfg = yaml.safe_load(_CONFIG_FILE.read_text(encoding="utf-8")) or {}
        _smtp = _cfg.get("smtp", {})
        SMTP_SERVER = _smtp.get("server", SMTP_SERVER)
        SMTP_PORT = int(_smtp.get("port", SMTP_PORT))
        SMTP_FROM = _smtp.get("from_address", SMTP_FROM)
        SMTP_USERNAME = _smtp.get("username", SMTP_USERNAME)
        SMTP_PASSWORD = _smtp.get("password", SMTP_PASSWORD)
    except Exception as _e:
        print(f"[WARN] Could not load config.yml: {_e}")

# IMPORTANT: When using Outlook COM (win32com.client), the system will try
# accounts in priority order:
#   1. info@resiliencescan.org (preferred)
#   2. r.deboer@windesheim.nl
#   3. cg.verhoef@windesheim.nl
#   4. Any available Outlook account
#   5. SMTP as final fallback (if configured above)


def find_report_file(company, person, reports_folder):
    """Find report file matching the new naming format: YYYYMMDD ResilienceScanReport (Company - Person).pdf"""
    # Sanitize names to match filename
    display_company = safe_display_name(company)
    display_person = safe_display_name(person)

    # Try to find file with today's date
    date_str = datetime.now().strftime("%Y%m%d")
    expected_filename = (
        f"{date_str} ResilienceScanReport ({display_company} - {display_person}).pdf"
    )
    expected_path = Path(reports_folder) / expected_filename

    if expected_path.exists():
        return str(expected_path)

    # Try to find file with any date (most recent)
    pattern = f"*ResilienceScanReport ({display_company} - {display_person}).pdf"
    matches = glob.glob(str(Path(reports_folder) / pattern))

    if matches:
        # Return most recent file
        return max(matches, key=os.path.getmtime)

    return None


def send_emails():
    df = pd.read_csv(CSV_PATH)
    df.columns = df.columns.str.lower().str.strip()

    required_cols = {"company_name", "email_address", "name"}
    if not required_cols.issubset(df.columns):
        print(f"[ERROR] Missing one or more required columns: {required_cols}")
        return

    sent_count = 0
    failed_count = 0

    # Try Outlook COM first
    outlook = None
    use_smtp = False

    print("[EMAIL] Trying to connect to Outlook...")
    try:
        import win32com.client as win32

        outlook = win32.Dispatch("Outlook.Application")
        print("[OK] Outlook COM connected successfully!")
    except Exception as e:
        print(f"[WARN] Outlook COM not available: {e}")
        print("[EMAIL] Falling back to SMTP...")
        use_smtp = True

        # Check SMTP configuration
        if not SMTP_FROM or not SMTP_USERNAME or not SMTP_PASSWORD:
            print("[ERROR] SMTP configuration also incomplete!")
            print("   Please either:")
            print("   1. Fix Outlook COM, OR")
            print(
                "   2. Set SMTP_FROM, SMTP_USERNAME, and SMTP_PASSWORD in send_email.py"
            )
            return

    # Only process emails that have reports
    print(f"[SCAN] Scanning reports folder: {REPORTS_FOLDER}")
    available_reports = list(Path(REPORTS_FOLDER).glob("*.pdf"))
    print(f"[OK] Found {len(available_reports)} PDF reports")

    for _, row in df.iterrows():
        company = str(row["company_name"])
        email = row["email_address"]
        name = row.get("name", "there")

        if pd.isna(email) or "@" not in email:
            print(f"[SKIP] {company} - invalid email")
            continue

        # Find report file using new naming format
        attachment_path = find_report_file(company, name, REPORTS_FOLDER)

        if not attachment_path:
            # Skip silently - only send for emails with reports
            continue

        recipient = TEST_EMAIL if TEST_MODE else email

        if TEST_MODE:
            print(f"[TEST] Sending to {TEST_EMAIL} for {company}")
            real_email = email
        else:
            print(f"[SEND] Sending to {email} for {company}")

        # Build email subject
        subject = f"Your Resilience Scan Report – {company}"

        # Build email body
        body = (
            f"Dear {name},\n\n"
            f"Please find attached your resilience scan report for {company}.\n\n"
            "If you have any questions, feel free to reach out.\n\n"
            "Best regards,\n\n"
            "Christiaan Verhoef\n"
            "Windesheim | Value Chain Hackers"
        )

        if TEST_MODE:
            body = (
                f"[TEST MODE]\nThis email was originally intended for: {real_email}\n\n"
                + body
            )

        try:
            if use_smtp:
                # Send via SMTP
                msg = MIMEMultipart()
                msg["From"] = SMTP_FROM
                msg["To"] = recipient
                msg["Subject"] = subject

                # Add body
                msg.attach(MIMEText(body, "plain"))

                # Add PDF attachment
                with open(attachment_path, "rb") as f:
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(f.read())
                    encoders.encode_base64(part)
                    part.add_header(
                        "Content-Disposition",
                        f"attachment; filename={Path(attachment_path).name}",
                    )
                    msg.attach(part)

                # Send via SMTP
                server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=30)
                server.starttls()
                server.login(SMTP_USERNAME, SMTP_PASSWORD)
                server.send_message(msg)
                server.quit()

            else:
                # Send via Outlook COM
                mail = outlook.CreateItem(0)
                mail.To = recipient
                mail.Subject = subject
                mail.Body = body
                mail.Attachments.Add(os.path.abspath(attachment_path))
                mail.Send()

            sent_count += 1
            print(f"   [OK] Sent successfully via {'SMTP' if use_smtp else 'Outlook'}")

        except smtplib.SMTPAuthenticationError as e:
            failed_count += 1
            print(f"   [FAIL] Authentication error (check username/password): {e}")
        except smtplib.SMTPException as e:
            failed_count += 1
            print(f"   [FAIL] SMTP error: {e}")
        except OSError as e:
            failed_count += 1
            print(f"   [FAIL] Network/connection error: {e}")
        except Exception as e:
            failed_count += 1
            print(f"   [FAIL] Unexpected error: {type(e).__name__}: {e}")

    print("\n[OK] Finished!")
    print(f"   Sent: {sent_count}")
    print(f"   Failed: {failed_count}")
    print(f"   Mode: {'TEST' if TEST_MODE else 'LIVE'}")


if __name__ == "__main__":
    send_emails()
