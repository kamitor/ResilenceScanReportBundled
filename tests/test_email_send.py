"""
test_email_send.py — tests for send_email.py helpers.

Covers:
- find_report_file() with today's date, any-date fallback, no match
- TEST_EMAIL default is test@example.com (not a real address)
- send_emails() SMTP path with mocked smtplib (success + auth error)
- email_tracker state after mark_sent / mark_failed / mark_pending
"""

import pathlib
import smtplib
import sys
from datetime import datetime
from unittest.mock import MagicMock, patch

import pandas as pd

ROOT = pathlib.Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import send_email  # noqa: E402
from email_tracker import EmailTracker  # noqa: E402


# ---------------------------------------------------------------------------
# TEST_EMAIL default
# ---------------------------------------------------------------------------


def test_test_email_default_is_placeholder(monkeypatch):
    """TEST_EMAIL must default to test@example.com, not a real address."""
    monkeypatch.delenv("RESILIENCESCAN_TEST_EMAIL", raising=False)
    import importlib

    importlib.reload(send_email)
    assert send_email.TEST_EMAIL == "test@example.com"


def test_test_email_overridable_via_env(monkeypatch):
    """TEST_EMAIL can be overridden via env var."""
    monkeypatch.setenv("RESILIENCESCAN_TEST_EMAIL", "override@example.com")
    import importlib

    importlib.reload(send_email)
    assert send_email.TEST_EMAIL == "override@example.com"


# ---------------------------------------------------------------------------
# find_report_file
# ---------------------------------------------------------------------------


def test_find_report_file_today(tmp_path):
    """find_report_file returns today's file when it exists."""
    date_str = datetime.now().strftime("%Y%m%d")
    fname = f"{date_str} ResilienceScanReport (Acme Corp - Alice Smith).pdf"
    (tmp_path / fname).write_bytes(b"%PDF-1.4")

    result = send_email.find_report_file("Acme Corp", "Alice Smith", str(tmp_path))
    assert result is not None
    assert "Acme Corp" in result
    assert "Alice Smith" in result


def test_find_report_file_any_date(tmp_path):
    """find_report_file falls back to any-date match when today's file absent."""
    old_fname = "20230101 ResilienceScanReport (Acme Corp - Alice Smith).pdf"
    (tmp_path / old_fname).write_bytes(b"%PDF-1.4")

    result = send_email.find_report_file("Acme Corp", "Alice Smith", str(tmp_path))
    assert result is not None
    assert "Acme Corp" in result


def test_find_report_file_no_match(tmp_path):
    """find_report_file returns None when no matching PDF is present."""
    result = send_email.find_report_file("NoSuch Corp", "Nobody", str(tmp_path))
    assert result is None


def test_find_report_file_sanitises_slash(tmp_path):
    """find_report_file handles company names with slashes."""
    # Company with slash should be sanitised to dash in filename
    fname = "20230101 ResilienceScanReport (A-B Corp - Alice Smith).pdf"
    (tmp_path / fname).write_bytes(b"%PDF-1.4")

    result = send_email.find_report_file("A/B Corp", "Alice Smith", str(tmp_path))
    assert result is not None


# ---------------------------------------------------------------------------
# send_emails() via mocked SMTP
# ---------------------------------------------------------------------------


def _write_csv(path, rows):
    pd.DataFrame(rows).to_csv(path, index=False)


def test_send_emails_smtp_success(tmp_path, monkeypatch):
    """send_emails() calls smtplib.SMTP.send_message for each matched report."""
    # Prepare CSV
    csv_path = tmp_path / "cleaned_master.csv"
    _write_csv(
        csv_path,
        [{"company_name": "Acme", "name": "Alice Smith", "email_address": "a@x.com"}],
    )

    # Prepare a fake PDF
    date_str = datetime.now().strftime("%Y%m%d")
    pdf_name = f"{date_str} ResilienceScanReport (Acme - Alice Smith).pdf"
    reports_dir = tmp_path / "reports"
    reports_dir.mkdir()
    (reports_dir / pdf_name).write_bytes(b"%PDF-1.4")

    monkeypatch.setattr(send_email, "CSV_PATH", str(csv_path))
    monkeypatch.setattr(send_email, "REPORTS_FOLDER", str(reports_dir))
    monkeypatch.setattr(send_email, "TEST_MODE", False)
    monkeypatch.setattr(send_email, "SMTP_FROM", "from@example.com")
    monkeypatch.setattr(send_email, "SMTP_USERNAME", "user")
    monkeypatch.setattr(send_email, "SMTP_PASSWORD", "pass")

    mock_smtp_instance = MagicMock()
    mock_smtp_cls = MagicMock(return_value=mock_smtp_instance)

    with patch("send_email.smtplib.SMTP", mock_smtp_cls):
        # Also patch win32com so it raises ImportError (forces SMTP path)
        with patch.dict(sys.modules, {"win32com": None, "win32com.client": None}):
            send_email.send_emails()

    mock_smtp_instance.send_message.assert_called_once()
    mock_smtp_instance.quit.assert_called_once()


def test_send_emails_smtp_auth_error(tmp_path, monkeypatch, capsys):
    """send_emails() prints auth error and continues when SMTPAuthenticationError."""
    csv_path = tmp_path / "cleaned_master.csv"
    _write_csv(
        csv_path,
        [{"company_name": "Acme", "name": "Alice Smith", "email_address": "a@x.com"}],
    )

    date_str = datetime.now().strftime("%Y%m%d")
    pdf_name = f"{date_str} ResilienceScanReport (Acme - Alice Smith).pdf"
    reports_dir = tmp_path / "reports"
    reports_dir.mkdir()
    (reports_dir / pdf_name).write_bytes(b"%PDF-1.4")

    monkeypatch.setattr(send_email, "CSV_PATH", str(csv_path))
    monkeypatch.setattr(send_email, "REPORTS_FOLDER", str(reports_dir))
    monkeypatch.setattr(send_email, "TEST_MODE", False)
    monkeypatch.setattr(send_email, "SMTP_FROM", "from@example.com")
    monkeypatch.setattr(send_email, "SMTP_USERNAME", "user")
    monkeypatch.setattr(send_email, "SMTP_PASSWORD", "pass")

    mock_smtp_instance = MagicMock()
    mock_smtp_instance.login.side_effect = smtplib.SMTPAuthenticationError(
        535, b"Auth failed"
    )
    mock_smtp_cls = MagicMock(return_value=mock_smtp_instance)

    with patch("send_email.smtplib.SMTP", mock_smtp_cls):
        with patch.dict(sys.modules, {"win32com": None, "win32com.client": None}):
            send_email.send_emails()

    captured = capsys.readouterr()
    assert "Authentication error" in captured.out or "FAIL" in captured.out


def test_send_emails_skips_invalid_email(tmp_path, monkeypatch, capsys):
    """send_emails() skips rows with invalid/missing email addresses."""
    csv_path = tmp_path / "cleaned_master.csv"
    _write_csv(
        csv_path,
        [
            {
                "company_name": "Acme",
                "name": "Alice Smith",
                "email_address": "not-an-email",
            }
        ],
    )
    reports_dir = tmp_path / "reports"
    reports_dir.mkdir()

    monkeypatch.setattr(send_email, "CSV_PATH", str(csv_path))
    monkeypatch.setattr(send_email, "REPORTS_FOLDER", str(reports_dir))
    monkeypatch.setattr(send_email, "TEST_MODE", False)
    monkeypatch.setattr(send_email, "SMTP_FROM", "from@example.com")
    monkeypatch.setattr(send_email, "SMTP_USERNAME", "user")
    monkeypatch.setattr(send_email, "SMTP_PASSWORD", "pass")

    with patch.dict(sys.modules, {"win32com": None, "win32com.client": None}):
        send_email.send_emails()

    captured = capsys.readouterr()
    assert "SKIP" in captured.out or "invalid" in captured.out.lower()


# ---------------------------------------------------------------------------
# EmailTracker state management
# ---------------------------------------------------------------------------


def test_tracker_mark_sent(tmp_path, monkeypatch):
    """mark_sent records the company+person as sent."""
    import email_tracker as et

    monkeypatch.setattr(et, "_TRACKER_FILE", tmp_path / "email_tracker.json")
    monkeypatch.setattr(et, "_DATA_DIR", tmp_path)

    tracker = EmailTracker()
    tracker.mark_sent("Acme", "Alice")
    stats = tracker.get_statistics()
    assert stats["sent"] >= 1


def test_tracker_mark_failed(tmp_path, monkeypatch):
    """mark_failed records the company+person as failed."""
    import email_tracker as et

    monkeypatch.setattr(et, "_TRACKER_FILE", tmp_path / "email_tracker.json")
    monkeypatch.setattr(et, "_DATA_DIR", tmp_path)

    tracker = EmailTracker()
    tracker.mark_failed("Acme", "Alice")
    stats = tracker.get_statistics()
    assert stats["failed"] >= 1


def test_tracker_mark_pending(tmp_path, monkeypatch):
    """mark_pending resets a previously sent entry to pending."""
    import email_tracker as et

    monkeypatch.setattr(et, "_TRACKER_FILE", tmp_path / "email_tracker.json")
    monkeypatch.setattr(et, "_DATA_DIR", tmp_path)

    tracker = EmailTracker()
    tracker.mark_sent("Acme", "Alice")
    tracker.mark_pending("Acme", "Alice")
    stats = tracker.get_statistics()
    assert stats["sent"] == 0
    assert stats["pending"] >= 1


def test_tracker_total_equals_sum(tmp_path, monkeypatch):
    """total always equals sent + pending + failed."""
    import email_tracker as et

    monkeypatch.setattr(et, "_TRACKER_FILE", tmp_path / "email_tracker.json")
    monkeypatch.setattr(et, "_DATA_DIR", tmp_path)

    tracker = EmailTracker()
    tracker.mark_sent("A", "1")
    tracker.mark_failed("B", "2")
    tracker.mark_pending("C", "3")

    stats = tracker.get_statistics()
    assert stats["total"] == stats["sent"] + stats["pending"] + stats["failed"]
