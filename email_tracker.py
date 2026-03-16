"""
email_tracker.py — tracks per-recipient email send status.

State is persisted to email_tracker.json in the user data directory so that
tracking survives app restarts and Excel re-imports.

Called by the GUI via:
  import_from_csv(path)            -> (imported: int, skipped: int)
  get_statistics()                 -> {"total", "sent", "pending", "failed"}
  mark_sent(company, person)
  mark_failed(company, person)
  mark_pending(company, person)
"""

import json
from datetime import datetime

from utils.path_utils import get_user_base_dir

_user_base = get_user_base_dir()

_DATA_DIR = _user_base / "data"
_TRACKER_FILE = _DATA_DIR / "email_tracker.json"


def _key(company: str, person: str) -> str:
    return f"{company.strip()}|{person.strip()}"


class EmailTracker:
    """Tracks per-recipient email send status, persisted to email_tracker.json.

    Each recipient is stored as::

        {
            "key":       "<company>|<person>",
            "company":   str,
            "person":    str,
            "email":     str,
            "status":    "pending" | "sent" | "failed",
            "sent_date": str | None,
        }
    """

    def __init__(self) -> None:
        self._recipients: dict[str, dict] = {}
        self._load()

    # ------------------------------------------------------------------ I/O

    def _load(self) -> None:
        if _TRACKER_FILE.exists():
            try:
                data = json.loads(_TRACKER_FILE.read_text(encoding="utf-8"))
                self._recipients = {r["key"]: r for r in data.get("recipients", [])}
            except (OSError, ValueError):
                self._recipients = {}

    def _save(self) -> None:
        _DATA_DIR.mkdir(parents=True, exist_ok=True)
        payload = {"recipients": list(self._recipients.values())}
        _TRACKER_FILE.write_text(
            json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8"
        )

    # ------------------------------------------------------------------ public API

    def import_from_csv(self, path: str) -> tuple[int, int]:
        """Import recipients from cleaned_master.csv.

        Adds each row that has company_name + name as a recipient (if not
        already tracked).  Rows where reportsent=True are imported with
        status='sent'.  Returns (imported, skipped).
        """
        try:
            import pandas as pd

            df = pd.read_csv(path, low_memory=False, encoding="utf-8")
            df.columns = df.columns.str.lower().str.strip()
        except (OSError, ValueError) as e:
            print(f"[email_tracker] Cannot read CSV: {e}")
            return 0, 0

        imported = 0
        skipped = 0

        for _, row in df.iterrows():
            company = str(row.get("company_name", "")).strip()
            person = str(row.get("name", "")).strip()
            if not company or company.lower() == "nan" or not person:
                skipped += 1
                continue

            email = str(row.get("email_address", "")).strip()
            if email.lower() == "nan":
                email = ""

            k = _key(company, person)
            if k in self._recipients:
                # Update blank email if we now have one
                if not self._recipients[k].get("email") and email:
                    self._recipients[k]["email"] = email
                skipped += 1
                continue

            is_sent = bool(row.get("reportsent", False))
            self._recipients[k] = {
                "key": k,
                "company": company,
                "person": person,
                "email": email,
                "status": "sent" if is_sent else "pending",
                "sent_date": None,
            }
            imported += 1

        self._save()
        return imported, skipped

    def get_statistics(self) -> dict:
        """Return aggregate send statistics."""
        total = len(self._recipients)
        sent = sum(1 for r in self._recipients.values() if r["status"] == "sent")
        failed = sum(1 for r in self._recipients.values() if r["status"] == "failed")
        pending = total - sent - failed
        return {"total": total, "sent": sent, "pending": pending, "failed": failed}

    def mark_sent(self, company: str, person: str) -> None:
        """Mark a recipient as successfully sent."""
        k = _key(company, person)
        entry = self._recipients.get(k) or {
            "key": k,
            "company": company.strip(),
            "person": person.strip(),
            "email": "",
        }
        entry["status"] = "sent"
        entry["sent_date"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self._recipients[k] = entry
        self._save()

    def mark_failed(self, company: str, person: str) -> None:
        """Mark a recipient as failed."""
        k = _key(company, person)
        entry = self._recipients.get(k) or {
            "key": k,
            "company": company.strip(),
            "person": person.strip(),
            "email": "",
            "sent_date": None,
        }
        entry["status"] = "failed"
        self._recipients[k] = entry
        self._save()

    def mark_pending(self, company: str, person: str) -> None:
        """Reset a recipient to pending."""
        k = _key(company, person)
        if k in self._recipients:
            self._recipients[k]["status"] = "pending"
            self._recipients[k]["sent_date"] = None
            self._save()

    def get_all(self) -> list[dict]:
        """Return all recipients as a list (for UI population)."""
        return list(self._recipients.values())
