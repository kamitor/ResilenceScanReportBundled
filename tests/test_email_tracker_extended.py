"""
test_email_tracker_extended.py — comprehensive tests for EmailTracker.

Covers: persistence across instances, import_from_csv edge cases, concurrent
thread safety, Unicode names, whitespace stripping, corrupted JSON recovery,
and the full public API lifecycle.
"""

import csv
import json
import pathlib
import sys
import threading

import pytest

ROOT = pathlib.Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import email_tracker as et  # noqa: E402
from email_tracker import EmailTracker  # noqa: E402


# ---------------------------------------------------------------------------
# Helper: patch module-level paths so tests are isolated in tmp_path
# ---------------------------------------------------------------------------


@pytest.fixture()
def tracker(monkeypatch, tmp_path):
    """Return a fresh EmailTracker backed by a tmp_path directory."""
    monkeypatch.setattr(et, "_TRACKER_FILE", tmp_path / "email_tracker.json")
    monkeypatch.setattr(et, "_DATA_DIR", tmp_path)
    return EmailTracker()


def _csv(tmp_path, rows):
    """Write rows (list of dicts) to a CSV file and return the path."""
    path = tmp_path / "test.csv"
    if rows:
        fieldnames = list(rows[0].keys())
        with path.open("w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(rows)
    else:
        path.write_text("company_name,name,email_address\n", encoding="utf-8")
    return path


# ---------------------------------------------------------------------------
# Construction and persistence
# ---------------------------------------------------------------------------


def test_fresh_tracker_is_empty(tracker):
    """A brand-new tracker starts with zero recipients."""
    assert tracker.get_statistics() == {
        "total": 0,
        "sent": 0,
        "pending": 0,
        "failed": 0,
    }


def test_load_from_nonexistent_file(monkeypatch, tmp_path):
    """EmailTracker initialises cleanly when the JSON file does not yet exist."""
    monkeypatch.setattr(et, "_TRACKER_FILE", tmp_path / "missing.json")
    monkeypatch.setattr(et, "_DATA_DIR", tmp_path)
    t = EmailTracker()
    assert t.get_all() == []


def test_load_from_corrupted_json(monkeypatch, tmp_path):
    """EmailTracker recovers silently from a corrupted JSON file."""
    tracker_file = tmp_path / "email_tracker.json"
    tracker_file.write_text("NOT VALID JSON {{{", encoding="utf-8")
    monkeypatch.setattr(et, "_TRACKER_FILE", tracker_file)
    monkeypatch.setattr(et, "_DATA_DIR", tmp_path)
    t = EmailTracker()
    assert t.get_all() == []


def test_persistence_across_instances(monkeypatch, tmp_path):
    """Data saved by one instance is loaded by the next."""
    monkeypatch.setattr(et, "_TRACKER_FILE", tmp_path / "email_tracker.json")
    monkeypatch.setattr(et, "_DATA_DIR", tmp_path)

    t1 = EmailTracker()
    t1.mark_sent("Acme", "Alice")
    t1.mark_failed("Beta", "Bob")

    t2 = EmailTracker()
    stats = t2.get_statistics()
    assert stats["sent"] == 1
    assert stats["failed"] == 1


def test_save_creates_data_dir(monkeypatch, tmp_path):
    """mark_sent creates the data directory if it does not exist."""
    nested = tmp_path / "deep" / "nested"
    monkeypatch.setattr(et, "_TRACKER_FILE", nested / "email_tracker.json")
    monkeypatch.setattr(et, "_DATA_DIR", nested)
    t = EmailTracker()
    t.mark_sent("X", "Y")
    assert (nested / "email_tracker.json").exists()


# ---------------------------------------------------------------------------
# mark_sent
# ---------------------------------------------------------------------------


def test_mark_sent_creates_new_entry(tracker):
    """mark_sent creates a recipient entry if it doesn't already exist."""
    tracker.mark_sent("Acme", "Alice")
    entries = tracker.get_all()
    assert len(entries) == 1
    assert entries[0]["status"] == "sent"


def test_mark_sent_updates_existing(tracker):
    """mark_sent updates an existing pending entry to sent."""
    tracker.mark_pending(
        "Acme", "Alice"
    )  # creates pending entry... no, mark_pending only updates if exists
    tracker.mark_failed("Acme", "Alice")  # creates failed entry
    tracker.mark_sent("Acme", "Alice")
    assert tracker.get_statistics()["sent"] == 1
    assert tracker.get_statistics()["failed"] == 0


def test_mark_sent_sets_sent_date(tracker):
    """mark_sent stores a non-None sent_date timestamp."""
    tracker.mark_sent("Acme", "Alice")
    entry = tracker.get_all()[0]
    assert entry["sent_date"] is not None
    assert len(entry["sent_date"]) > 0


def test_mark_sent_twice_does_not_duplicate(tracker):
    """Calling mark_sent twice for the same recipient does not create duplicates."""
    tracker.mark_sent("Acme", "Alice")
    tracker.mark_sent("Acme", "Alice")
    assert tracker.get_statistics()["total"] == 1


# ---------------------------------------------------------------------------
# mark_failed
# ---------------------------------------------------------------------------


def test_mark_failed_creates_entry(tracker):
    """mark_failed creates a recipient entry marked as failed."""
    tracker.mark_failed("Acme", "Alice")
    assert tracker.get_statistics()["failed"] == 1


def test_mark_failed_no_sent_date(tracker):
    """A failed entry has no sent_date."""
    tracker.mark_failed("Acme", "Alice")
    entry = tracker.get_all()[0]
    assert entry.get("sent_date") is None


def test_mark_failed_overrides_sent(tracker):
    """mark_failed downgrades a previously sent entry to failed."""
    tracker.mark_sent("Acme", "Alice")
    tracker.mark_failed("Acme", "Alice")
    stats = tracker.get_statistics()
    assert stats["sent"] == 0
    assert stats["failed"] == 1


# ---------------------------------------------------------------------------
# mark_pending
# ---------------------------------------------------------------------------


def test_mark_pending_resets_failed(tracker):
    """mark_pending resets a failed entry back to pending."""
    tracker.mark_failed("Acme", "Alice")
    tracker.mark_pending("Acme", "Alice")
    stats = tracker.get_statistics()
    assert stats["failed"] == 0
    assert stats["pending"] == 1


def test_mark_pending_clears_sent_date(tracker):
    """mark_pending clears the sent_date."""
    tracker.mark_sent("Acme", "Alice")
    tracker.mark_pending("Acme", "Alice")
    entry = tracker.get_all()[0]
    assert entry["sent_date"] is None


def test_mark_pending_nonexistent_is_noop(tracker):
    """mark_pending on an unknown recipient does NOT create a new entry."""
    tracker.mark_pending("Nobody", "Unknown")
    assert tracker.get_statistics()["total"] == 0


# ---------------------------------------------------------------------------
# get_statistics consistency
# ---------------------------------------------------------------------------


def test_statistics_total_always_equals_sum(tracker):
    """total == sent + pending + failed for any mix of statuses."""
    tracker.mark_sent("A", "1")
    tracker.mark_failed("B", "2")
    tracker.mark_sent("C", "3")
    tracker.mark_failed("D", "4")
    # "E" imported via import_from_csv would be pending; use direct manipulation
    tracker.mark_sent("E", "5")
    tracker.mark_pending("E", "5")

    stats = tracker.get_statistics()
    assert stats["total"] == stats["sent"] + stats["pending"] + stats["failed"]


def test_get_all_returns_list_of_dicts(tracker):
    """get_all() returns a list of dict entries with expected keys."""
    tracker.mark_sent("Acme", "Alice")
    entries = tracker.get_all()
    assert isinstance(entries, list)
    assert "key" in entries[0]
    assert "status" in entries[0]
    assert "company" in entries[0]
    assert "person" in entries[0]


# ---------------------------------------------------------------------------
# import_from_csv
# ---------------------------------------------------------------------------


def test_import_csv_basic(monkeypatch, tmp_path):
    """import_from_csv imports rows as pending entries."""
    monkeypatch.setattr(et, "_TRACKER_FILE", tmp_path / "email_tracker.json")
    monkeypatch.setattr(et, "_DATA_DIR", tmp_path)
    t = EmailTracker()

    path = _csv(
        tmp_path,
        [
            {"company_name": "Acme", "name": "Alice", "email_address": "a@x.com"},
            {"company_name": "Beta", "name": "Bob", "email_address": "b@x.com"},
        ],
    )
    imported, skipped = t.import_from_csv(str(path))
    assert imported == 2
    assert skipped == 0
    assert t.get_statistics()["pending"] == 2


def test_import_csv_skips_existing(monkeypatch, tmp_path):
    """import_from_csv does not add duplicate entries."""
    monkeypatch.setattr(et, "_TRACKER_FILE", tmp_path / "email_tracker.json")
    monkeypatch.setattr(et, "_DATA_DIR", tmp_path)
    t = EmailTracker()

    path = _csv(
        tmp_path,
        [{"company_name": "Acme", "name": "Alice", "email_address": "a@x.com"}],
    )
    t.import_from_csv(str(path))
    imported, skipped = t.import_from_csv(str(path))
    assert imported == 0
    assert skipped == 1
    assert t.get_statistics()["total"] == 1


def test_import_csv_reportsent_true(monkeypatch, tmp_path):
    """Rows with reportsent=True are imported with status='sent'."""
    monkeypatch.setattr(et, "_TRACKER_FILE", tmp_path / "email_tracker.json")
    monkeypatch.setattr(et, "_DATA_DIR", tmp_path)
    t = EmailTracker()

    path = _csv(
        tmp_path,
        [
            {
                "company_name": "Acme",
                "name": "Alice",
                "email_address": "a@x.com",
                "reportsent": True,
            },
        ],
    )
    t.import_from_csv(str(path))
    assert t.get_statistics()["sent"] == 1
    assert t.get_statistics()["pending"] == 0


def test_import_csv_missing_company_skipped(monkeypatch, tmp_path):
    """Rows where company_name is blank are skipped during import."""
    monkeypatch.setattr(et, "_TRACKER_FILE", tmp_path / "email_tracker.json")
    monkeypatch.setattr(et, "_DATA_DIR", tmp_path)
    t = EmailTracker()

    path = _csv(
        tmp_path, [{"company_name": "", "name": "Alice", "email_address": "a@x.com"}]
    )
    imported, skipped = t.import_from_csv(str(path))
    assert imported == 0
    assert skipped == 1


def test_import_csv_missing_file_returns_zero(monkeypatch, tmp_path):
    """import_from_csv returns (0, 0) gracefully when file doesn't exist."""
    monkeypatch.setattr(et, "_TRACKER_FILE", tmp_path / "email_tracker.json")
    monkeypatch.setattr(et, "_DATA_DIR", tmp_path)
    t = EmailTracker()
    imported, skipped = t.import_from_csv(str(tmp_path / "nonexistent.csv"))
    assert imported == 0
    assert skipped == 0


def test_import_csv_updates_missing_email(monkeypatch, tmp_path):
    """If an existing entry has no email, import_from_csv fills it in."""
    monkeypatch.setattr(et, "_TRACKER_FILE", tmp_path / "email_tracker.json")
    monkeypatch.setattr(et, "_DATA_DIR", tmp_path)
    t = EmailTracker()

    # First import: no email
    path1 = _csv(
        tmp_path, [{"company_name": "Acme", "name": "Alice", "email_address": ""}]
    )
    t.import_from_csv(str(path1))
    entry_before = t.get_all()[0]
    assert entry_before["email"] == ""

    # Second import: same company/person but now has email
    tmp_path2 = tmp_path / "v2"
    tmp_path2.mkdir()
    path2 = _csv(
        tmp_path2,
        [{"company_name": "Acme", "name": "Alice", "email_address": "alice@acme.com"}],
    )
    t.import_from_csv(str(path2))

    entry_after = t.get_all()[0]
    assert entry_after["email"] == "alice@acme.com"


# ---------------------------------------------------------------------------
# Key normalisation (whitespace stripping)
# ---------------------------------------------------------------------------


def test_key_strips_whitespace(tracker):
    """Company and person names with leading/trailing spaces map to the same key."""
    tracker.mark_sent("  Acme  ", "  Alice  ")
    tracker.mark_sent("Acme", "Alice")  # should be the same entry
    assert tracker.get_statistics()["total"] == 1


# ---------------------------------------------------------------------------
# Unicode names
# ---------------------------------------------------------------------------


def test_unicode_company_name(tracker):
    """EmailTracker handles Unicode company names (e.g. accented characters)."""
    tracker.mark_sent("Réseau SA", "Ångström AB")
    entries = tracker.get_all()
    assert len(entries) == 1
    assert entries[0]["company"] == "Réseau SA"


def test_unicode_persistence(monkeypatch, tmp_path):
    """Unicode names survive a save/reload cycle."""
    monkeypatch.setattr(et, "_TRACKER_FILE", tmp_path / "email_tracker.json")
    monkeypatch.setattr(et, "_DATA_DIR", tmp_path)

    t1 = EmailTracker()
    t1.mark_sent("日本語株式会社", "田中太郎")

    t2 = EmailTracker()
    entries = t2.get_all()
    assert len(entries) == 1
    assert entries[0]["company"] == "日本語株式会社"


# ---------------------------------------------------------------------------
# Thread safety
# ---------------------------------------------------------------------------


def test_concurrent_mark_sent_no_data_loss(monkeypatch, tmp_path):
    """Concurrent mark_sent calls from multiple threads produce consistent totals."""
    monkeypatch.setattr(et, "_TRACKER_FILE", tmp_path / "email_tracker.json")
    monkeypatch.setattr(et, "_DATA_DIR", tmp_path)
    t = EmailTracker()

    n_threads = 20
    errors = []

    def worker(i):
        try:
            t.mark_sent(f"Company{i}", f"Person{i}")
        except Exception as exc:
            errors.append(exc)

    threads = [threading.Thread(target=worker, args=(i,)) for i in range(n_threads)]
    for th in threads:
        th.start()
    for th in threads:
        th.join()

    assert errors == [], f"Exceptions in threads: {errors}"
    assert t.get_statistics()["total"] == n_threads


def test_concurrent_mixed_operations(monkeypatch, tmp_path):
    """Mixed concurrent operations (sent/failed/pending) do not corrupt the store."""
    monkeypatch.setattr(et, "_TRACKER_FILE", tmp_path / "email_tracker.json")
    monkeypatch.setattr(et, "_DATA_DIR", tmp_path)
    t = EmailTracker()

    def worker(i):
        t.mark_sent(f"C{i}", f"P{i}")
        t.mark_failed(f"C{i}", f"P{i}")
        t.mark_pending(f"C{i}", f"P{i}")

    threads = [threading.Thread(target=worker, args=(i,)) for i in range(10)]
    for th in threads:
        th.start()
    for th in threads:
        th.join()

    stats = t.get_statistics()
    assert stats["total"] == stats["sent"] + stats["pending"] + stats["failed"]


def test_json_file_is_valid_after_concurrent_writes(monkeypatch, tmp_path):
    """The persisted JSON file remains valid after concurrent writes."""
    tracker_file = tmp_path / "email_tracker.json"
    monkeypatch.setattr(et, "_TRACKER_FILE", tracker_file)
    monkeypatch.setattr(et, "_DATA_DIR", tmp_path)
    t = EmailTracker()

    threads = [
        threading.Thread(target=t.mark_sent, args=(f"Co{i}", f"Pe{i}"))
        for i in range(15)
    ]
    for th in threads:
        th.start()
    for th in threads:
        th.join()

    data = json.loads(tracker_file.read_text(encoding="utf-8"))
    assert "recipients" in data
    assert isinstance(data["recipients"], list)
