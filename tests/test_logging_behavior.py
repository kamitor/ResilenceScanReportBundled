"""
test_logging_behavior.py — tests for the log-file system.

Verifies that:
- LOG_FILE is located inside the data root
- File-level log writes are thread-safe under concurrent access
- Log entries survive a read-back (not truncated or corrupted)
- Unicode content is handled without error
"""

import pathlib
import sys
import threading

ROOT = pathlib.Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))


# ---------------------------------------------------------------------------
# LOG_FILE location
# ---------------------------------------------------------------------------


def test_log_file_is_in_data_root():
    """LOG_FILE constant is a child of the data root directory."""
    import app.app_paths as ap

    assert ap.LOG_FILE.parent == ap._DATA_ROOT


def test_log_file_name():
    """LOG_FILE has the expected filename."""
    import app.app_paths as ap

    assert ap.LOG_FILE.name == "gui_log.txt"


# ---------------------------------------------------------------------------
# Low-level thread-safe file writes (mirrors what LogsMixin._write_log does)
# ---------------------------------------------------------------------------


def _make_log_writer(log_file: pathlib.Path, lock: threading.Lock):
    """Return a write function that mimics the app's thread-safe log helper."""

    def write(message: str) -> None:
        with lock:
            with log_file.open("a", encoding="utf-8") as fh:
                fh.write(message + "\n")

    return write


def test_sequential_writes_appear_in_file(tmp_path):
    """Sequential log writes are all present in the output file."""
    log_file = tmp_path / "gui_log.txt"
    lock = threading.Lock()
    write = _make_log_writer(log_file, lock)

    messages = [f"line {i}" for i in range(20)]
    for m in messages:
        write(m)

    content = log_file.read_text(encoding="utf-8")
    for m in messages:
        assert m in content


def test_concurrent_writes_no_corruption(tmp_path):
    """Concurrent writes from many threads do not corrupt the log file."""
    log_file = tmp_path / "gui_log.txt"
    lock = threading.Lock()
    write = _make_log_writer(log_file, lock)

    n_threads = 30
    lines_per_thread = 10

    def worker(thread_id: int) -> None:
        for i in range(lines_per_thread):
            write(f"thread={thread_id} line={i}")

    threads = [threading.Thread(target=worker, args=(t,)) for t in range(n_threads)]
    for th in threads:
        th.start()
    for th in threads:
        th.join()

    lines = log_file.read_text(encoding="utf-8").splitlines()
    assert len(lines) == n_threads * lines_per_thread


def test_unicode_log_message(tmp_path):
    """Log files handle Unicode characters without encoding errors."""
    log_file = tmp_path / "gui_log.txt"
    lock = threading.Lock()
    write = _make_log_writer(log_file, lock)

    write("Bédrijf: 日本語株式会社")
    write("Person: Åsa Lindström")

    content = log_file.read_text(encoding="utf-8")
    assert "日本語株式会社" in content
    assert "Åsa Lindström" in content


def test_log_file_created_on_first_write(tmp_path):
    """The log file is created automatically on the first write."""
    log_file = tmp_path / "subdir" / "gui_log.txt"
    log_file.parent.mkdir(parents=True, exist_ok=True)
    lock = threading.Lock()
    write = _make_log_writer(log_file, lock)

    assert not log_file.exists()
    write("hello")
    assert log_file.exists()


def test_log_appends_not_overwrites(tmp_path):
    """Subsequent writes append to an existing log file rather than overwriting it."""
    log_file = tmp_path / "gui_log.txt"
    log_file.write_text("first entry\n", encoding="utf-8")
    lock = threading.Lock()
    write = _make_log_writer(log_file, lock)

    write("second entry")
    content = log_file.read_text(encoding="utf-8")
    assert "first entry" in content
    assert "second entry" in content


def test_log_long_message(tmp_path):
    """A very long log message (>1 MB) is written and read back without truncation."""
    log_file = tmp_path / "gui_log.txt"
    lock = threading.Lock()
    write = _make_log_writer(log_file, lock)

    long_msg = "x" * (1024 * 1024)  # 1 MB
    write(long_msg)
    content = log_file.read_text(encoding="utf-8")
    assert long_msg in content


# ---------------------------------------------------------------------------
# Log lock prevents interleaving under stress
# ---------------------------------------------------------------------------


def test_log_lines_are_not_interleaved(tmp_path):
    """Each log line is written atomically — no partial line interleaving."""
    log_file = tmp_path / "gui_log.txt"
    lock = threading.Lock()
    write = _make_log_writer(log_file, lock)

    marker = "COMPLETE_LINE"
    n = 50

    def worker():
        for _ in range(n):
            write(marker)

    threads = [threading.Thread(target=worker) for _ in range(5)]
    for th in threads:
        th.start()
    for th in threads:
        th.join()

    lines = log_file.read_text(encoding="utf-8").splitlines()
    for line in lines:
        assert line == marker, f"Interleaved or corrupted line: {line!r}"


# ---------------------------------------------------------------------------
# app.app_paths constants
# ---------------------------------------------------------------------------


def test_log_file_type():
    """LOG_FILE is a pathlib.Path instance."""
    import app.app_paths as ap

    assert isinstance(ap.LOG_FILE, pathlib.Path)


def test_config_file_in_same_root_as_log():
    """CONFIG_FILE and LOG_FILE are siblings in the same data-root directory."""
    import app.app_paths as ap

    assert ap.CONFIG_FILE.parent == ap.LOG_FILE.parent
