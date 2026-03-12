"""
test_thread_safety.py — tests for thread-safety primitives in the generation
pipeline.

These tests do NOT start a Tk window or a real GUI.  They directly instantiate
the threading primitives and helper logic extracted from the GUI class in order
to verify:

- _gen_proc_lock protects concurrent access to _gen_proc
- _stop_gen (threading.Event) can be set/cleared/checked correctly
- cancel_generation pattern (lock-snapshot + kill) tolerates None proc
- cancel_generation pattern tolerates a proc that has already exited
- Rapid set/clear of _stop_gen from multiple threads is race-free
"""

import pathlib
import sys
import threading
import time
from unittest.mock import MagicMock


ROOT = pathlib.Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))


# ---------------------------------------------------------------------------
# Minimal stand-alone replica of the relevant GUI state (no Tk dependency)
# ---------------------------------------------------------------------------


class _FakeGen:
    """Replicates the threading primitives used by ResilienceScanGUI."""

    def __init__(self):
        self._gen_proc = None
        self._gen_proc_lock = threading.Lock()
        self._stop_gen = threading.Event()

    def _set_proc(self, proc):
        with self._gen_proc_lock:
            self._gen_proc = proc

    def _clear_proc(self):
        with self._gen_proc_lock:
            self._gen_proc = None

    def cancel(self):
        """Mirror of cancel_generation (without the messagebox confirmation)."""
        self._stop_gen.set()
        with self._gen_proc_lock:
            proc = self._gen_proc
        if proc is not None:
            try:
                proc.kill()
                proc.wait()
            except (OSError, AttributeError):
                pass


# ---------------------------------------------------------------------------
# _stop_gen (threading.Event) behaviour
# ---------------------------------------------------------------------------


def test_stop_gen_initially_clear():
    gen = _FakeGen()
    assert not gen._stop_gen.is_set()


def test_stop_gen_set_is_detectable():
    gen = _FakeGen()
    gen._stop_gen.set()
    assert gen._stop_gen.is_set()


def test_stop_gen_clear_resets():
    gen = _FakeGen()
    gen._stop_gen.set()
    gen._stop_gen.clear()
    assert not gen._stop_gen.is_set()


def test_stop_gen_thread_sees_set():
    """A background thread polling _stop_gen detects the set() from main thread."""
    gen = _FakeGen()
    detected = threading.Event()

    def _worker():
        while not gen._stop_gen.is_set():
            time.sleep(0.001)
        detected.set()

    t = threading.Thread(target=_worker, daemon=True)
    t.start()
    time.sleep(0.01)
    gen._stop_gen.set()
    assert detected.wait(timeout=1.0), "Worker never detected _stop_gen.set()"
    t.join(timeout=1.0)


# ---------------------------------------------------------------------------
# _gen_proc_lock protects _gen_proc
# ---------------------------------------------------------------------------


def test_gen_proc_initially_none():
    gen = _FakeGen()
    assert gen._gen_proc is None


def test_gen_proc_set_and_clear():
    gen = _FakeGen()
    fake_proc = MagicMock()
    gen._set_proc(fake_proc)
    assert gen._gen_proc is fake_proc
    gen._clear_proc()
    assert gen._gen_proc is None


def test_gen_proc_lock_concurrent_writes():
    """Many threads writing _gen_proc concurrently should never corrupt state."""
    gen = _FakeGen()
    errors = []

    def _writer(val):
        try:
            for _ in range(50):
                gen._set_proc(val)
                gen._clear_proc()
        except Exception as e:
            errors.append(e)

    threads = [threading.Thread(target=_writer, args=(i,)) for i in range(8)]
    for t in threads:
        t.start()
    for t in threads:
        t.join()

    assert not errors, f"Errors during concurrent writes: {errors}"


# ---------------------------------------------------------------------------
# cancel() with various _gen_proc states
# ---------------------------------------------------------------------------


def test_cancel_with_none_proc_does_not_raise():
    """cancel() must not raise when _gen_proc is None (not yet started)."""
    gen = _FakeGen()
    gen.cancel()  # should be a no-op


def test_cancel_calls_kill_on_running_proc():
    """cancel() calls proc.kill() and proc.wait() when proc is set."""
    gen = _FakeGen()
    fake_proc = MagicMock()
    gen._set_proc(fake_proc)

    gen.cancel()

    fake_proc.kill.assert_called_once()
    fake_proc.wait.assert_called_once()


def test_cancel_tolerates_oserror_on_kill():
    """cancel() suppresses OSError from proc.kill() (process already dead)."""
    gen = _FakeGen()
    fake_proc = MagicMock()
    fake_proc.kill.side_effect = OSError("already dead")
    gen._set_proc(fake_proc)

    gen.cancel()  # must not raise


def test_cancel_tolerates_attribute_error():
    """cancel() suppresses AttributeError (proc became None between snapshot and kill)."""
    gen = _FakeGen()

    class _RacyProc:
        def kill(self):
            raise AttributeError("proc gone")

        def wait(self):
            pass

    gen._set_proc(_RacyProc())
    gen.cancel()  # must not raise


def test_cancel_sets_stop_gen():
    """cancel() always sets _stop_gen regardless of proc state."""
    gen = _FakeGen()
    gen.cancel()
    assert gen._stop_gen.is_set()


# ---------------------------------------------------------------------------
# Rapid concurrent set/clear race test
# ---------------------------------------------------------------------------


def test_stop_gen_no_race_under_concurrent_access():
    """Concurrent set/clear of _stop_gen from many threads must not deadlock."""
    gen = _FakeGen()
    stop = threading.Event()

    def _toggler():
        while not stop.is_set():
            gen._stop_gen.set()
            gen._stop_gen.clear()

    threads = [threading.Thread(target=_toggler, daemon=True) for _ in range(4)]
    for t in threads:
        t.start()
    time.sleep(0.05)
    stop.set()
    for t in threads:
        t.join(timeout=1.0)
    # If we reach here without deadlock/hang the test passes
