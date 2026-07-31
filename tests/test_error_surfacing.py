"""
Validates the fix for a real reported bug: clicking "View Data" on an
increment that fails to load showed the loading spinner, then silently
did nothing -- no data view, no error, no crash. Root cause was a bare
`if increment is None: return` in ui.main_window.MainWindow.show_data_view's
on_finished callback (core.ui.mock_data.MockDataStore.get_increment_for_display
returns None rather than raising when the increment record or a specific
version's file can't be found -- see that function's docstring). Fixed
alongside two related robustness gaps found during investigation:

  1. No logging was configured anywhere in the app at all -- see
     core/app_logging.py's module docstring for why that matters (the
     release build is --windowed, so a bare print()/logger call before
     had nowhere visible to go). ui.workers._Worker.run() now logs
     every background-task exception (with full traceback) BEFORE
     forwarding it to the caller's on_error -- centralizing this once,
     rather than requiring every run_with_progress caller to remember
     to log it themselves.
  2. ui.mock_data.MockDataStore.add_new_increment() named a newly
     created increment from core.excel_reader.parse_project_info()'s
     raw fixed-cell "record_name" -- a different, more tolerant read
     meant only for cosmetic display -- instead of get_record_name()
     (search-based, whitespace-normalized, the actual documented
     identity source of truth core.increment_matcher already used to
     decide this WAS a new increment in the first place). Not confirmed
     as the specific cause of the reported bug, but a real inconsistency
     worth fixing regardless.

This test deliberately stays at the core/ui.mock_data/ui.workers layer,
not a full QApplication+MainWindow GUI test -- no other test in this
suite spins up a Qt event loop (see ui.workers._Worker.run(), called
directly here: it's a plain method underneath its @Slot() decorator,
so it can be exercised synchronously with no QThreadPool/event loop at
all), and the underlying conditions (get_increment_for_display()
returning None, a malformed file raising a clear exception, an
exception getting logged) are exactly what ui/main_window.py's fixed
on_finished/on_error handlers now branch on.

Run directly: `python tests/test_error_surfacing.py`
"""

import os
import shutil
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import openpyxl

from core import app_logging
from core.project_store import ProjectStore
from ui.mock_data import MockDataStore
from ui.workers import _Worker

FIXTURE = os.path.join(os.path.dirname(__file__), "fixtures", "sample_increment.xlsm")
PROJECT_NAME = "Error Surfacing Validation"


def _make_malformed_file(tmp: Path) -> Path:
    """A copy of the real fixture with its A-Project Info sheet deleted
    -- "a genuinely different project structure" (missing a required
    sheet), not just a truncated/corrupted file.
    """
    malformed_path = tmp / "malformed_missing_project_info.xlsm"
    shutil.copy(FIXTURE, malformed_path)
    wb = openpyxl.load_workbook(malformed_path, keep_vba=True)
    del wb["A-Project Info"]
    wb.save(malformed_path)
    return malformed_path


def main():
    failures = []

    with tempfile.TemporaryDirectory() as tmp_str:
        tmp = Path(tmp_str)
        log_path = tmp / "app.log"
        app_logging.configure_logging(log_path=log_path)

        store = MockDataStore(ProjectStore(base_dir=tmp / "data"))
        store.add_project(PROJECT_NAME)

        # ------------------------------------------------------------
        # 1 -- get_increment_for_display() returns None (doesn't raise)
        # for a genuinely nonexistent increment -- the exact precondition
        # ui.main_window.show_data_view's on_finished must now handle
        # instead of silently no-op'ing.
        # ------------------------------------------------------------
        print("=" * 70)
        print("1 -- get_increment_for_display returns None, doesn't raise, for a missing increment")
        print("=" * 70)

        result = store.get_increment_for_display(PROJECT_NAME, "INC 999 - Does Not Exist")
        print(f"Result: {result!r}")
        if result is not None:
            failures.append("1: expected None for a nonexistent increment, got a real Increment")

        # Same for a missing VERSION of a real increment.
        real_inc = store.add_new_increment(PROJECT_NAME, FIXTURE)
        result_bad_version = store.get_increment_for_display(PROJECT_NAME, real_inc.name, version=99)
        print(f"Result for version=99 (only v1 exists): {result_bad_version!r}")
        if result_bad_version is not None:
            failures.append("1: expected None for a nonexistent version, got a real Increment")

        # ------------------------------------------------------------
        # 2 -- a malformed file (missing a required sheet) raises a
        # real, specific exception from add_new_increment() -- never
        # silently returns something bogus, never hangs.
        # ------------------------------------------------------------
        print("\n" + "=" * 70)
        print("2 -- malformed file (missing A-Project Info) raises clearly")
        print("=" * 70)

        malformed_path = _make_malformed_file(tmp)
        raised = None
        try:
            store.add_new_increment(PROJECT_NAME, str(malformed_path))
        except Exception as exc:  # noqa: BLE001 -- deliberately broad, this IS the assertion
            raised = exc
        print(f"Raised: {type(raised).__name__ if raised else None}: {raised}")
        if raised is None:
            failures.append("2: expected add_new_increment() to raise on a file missing A-Project Info, it didn't")

        # ------------------------------------------------------------
        # 3 -- ui.workers._Worker.run() logs the full exception (with
        # traceback) to the file BEFORE forwarding it to on_error --
        # exercised directly and synchronously (see module docstring
        # for why this doesn't need a QThreadPool/event loop).
        # ------------------------------------------------------------
        print("\n" + "=" * 70)
        print("3 -- background task exceptions are logged with a full traceback")
        print("=" * 70)

        def _raises():
            raise ValueError("deliberate failure for test 3")

        captured = []
        worker = _Worker(_raises, (), {}, description="test 3 task")
        worker.signals.error.connect(lambda exc: captured.append(exc))
        worker.run()

        print(f"Signal captured exception: {captured}")
        if len(captured) != 1 or not isinstance(captured[0], ValueError):
            failures.append(f"3: expected exactly one ValueError forwarded via signals.error, got {captured}")

        log_text = log_path.read_text()
        print(f"Log file size: {len(log_text)} bytes")
        if "Background task failed" not in log_text:
            failures.append("3: log file should record that a background task failed")
        if "deliberate failure for test 3" not in log_text:
            failures.append("3: log file should contain the actual exception message")
        if "Traceback" not in log_text:
            failures.append("3: log file should contain a full traceback, not just the exception message")

        # ------------------------------------------------------------
        # 4 -- add_new_increment() names the increment via
        # get_record_name() (the documented identity source of truth),
        # not parse_project_info()'s cosmetic fixed-cell read -- both
        # happen to agree for this real fixture, so this just confirms
        # the fix didn't change the correct, common-case name.
        # ------------------------------------------------------------
        print("\n" + "=" * 70)
        print("4 -- increment naming still correct for a real file")
        print("=" * 70)

        print(f"Increment name: {real_inc.name!r}")
        if real_inc.name != "INC 2 - Foundation and Underground Utilities":
            failures.append(f"4: unexpected increment name: {real_inc.name!r}")

    print("\n" + "=" * 70)
    if failures:
        print("RESULT: FAIL")
        for f in failures:
            print(" -", f)
        raise SystemExit(1)
    else:
        print("RESULT: PASS -- get_increment_for_display's None-return precondition confirmed, a malformed "
              "file raises clearly instead of hanging, background-task exceptions are logged with a full "
              "traceback, and increment naming is unaffected for real files")


if __name__ == "__main__":
    main()
