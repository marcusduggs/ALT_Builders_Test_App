"""
Regression test confirming core.excel_reader.detect_stage_count()'s
ValueError (raised for an unrecognized grid-sheet template -- see
tests/test_legacy_stage_header.py) actually reaches the user as a
QMessageBox dialog through the REAL upload flow, not just a direct
unit-test call to detect_stage_count() itself.

Drives ui.pages.home_page.HomePage._upload_as_new_increment() -- the
exact method HomePage._on_upload_file() calls for a NO_MATCH file --
end to end: a real run_with_progress() call, dispatched onto a real
QThreadPool background worker (ui/workers.py's _Worker.run(), which
catches ANY exception and emits it on the `error` signal), delivered
back to the UI thread's on_error closure, which shows a QMessageBox.
QT_QPA_PLATFORM=offscreen (set below, before any Qt import) lets this
run headless; QMessageBox.critical/warning/information are patched to
CAPTURE the call instead of blocking on a real modal dialog (there's no
user present to click it in a test).

Confirms (and documents) a real wiring detail found during
investigation: _upload_as_new_increment()'s on_error has a
`isinstance(exc, ValueError)` branch written for a DIFFERENT cause (an
increment with this Record Name already existing -- raised by
create_increment(), see its docstring) that also happens to catch this
ValueError, since detect_stage_count()'s error is raised earlier in the
same call (inside normalize_workbook(), before create_increment() ever
runs) and is also a plain ValueError. The dialog still correctly shows
up with the real, specific error message -- this test locks in that it
reaches the user at all, not that the title/icon are necessarily the
ones you'd write from scratch for this specific cause.

Run directly: `python tests/test_unrecognized_template_error_dialog.py`
"""

import os
import sys
import tempfile
from pathlib import Path
from unittest import mock

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import openpyxl
from PySide6.QtCore import QThreadPool
from PySide6.QtWidgets import QApplication, QMessageBox

from core import excel_reader
from core.project_store import ProjectStore
from ui.mock_data import MockDataStore
from ui.pages.home_page import HomePage

GOOD_FIXTURE = os.path.join(os.path.dirname(__file__), "fixtures", "sample_increment.xlsm")
PROJECT_NAME = "Error Dialog Real-Flow Regression"


def _build_unrecognized_template_file(dest_path: str) -> None:
    """A real .xlsm file, otherwise identical to a known-good fixture,
    with its B-Tests header corrupted to match NEITHER the new "Stage N
    Required" nor the legacy "REQUIRED (Select with..." pattern -- a
    genuine third, unrecognized format, not a synthetic in-memory
    stand-in for one.
    """
    wb = openpyxl.load_workbook(GOOD_FIXTURE, keep_vba=True)
    ws = wb["B-Tests"]
    header_row = excel_reader._find_header_row(ws, excel_reader.GRID_INDEX_COL, excel_reader.GRID_HEADER_LABEL)
    ws.cell(row=header_row, column=excel_reader.GRID_STAGE_FIRST_COL, value="Totally Unrecognized Column Header")
    wb.save(dest_path)


def main():
    failures = []

    with tempfile.TemporaryDirectory() as tmp:
        bad_file = str(Path(tmp) / "bad_template.xlsm")
        _build_unrecognized_template_file(bad_file)

        # Fixture assumption guard: confirm the corrupted file really
        # does raise ValueError when parsed directly, before testing
        # whether that propagates through the UI.
        try:
            excel_reader.normalize_workbook(bad_file)
            failures.append("setup: expected the corrupted file to raise ValueError, it parsed cleanly instead")
        except ValueError:
            pass

        app = QApplication.instance() or QApplication(sys.argv)
        store = MockDataStore(ProjectStore(base_dir=Path(tmp) / "data"))
        store.add_project(PROJECT_NAME)
        page = HomePage(store, on_view_data=lambda *a: None, on_view_combined=lambda *a: None)

        captured = []
        with mock.patch.object(QMessageBox, "critical", staticmethod(lambda *a, **kw: captured.append(("critical", a)))), \
             mock.patch.object(QMessageBox, "warning", staticmethod(lambda *a, **kw: captured.append(("warning", a)))), \
             mock.patch.object(QMessageBox, "information", staticmethod(lambda *a, **kw: captured.append(("information", a)))):

            # The exact method _on_upload_file() calls for a NO_MATCH
            # file -- real run_with_progress() call, real background
            # worker, real on_error wiring.
            page._upload_as_new_increment(PROJECT_NAME, bad_file)

            QThreadPool.globalInstance().waitForDone()
            for _ in range(20):
                app.processEvents()

        print(f"Captured dialog calls: {len(captured)}")
        for kind, args in captured:
            print(f"  QMessageBox.{kind}: title={args[1]!r}")

        if not captured:
            failures.append(
                "No dialog was shown at all -- the ValueError did not reach the UI error path (this would "
                "mean a crash, hang, or silent swallow instead)"
            )
        else:
            kind, args = captured[0]
            text = args[2] if len(args) > 2 else ""
            print(f"  Dialog text: {text!r}")
            if "unrecognized template format" not in str(text) and "Totally Unrecognized Column Header" not in str(text):
                failures.append(
                    f"A dialog was shown, but its text doesn't contain the actual ValueError message: {text!r}"
                )

    print("\n" + "=" * 70)
    if failures:
        print("RESULT: FAIL")
        for f in failures:
            print(" -", f)
        raise SystemExit(1)
    else:
        print("RESULT: PASS -- an unrecognized-template ValueError, raised while uploading a real file "
              "through the ACTUAL HomePage._upload_as_new_increment() -> run_with_progress() -> QThreadPool "
              "background worker -> on_error path, reaches the user as a QMessageBox dialog containing the "
              "real error message -- not a crash, hang, or silent failure.")


if __name__ == "__main__":
    main()
