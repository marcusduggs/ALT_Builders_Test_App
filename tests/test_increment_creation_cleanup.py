"""
Regression test for three related fixes that came out of a real client
Windows path-length failure (see tests/test_increment_path_length.py
for the path-length fix itself):

  1. create_increment() cleans up after itself: if any of its three
     writes (increment.json, status.json, the version file copy)
     fails partway through, the half-created increment folder is
     removed rather than left behind as an orphan.
  2. list_increments() skips (and logs) any increment directory
     without a valid increment.json -- confirmed with a simulated
     orphan folder, so a leftover empty/half-written directory (from
     before this cleanup existed, or from any other cause) can never
     break the project's increment list.
  3. The error message fix: an OSError reaching the UI through the
     REAL upload flow (HomePage._upload_as_new_increment(), the same
     real run_with_progress() -> QThreadPool -> on_error path
     tests/test_unrecognized_template_error_dialog.py drives) shows
     "Could Not Save... The folder path may be too long" -- NOT the
     generic "This file couldn't be parsed" message a filesystem
     failure used to be shown as, which actively misled about the
     actual cause (the file parsed fine; saving it is what failed).
     The existing duplicate-name (ValueError) and unrecognized-template
     (also ValueError, see that other test) messages are confirmed
     UNCHANGED -- this only adds a new branch, it doesn't touch those.

Run directly: `python tests/test_increment_creation_cleanup.py`
"""

import logging
import os
import sys
import tempfile
from pathlib import Path
from unittest import mock

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from PySide6.QtCore import QThreadPool
from PySide6.QtWidgets import QApplication, QDialog, QMessageBox

from core import project_store as project_store_module
from ui.dialogs.display_name_dialog import DisplayNameDialog
from core.project_store import ProjectStore
from ui.mock_data import MockDataStore
from ui.pages.home_page import HomePage

GOOD_FIXTURE = os.path.join(os.path.dirname(__file__), "fixtures", "sample_increment.xlsm")
PROJECT_NAME = "Creation Cleanup Regression"


def main():
    failures = []

    with tempfile.TemporaryDirectory() as tmp:
        data_dir = Path(tmp) / "data"
        store = ProjectStore(base_dir=data_dir)
        store.create_project(PROJECT_NAME)
        project = store.get_project(PROJECT_NAME)

        # ------------------------------------------------------------
        # 1 -- create_increment() cleans up a half-created folder when
        # a write fails partway through (simulated by making
        # _copy_version_file's shutil.copy2 raise OSError -- the LAST
        # of the three writes, so increment.json and status.json have
        # already been written by the time it fails, exactly the
        # "half-created folder" scenario).
        # ------------------------------------------------------------
        print("=" * 70)
        print("1 -- create_increment() cleans up after a failed write")
        print("=" * 70)

        with mock.patch("shutil.copy2", side_effect=OSError(2, "No such file or directory")):
            try:
                store.create_increment(PROJECT_NAME, "INC 1 - Will Fail", GOOD_FIXTURE)
                failures.append("1: expected create_increment() to re-raise the simulated OSError, it didn't raise at all")
            except OSError:
                pass  # expected -- re-raised, not swallowed

        increments_dir = store._increments_dir(project.slug)
        leftover_dirs = list(increments_dir.iterdir()) if increments_dir.exists() else []
        print(f"  Increment directories left behind after the failed create: {leftover_dirs}")
        if leftover_dirs:
            failures.append(f"1: expected NO leftover increment directory after a failed create_increment(), found {leftover_dirs}")

        # Confirm a SUBSEQUENT, real (non-failing) create_increment()
        # for the same name now succeeds cleanly -- proves the failed
        # attempt didn't leave anything behind that would collide with
        # or otherwise interfere with a real retry.
        record = store.create_increment(PROJECT_NAME, "INC 1 - Will Fail", GOOD_FIXTURE)
        print(f"  Retried create_increment() after cleanup: slug={record.slug!r}, version={record.version}")
        if record.version != 1:
            failures.append(f"1: retried create_increment() should succeed as version 1, got version {record.version}")

        # ------------------------------------------------------------
        # 2 -- list_increments() skips (and logs) a simulated orphan
        # directory with no valid increment.json.
        # ------------------------------------------------------------
        print("\n" + "=" * 70)
        print("2 -- list_increments() skips and logs an orphan directory")
        print("=" * 70)

        orphan_dir = increments_dir / "orphan-with-no-increment-json"
        orphan_dir.mkdir(parents=True)
        (orphan_dir / "files").mkdir()  # a leftover partial write, but genuinely no increment.json

        corrupt_dir = increments_dir / "orphan-with-corrupt-increment-json"
        corrupt_dir.mkdir(parents=True)
        (corrupt_dir / "increment.json").write_text("{not valid json")

        logged_warnings = []
        with mock.patch.object(project_store_module.logger, "warning", side_effect=lambda *a, **kw: logged_warnings.append(a)):
            records = store.list_increments(PROJECT_NAME)

        names = [r.name for r in records]
        print(f"  list_increments() returned: {names}")
        print(f"  Warnings logged: {len(logged_warnings)}")
        if "INC 1 - Will Fail" not in names:
            failures.append("2: the real, valid increment should still be listed")
        if len(records) != 1:
            failures.append(f"2: expected exactly 1 valid increment (orphan/corrupt directories skipped), got {len(records)}: {names}")
        if len(logged_warnings) < 2:
            failures.append(f"2: expected a logged warning for EACH of the 2 bad directories (orphan + corrupt), got {len(logged_warnings)}")

        # A directory-listing call must never raise just because one of
        # its sibling directories is broken.
        try:
            store.get_increment(PROJECT_NAME, "INC 1 - Will Fail")
        except Exception as exc:  # noqa: BLE001 -- this call must not raise AT ALL here
            failures.append(f"2: get_increment() raised unexpectedly with orphan directories present: {exc!r}")

    # ------------------------------------------------------------
    # 3 -- the UI error message: an OSError reaching the REAL upload
    # flow shows "Could Not Save"/"folder path may be too long", not
    # the generic "couldn't be parsed" message.
    # ------------------------------------------------------------
    print("\n" + "=" * 70)
    print("3 -- UI error message for a storage failure (real upload flow)")
    print("=" * 70)

    with tempfile.TemporaryDirectory() as tmp:
        app = QApplication.instance() or QApplication(sys.argv)
        ui_store = MockDataStore(ProjectStore(base_dir=Path(tmp) / "data"))
        ui_store.add_project(PROJECT_NAME)
        page = HomePage(ui_store, on_view_data=lambda *a: None, on_view_combined=lambda *a: None)

        captured = []
        # _upload_as_new_increment() now shows a DisplayNameDialog
        # between the real parse and the actual save (Rey's
        # TIO_Changes.xlsx "Sum Data" tab change request #4) -- mocked
        # here to auto-accept the derived name, same as a user clicking
        # OK without changing anything, so this test can still reach
        # the real save step (where the simulated OSError below fires)
        # without a real modal dialog blocking on exec().
        with mock.patch.object(QMessageBox, "critical", staticmethod(lambda *a, **kw: captured.append(("critical", a)))), \
             mock.patch.object(QMessageBox, "warning", staticmethod(lambda *a, **kw: captured.append(("warning", a)))), \
             mock.patch.object(DisplayNameDialog, "exec", lambda self: QDialog.Accepted), \
             mock.patch.object(DisplayNameDialog, "value", lambda self: ""), \
             mock.patch("shutil.copy2", side_effect=OSError(2, "No such file or directory", "...increment.json.tmp")):

            page._upload_as_new_increment(PROJECT_NAME, GOOD_FIXTURE)
            QThreadPool.globalInstance().waitForDone()
            for _ in range(20):
                app.processEvents()

        print(f"Captured dialog calls: {len(captured)}")
        for kind, args in captured:
            print(f"  QMessageBox.{kind}: title={args[1]!r}")

        if not captured:
            failures.append("3: no dialog was shown at all for the simulated OSError -- would mean a crash, hang, or silent swallow")
        else:
            kind, args = captured[0]
            title = args[1] if len(args) > 1 else ""
            text = args[2] if len(args) > 2 else ""
            print(f"  Dialog title={title!r} text={text!r}")
            if "couldn't be parsed" in str(text) or "couldn't be parsed" in str(title):
                failures.append(f"3: the OSError dialog still shows the misleading 'couldn't be parsed' message: {text!r}")
            if "too long" not in str(text).lower():
                failures.append(f"3: expected the dialog to mention the folder path may be too long, got {text!r}")
            if kind != "critical":
                failures.append(f"3: expected QMessageBox.critical for a storage failure, got QMessageBox.{kind}")

    print("\n" + "=" * 70)
    if failures:
        print("RESULT: FAIL")
        for f in failures:
            print(" -", f)
        raise SystemExit(1)
    else:
        print(
            "RESULT: PASS -- create_increment() removes a half-created folder when a write fails partway "
            "through (and a retry afterward succeeds cleanly), list_increments() skips and logs any "
            "increment directory with a missing or corrupt increment.json without breaking the rest of the "
            "list, and a real OSError reaching the actual upload flow shows the clear 'could not save to "
            "the data folder, the folder path may be too long' message instead of the misleading 'couldn't "
            "be parsed' one."
        )


if __name__ == "__main__":
    main()
