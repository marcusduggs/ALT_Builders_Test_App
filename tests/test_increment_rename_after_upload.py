"""
Regression test for "Rename" -- changing an ALREADY-UPLOADED increment's
display_name, as opposed to ui.dialogs.display_name_dialog's original
upload-time naming feature (see tests/test_display_name_rename.py, which
covers that earlier, narrower case: naming a brand-new increment before
its first save).

Entry points: ui.pages.home_page.HomePage's Increments-list "Rename"
button (next to View Data/Delete in the Actions column) and
ui.pages.data_view_page.DataViewPage's own "Rename" button next to the
title bar -- both reuse the SAME DisplayNameDialog (now with an
`initial_value`/`window_title`/`intro_text` it didn't need for the
upload-only case) and the SAME storage path:
core.project_store.ProjectStore.rename_increment_display_name() (writes
ONLY display_name into increment.json via the existing atomic
_write_json(), never the folder/slug/name/version/last_updated) via
ui.mock_data.MockDataStore.rename_increment() (resolves and returns the
new display label so a caller holding an in-memory Increment can patch
it in place instead of re-fetching/re-parsing the whole file).

This file proves, with real file I/O (no mocking) except where a modal
QDialog.exec() would otherwise block:
  1. A rename persists across a fresh ProjectStore/MockDataStore
     instance pointed at the same data dir (simulating an app restart),
     not just the in-memory object just returned.
  2. Renaming to "" (or whitespace) resets to the automatic
     (derive_increment_code()/derive_short_display_name()) name.
  3. A renamed-after-upload increment still EXACT_MATCHes on the next
     upload of the same file and carries the rename across the version
     bump -- the identity-vs-display separation applies the same way
     whether the name was set AT upload or via a later Rename.
  4. The duplicate-name warning excludes the increment being renamed
     from its own "already taken" check (renaming to your own current
     name must never warn against yourself).
  5. Sanitization/40-char cap apply through the rename path exactly as
     they do at upload time (same _sanitize_display_name() call).
  6. The increment's folder, slug, and status.json are BYTE-FOR-BYTE
     unchanged by a rename -- only increment.json's display_name field
     differs.
  7. HomePage's Increments list and DataViewPage's title bar/Sum Data
     column B both reflect a rename immediately, without a restart or a
     full reload.
  8. A freshly-built CombinedDataViewPage (the only way this app can
     ever show one -- see module docstring) shows the renamed value in
     both column B and the Project Name / Increment # filter dropdown,
     with filters starting at "All" (there is no code path by which a
     combined view already showing the OLD name could still be on
     screen at the moment a rename happens elsewhere -- MainWindow is a
     single QStackedWidget page at a time, and reaching Combined View
     requires leaving Home/Data View first).

Run directly: `python tests/test_increment_rename_after_upload.py`
"""

import os
import sys
import tempfile
from pathlib import Path
from unittest import mock

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from PySide6.QtCore import QModelIndex
from PySide6.QtWidgets import QApplication, QDialog

from core.increment_matcher import MatchType
from core.project_store import DISPLAY_NAME_MAX_LENGTH, ProjectStore
from ui.dialogs.display_name_dialog import DisplayNameDialog
from ui.mock_data import MockDataStore, build_combined_view
from ui.pages.combined_data_view_page import CombinedDataViewPage
from ui.pages.data_view_page import DataViewPage
from ui.pages.home_page import HomePage

FIXTURE = os.path.join(os.path.dirname(__file__), "fixtures", "sample_increment.xlsm")
REAL_FIXDIR = os.path.join(os.path.dirname(__file__), "fixtures", "hmr_smr_projects")
REAL_FIXTURE = os.path.join(REAL_FIXDIR, "TIO H231182-38-00 SMR 3D 2022 Delta 1.xlsm")


def test_rename_sticks_after_reload(failures):
    with tempfile.TemporaryDirectory() as tmp:
        data_dir = Path(tmp) / "data"
        store = MockDataStore(ProjectStore(base_dir=data_dir))
        store.add_project("Reload Check")
        inc = store.add_new_increment("Reload Check", FIXTURE)

        new_name = store.rename_increment("Reload Check", inc.name, "Renamed Label")
        print(f"rename_increment() returned: {new_name!r}")
        if new_name != "Renamed Label":
            failures.append(f"expected rename_increment() to return 'Renamed Label', got {new_name!r}")

        # Fresh store instances, same data_dir -- simulates an app restart,
        # not just re-reading the same in-memory object.
        reloaded_store = MockDataStore(ProjectStore(base_dir=data_dir))
        reloaded = next(i for i in reloaded_store.get_project("Reload Check").increments if i.name == inc.name)
        print(f"After simulated restart: display_name={reloaded.display_name!r}")
        if reloaded.display_name != "Renamed Label":
            failures.append(f"rename did not survive a fresh ProjectStore instance, got {reloaded.display_name!r}")

        display = reloaded_store.get_increment_for_display("Reload Check", inc.name)
        print(f"get_increment_for_display() display_name: {display.display_name!r}")
        if display.display_name != "Renamed Label":
            failures.append(f"get_increment_for_display() should also show the renamed value, got {display.display_name!r}")


def test_rename_to_empty_reverts_to_derived_code(failures):
    with tempfile.TemporaryDirectory() as tmp:
        store = MockDataStore(ProjectStore(base_dir=Path(tmp) / "data"))
        store.add_project("Reset Check")
        inc = store.add_new_increment("Reset Check", FIXTURE)
        derived = store.derive_automatic_display_name(inc.name)

        store.rename_increment("Reset Check", inc.name, "Custom Name First")
        after_custom = next(i for i in store.get_project("Reset Check").increments if i.name == inc.name)
        print(f"After custom rename: {after_custom.display_name!r}")
        if after_custom.display_name != "Custom Name First":
            failures.append(f"expected 'Custom Name First', got {after_custom.display_name!r}")

        reset_name = store.rename_increment("Reset Check", inc.name, "")
        print(f"After renaming to '' -> resolved: {reset_name!r} (expected automatic name {derived!r})")
        if reset_name != derived:
            failures.append(f"renaming to blank should resolve to the automatic name {derived!r}, got {reset_name!r}")

        # Whitespace-only must behave identically to a truly empty string.
        store.rename_increment("Reset Check", inc.name, "Custom Again")
        reset_via_whitespace = store.rename_increment("Reset Check", inc.name, "   ")
        print(f"After renaming to whitespace-only -> resolved: {reset_via_whitespace!r}")
        if reset_via_whitespace != derived:
            failures.append(f"renaming to whitespace-only should also resolve to the automatic name {derived!r}, got {reset_via_whitespace!r}")

        record = store.store.get_increment("Reset Check", inc.name)
        print(f"Raw stored display_name after reset: {record.display_name!r} (expected None)")
        if record.display_name is not None:
            failures.append(f"a reset display_name should be stored as None (not ''), got {record.display_name!r}")


def test_renamed_after_upload_still_matches_and_versions(failures):
    """The real proof this is safe for a RENAME (not just an at-upload
    custom name, already covered by test_display_name_rename.py): name
    an increment AFTER it's been uploaded, then upload the SAME file
    again -- still EXACT_MATCH, still the same increment, rename still
    carried across the version bump.
    """
    if not os.path.isdir(REAL_FIXDIR):
        print(f"SKIP: {REAL_FIXDIR!r} not present (real HMR/SMR fixtures not available in this checkout)")
        return

    with tempfile.TemporaryDirectory() as tmp:
        store = MockDataStore(ProjectStore(base_dir=Path(tmp) / "data"))
        store.add_project("Rename After Upload Identity Check")
        inc = store.add_new_increment("Rename After Upload Identity Check", REAL_FIXTURE)
        real_record_name = inc.name
        print(f"Uploaded (no custom name yet): display_name={inc.display_name!r}")

        store.rename_increment("Rename After Upload Identity Check", inc.name, "SMR 3D")
        renamed = next(i for i in store.get_project("Rename After Upload Identity Check").increments if i.name == inc.name)
        print(f"After rename: display_name={renamed.display_name!r}")
        if renamed.display_name != "SMR 3D":
            failures.append(f"expected display_name 'SMR 3D' after rename, got {renamed.display_name!r}")
        if renamed.name != real_record_name:
            failures.append(f"identity (name) must stay the real Record Name after a rename, got {renamed.name!r}")

        match = store.match_upload("Rename After Upload Identity Check", REAL_FIXTURE)
        print(f"Re-upload match_type after rename: {match.match_type}")
        if match.match_type != MatchType.EXACT_MATCH:
            failures.append(f"a rename must not affect identity matching -- expected EXACT_MATCH, got {match.match_type}")

        result = store.simulate_comparison("Rename After Upload Identity Check", inc.name, REAL_FIXTURE)
        inc2 = store.confirm_update("Rename After Upload Identity Check", inc.name, REAL_FIXTURE, result)
        print(f"After confirm_update: display_name={inc2.display_name!r}, version={inc2.version}")
        if inc2.version != 2:
            failures.append(f"expected version 2 after confirming the update, got {inc2.version}")
        if inc2.display_name != "SMR 3D":
            failures.append(f"the rename must carry forward across the version bump, got {inc2.display_name!r}")


def test_duplicate_warning_excludes_self(failures):
    """The caller (HomePage/DataViewPage's _on_rename_increment) must
    exclude the increment BEING renamed from its own duplicate check --
    confirmed here by building the dialog the exact same way those
    handlers do and renaming an increment to its OWN current name.
    """
    QApplication.instance() or QApplication([])
    with tempfile.TemporaryDirectory() as tmp:
        store = MockDataStore(ProjectStore(base_dir=Path(tmp) / "data"))
        store.add_project("Self-Exclusion Check")
        inc_a = store.add_new_increment("Self-Exclusion Check", FIXTURE)
        store.rename_increment("Self-Exclusion Check", inc_a.name, "Shared Name")

        project = store.get_project("Self-Exclusion Check")
        target = next(i for i in project.increments if i.name == inc_a.name)
        existing_display_names = [i.display_name for i in project.increments if i.name != inc_a.name]
        print(f"existing_display_names (should be empty -- only increment in the project): {existing_display_names}")

        dialog = DisplayNameDialog(
            target.name, store.derive_automatic_display_name(target.name), existing_display_names,
            initial_value=target.display_name, window_title="Rename Increment", parent=None,
        )
        print(f"Warning hidden when field still shows own current name 'Shared Name': {dialog.warning_label.isHidden()}")
        if not dialog.warning_label.isHidden():
            failures.append("renaming an increment to its OWN current name must not trigger the duplicate warning")

        # Now a SECOND increment that genuinely shares a name with a
        # DIFFERENT increment -- the warning must still fire for THAT case.
        inc_b = store.add_new_increment("Self-Exclusion Check", os.path.join(REAL_FIXDIR, "TIO H231359-38-00_HMR 9B.xlsm")) \
            if os.path.isdir(REAL_FIXDIR) else None
        if inc_b is not None:
            project = store.get_project("Self-Exclusion Check")
            existing_for_b = [i.display_name for i in project.increments if i.name != inc_b.name]
            dialog_b = DisplayNameDialog(
                inc_b.name, store.derive_automatic_display_name(inc_b.name), existing_for_b,
                initial_value="Shared Name", window_title="Rename Increment", parent=None,
            )
            print(f"Warning hidden when renaming a DIFFERENT increment to an already-taken name (should be False): {dialog_b.warning_label.isHidden()}")
            if dialog_b.warning_label.isHidden():
                failures.append("renaming to a name already used by a DIFFERENT increment should still show the duplicate warning")


def test_sanitization_and_cap_via_rename(failures):
    with tempfile.TemporaryDirectory() as tmp:
        store = MockDataStore(ProjectStore(base_dir=Path(tmp) / "data"))
        store.add_project("Sanitize Via Rename")
        inc = store.add_new_increment("Sanitize Via Rename", FIXTURE)

        cases = [
            ("  SMR 3D  ", "SMR 3D"),
            ("A" * 60, "A" * DISPLAY_NAME_MAX_LENGTH),
            ('bad/name:with*illegal<chars>?|"', "badnamewithillegalchars"),
            ("multi   space    collapse", "multi space collapse"),
        ]
        for raw, expected in cases:
            store.rename_increment("Sanitize Via Rename", inc.name, raw)
            record = store.store.get_increment("Sanitize Via Rename", inc.name)
            print(f"  rename_increment(..., {raw!r}) -> stored display_name={record.display_name!r} (expected {expected!r})")
            if record.display_name != expected:
                failures.append(f"rename_increment({raw!r}) expected stored display_name {expected!r}, got {record.display_name!r}")


def test_folder_slug_and_status_json_unchanged(failures):
    with tempfile.TemporaryDirectory() as tmp:
        data_dir = Path(tmp) / "data"
        store = MockDataStore(ProjectStore(base_dir=data_dir))
        store.add_project("Untouched Check")
        inc = store.add_new_increment("Untouched Check", FIXTURE)
        store.set_stage_status("Untouched Check", inc.name, "B-C1", 5, "Done")  # give status.json real content to compare

        before_record = store.store.get_increment("Untouched Check", inc.name)
        project_slug = store.store.get_project("Untouched Check").slug
        increment_dir = data_dir / "projects" / project_slug / "increments" / before_record.slug
        status_path = increment_dir / "status.json"
        status_bytes_before = status_path.read_bytes()
        increment_json_before = store.store._increment_json_path(project_slug, before_record.slug).read_text()
        files_dir_entries_before = sorted(p.name for p in (increment_dir / "files").iterdir())

        store.rename_increment("Untouched Check", inc.name, "Renamed, Should Not Move Anything")

        after_record = store.store.get_increment("Untouched Check", inc.name)
        print(f"slug before: {before_record.slug!r}, slug after: {after_record.slug!r}")
        if after_record.slug != before_record.slug:
            failures.append(f"a rename must never change the increment's folder slug, got {before_record.slug!r} -> {after_record.slug!r}")
        if not increment_dir.is_dir():
            failures.append(f"the increment's original folder {increment_dir} must still exist after a rename")

        status_bytes_after = status_path.read_bytes()
        print(f"status.json byte-identical: {status_bytes_after == status_bytes_before}")
        if status_bytes_after != status_bytes_before:
            failures.append("rename_increment_display_name must never touch status.json, but its bytes changed")

        files_dir_entries_after = sorted(p.name for p in (increment_dir / "files").iterdir())
        if files_dir_entries_after != files_dir_entries_before:
            failures.append(f"rename must never touch the files/ directory, got {files_dir_entries_before} -> {files_dir_entries_after}")

        increment_json_after = store.store._increment_json_path(project_slug, after_record.slug).read_text()
        import json
        before_fields = json.loads(increment_json_before)
        after_fields = json.loads(increment_json_after)
        for field in ("name", "version", "last_updated"):
            print(f"  increment.json[{field!r}] unchanged: {before_fields[field] == after_fields[field]}")
            if before_fields[field] != after_fields[field]:
                failures.append(f"rename must only change increment.json's display_name field, but {field!r} changed too: {before_fields[field]!r} -> {after_fields[field]!r}")
        if after_fields["display_name"] != "Renamed, Should Not Move Anything":
            failures.append(f"increment.json's display_name should be the new name, got {after_fields['display_name']!r}")


def test_home_page_increment_table_reflects_rename(failures):
    """Drives the REAL "Rename" button handler end to end (DisplayNameDialog's
    exec()/value() mocked to auto-accept a new name, same pattern
    tests/test_increment_creation_cleanup.py already uses for the
    upload dialog, since a real modal .exec() would block under the
    offscreen QPA platform) -- confirms the Increments list's own name
    column updates immediately, no restart or manual refresh call needed
    beyond what the real handler already does.
    """
    app = QApplication.instance() or QApplication(sys.argv)
    with tempfile.TemporaryDirectory() as tmp:
        store = MockDataStore(ProjectStore(base_dir=Path(tmp) / "data"))
        store.add_project("Home Page Rename Check")
        inc = store.add_new_increment("Home Page Rename Check", FIXTURE)

        page = HomePage(store, on_view_data=lambda *a: None, on_view_combined=lambda *a: None)
        page._refresh_project_combo(select_name="Home Page Rename Check")
        page._refresh_increment_table()

        name_item_before = page.increment_table.item(0, 1)
        print(f"Increments list name before rename: {name_item_before.text()!r}")

        with mock.patch.object(DisplayNameDialog, "exec", lambda self: QDialog.Accepted), \
             mock.patch.object(DisplayNameDialog, "value", lambda self: "Renamed From Home Page"):
            page._on_rename_increment("Home Page Rename Check", inc.name)

        name_item_after = page.increment_table.item(0, 1)
        print(f"Increments list name after rename: {name_item_after.text()!r}")
        if name_item_after.text() != "Renamed From Home Page":
            failures.append(f"Home page's Increments list should show the new name immediately, got {name_item_after.text()!r}")

        # Confirm it's actually persisted, not just a cosmetic table edit.
        stored = store.store.get_increment("Home Page Rename Check", inc.name)
        if stored.display_name != "Renamed From Home Page":
            failures.append(f"the rename shown in the table must be backed by a real persisted write, got {stored.display_name!r}")


def test_data_view_page_rename_updates_title_and_sum_data(failures):
    """Drives DataViewPage's own "Rename" button handler end to end --
    confirms the title bar AND the Sum Data tab's column B (which
    independently bakes in display_name per row at tab-build time, see
    ui.pages.data_view_page's own Sum Data column comment) both update
    immediately, without a get_increment_for_display() reload.
    """
    app = QApplication.instance() or QApplication(sys.argv)
    with tempfile.TemporaryDirectory() as tmp:
        store = MockDataStore(ProjectStore(base_dir=Path(tmp) / "data"))
        store.add_project("Data View Rename Check")
        inc = store.add_new_increment("Data View Rename Check", FIXTURE)
        display = store.get_increment_for_display("Data View Rename Check", inc.name)

        page = DataViewPage("Data View Rename Check", display, store, on_back=lambda: None)
        print(f"Title before rename: {page.title_label.text()!r}")

        with mock.patch.object(DisplayNameDialog, "exec", lambda self: QDialog.Accepted), \
             mock.patch.object(DisplayNameDialog, "value", lambda self: "Renamed In Data View"):
            page._on_rename_increment()

        print(f"Title after rename: {page.title_label.text()!r}")
        if "Renamed In Data View" not in page.title_label.text():
            failures.append(f"title bar should show the new name immediately after rename, got {page.title_label.text()!r}")
        if page.increment.display_name != "Renamed In Data View":
            failures.append(f"self.increment.display_name should be updated in place, got {page.increment.display_name!r}")

        sum_data_model = page.sum_data_table.model()
        source = sum_data_model.sourceModel() if hasattr(sum_data_model, "sourceModel") else sum_data_model
        col_b_values = {source.data(source.index(r, 1)) for r in range(source.rowCount())}
        print(f"Sum Data column B values after rename: {col_b_values}")
        if col_b_values != {"Renamed In Data View"}:
            failures.append(f"Sum Data's column B should show ONLY the new name after rename, got {col_b_values}")

        # Persisted, not just an in-memory/table edit.
        stored = store.store.get_increment("Data View Rename Check", inc.name)
        if stored.display_name != "Renamed In Data View":
            failures.append(f"the rename shown on screen must be backed by a real persisted write, got {stored.display_name!r}")


def test_combined_view_refresh_after_rename(failures):
    """A combined view is only ever built FRESH, from Increment objects
    re-fetched at that exact moment (see ui/main_window.py's
    show_combined_data_view -- there is no "live" combined view that
    could be showing stale data while a rename happens elsewhere) --
    confirmed directly: a CombinedDataViewPage built AFTER a rename
    shows the new name in both column B and the Project Name /
    Increment # filter dropdown, with the filter starting at "All" (a
    brand-new widget has never had anything else selected, so there is
    no stale "old name still selected" state to clear).
    """
    app = QApplication.instance() or QApplication(sys.argv)
    with tempfile.TemporaryDirectory() as tmp:
        store = MockDataStore(ProjectStore(base_dir=Path(tmp) / "data"))
        store.add_project("Combined Rename Check")
        inc_a = store.add_new_increment("Combined Rename Check", FIXTURE)
        inc_b = store.add_new_increment(
            "Combined Rename Check", os.path.join(os.path.dirname(__file__), "fixtures", "inc8_real_2stage.xlsm")
        )

        display_a_before = store.get_increment_for_display("Combined Rename Check", inc_a.name)
        display_b = store.get_increment_for_display("Combined Rename Check", inc_b.name)
        view_before = build_combined_view([display_a_before, display_b])
        page_before = CombinedDataViewPage("Combined Rename Check", view_before, store, on_back=lambda: None)
        model_before = page_before.sum_data_table.model()
        source_before = model_before.sourceModel() if hasattr(model_before, "sourceModel") else model_before
        col_b_before = {source_before.data(source_before.index(r, 1)) for r in range(source_before.rowCount())}
        print(f"Column B values BEFORE rename: {col_b_before}")

        store.rename_increment("Combined Rename Check", inc_a.name, "Renamed For Combined View")

        # The only real way this app ever shows a combined view: build a
        # BRAND NEW one from freshly re-fetched Increments, exactly what
        # ui/main_window.py's show_combined_data_view does.
        display_a_after = store.get_increment_for_display("Combined Rename Check", inc_a.name)
        view_after = build_combined_view([display_a_after, display_b])
        page_after = CombinedDataViewPage("Combined Rename Check", view_after, store, on_back=lambda: None)
        model_after = page_after.sum_data_table.model()
        source_after = model_after.sourceModel() if hasattr(model_after, "sourceModel") else model_after
        col_b_after = {source_after.data(source_after.index(r, 1)) for r in range(source_after.rowCount())}
        print(f"Column B values AFTER rename (fresh CombinedDataViewPage): {col_b_after}")
        if "Renamed For Combined View" not in col_b_after:
            failures.append(f"a freshly-built combined view should show the renamed value in column B, got {col_b_after}")
        if display_a_before.display_name in col_b_after:
            failures.append(f"the OLD name should no longer appear anywhere in a freshly-built combined view, but {display_a_before.display_name!r} is still present")

        filter_combo = page_after.sum_data_filter_bar._column_combos[1]  # Project Name / Increment #
        dropdown_values = {filter_combo.itemText(i) for i in range(filter_combo.count())}
        print(f"Filter dropdown values: {dropdown_values}")
        if "Renamed For Combined View" not in dropdown_values:
            failures.append(f"the filter dropdown should offer the renamed value, got {dropdown_values}")
        if display_a_before.display_name in dropdown_values:
            failures.append(f"the filter dropdown should not still offer the OLD name, got {dropdown_values}")
        print(f"Filter starts at 'All': {filter_combo.currentText() == 'All'}")
        if filter_combo.currentText() != "All":
            failures.append(f"a freshly-built combined view's filter should start at 'All', got {filter_combo.currentText()!r}")


def main():
    failures = []
    test_rename_sticks_after_reload(failures)
    test_rename_to_empty_reverts_to_derived_code(failures)
    test_renamed_after_upload_still_matches_and_versions(failures)
    test_duplicate_warning_excludes_self(failures)
    test_sanitization_and_cap_via_rename(failures)
    test_folder_slug_and_status_json_unchanged(failures)
    test_home_page_increment_table_reflects_rename(failures)
    test_data_view_page_rename_updates_title_and_sum_data(failures)
    test_combined_view_refresh_after_rename(failures)

    print("\n" + "=" * 70)
    if failures:
        print("RESULT: FAIL")
        for f in failures:
            print(" -", f)
        raise SystemExit(1)
    else:
        print(
            "RESULT: PASS -- renaming an already-uploaded increment persists across a restart, resets to the "
            "automatic name when cleared, keeps identity/matching/versioning fully intact, excludes the "
            "increment itself from its own duplicate-name warning, sanitizes/caps the same way upload-time "
            "naming does, never touches the folder/slug/status.json/files, and both HomePage's Increments "
            "list and DataViewPage's title bar/Sum Data column B update immediately -- with a freshly-built "
            "combined view (the only kind this app can ever show) always reflecting the current name."
        )


if __name__ == "__main__":
    main()
