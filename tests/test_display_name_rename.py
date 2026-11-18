"""
Regression test for Rey's TIO_Changes.xlsx "Sum Data" tab change
request #4: "Add the Option to change the name of the project before
the Upload". core.project_store.ProjectStore.create_increment() now
takes an optional display_name, persisted in increment.json as a
separate field (core.project_store.IncrementRecord.display_name) and
NEVER fed into _slugify_increment() -- it is a purely cosmetic label,
distinct from the real identity (get_record_name()'s Record Name
string, which core.increment_matcher and status.json keys always use
unchanged).

This file proves, with real file I/O (no mocking):
  1. Setting a custom display_name at upload has NO effect on the
     on-disk folder slug (confirmed identical to the no-display_name
     case) -- the Windows MAX_PATH budget logic in
     core.project_store._slugify_increment()/_increment_slug_max_length()
     is therefore unaffected.
  2. A renamed increment still EXACT_MATCHes (core.increment_matcher)
     when the SAME file is uploaded again -- identity is untouched.
  3. Re-uploading as a new version carries display_name forward
     unchanged (save_new_version() must not drop it), and bumps the
     version correctly.
  4. sanitization (trim, 40-char cap, filename-illegal characters
     stripped) and the blank-falls-back-to-derived rule.
  5. An OLDER increment (created before display_name existed, or with
     one never set) still works -- falls back to the derived short
     name, never crashes, never shows None/blank.

Run directly: `python tests/test_display_name_rename.py`
"""

import os
import sys
import tempfile
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from PySide6.QtWidgets import QApplication

from core import excel_reader
from core.increment_matcher import MatchType
from core.project_store import DISPLAY_NAME_MAX_LENGTH, ProjectStore, _sanitize_display_name
from ui.mock_data import MockDataStore
from ui.pages.data_view_page import DataViewPage

FIXTURE = os.path.join(os.path.dirname(__file__), "fixtures", "sample_increment.xlsm")
REAL_FIXDIR = os.path.join(os.path.dirname(__file__), "fixtures", "hmr_smr_projects")
REAL_FIXTURE = os.path.join(REAL_FIXDIR, "TIO H231182-38-00 SMR 3D 2022 Delta 1.xlsm")


def test_display_name_never_affects_slug(failures):
    with tempfile.TemporaryDirectory() as tmp:
        store = MockDataStore(ProjectStore(base_dir=Path(tmp) / "data"))

        store.add_project("With Custom Name")
        prepared_a = store.prepare_new_increment("With Custom Name", FIXTURE)
        inc_a = store.finalize_new_increment(prepared_a, "A" * DISPLAY_NAME_MAX_LENGTH)
        slug_a = store.store.get_increment("With Custom Name", inc_a.name).slug

        store.add_project("No Custom Name")
        prepared_b = store.prepare_new_increment("No Custom Name", FIXTURE)
        inc_b = store.finalize_new_increment(prepared_b, None)
        slug_b = store.store.get_increment("No Custom Name", inc_b.name).slug

        print(f"slug with max-length custom display_name: {slug_a!r}")
        print(f"slug with no display_name:                 {slug_b!r}")
        if slug_a != slug_b:
            failures.append(f"display_name must have ZERO effect on the folder slug, got {slug_a!r} vs {slug_b!r}")


def test_renamed_increment_still_matches_and_versions(failures):
    """The real proof this fix is safe: rename an increment, then
    upload the SAME real file again -- it must still be recognized as
    an update to the SAME increment (not a new one), and the rename
    must survive the version bump.
    """
    if not os.path.isdir(REAL_FIXDIR):
        print(f"SKIP: {REAL_FIXDIR!r} not present (real HMR/SMR fixtures not available in this checkout)")
        return

    with tempfile.TemporaryDirectory() as tmp:
        store = MockDataStore(ProjectStore(base_dir=Path(tmp) / "data"))
        store.add_project("Rename Identity Check")

        prepared = store.prepare_new_increment("Rename Identity Check", REAL_FIXTURE)
        real_record_name = prepared["record_name"]
        inc = store.finalize_new_increment(prepared, "SMR 3D")
        print(f"Created: real name={real_record_name!r}, display_name={inc.display_name!r}, version={inc.version}")

        if inc.display_name != "SMR 3D":
            failures.append(f"expected display_name 'SMR 3D', got {inc.display_name!r}")
        if inc.name != real_record_name:
            failures.append(f"identity (name) must stay the real Record Name, got {inc.name!r}")

        match = store.match_upload("Rename Identity Check", REAL_FIXTURE)
        print(f"Re-upload match_type: {match.match_type}")
        if match.match_type != MatchType.EXACT_MATCH:
            failures.append(
                f"renaming display_name must not affect identity matching -- expected EXACT_MATCH on the "
                f"same file, got {match.match_type}"
            )
        if match.matched_increment_name != real_record_name:
            failures.append(
                f"match_upload() must still resolve to the real Record Name, got {match.matched_increment_name!r}"
            )

        result = store.simulate_comparison("Rename Identity Check", inc.name, REAL_FIXTURE)
        if result.has_changes:
            failures.append(f"re-uploading the identical file should show no changes, got has_changes=True")

        inc2 = store.confirm_update("Rename Identity Check", inc.name, REAL_FIXTURE, result)
        print(f"After confirm_update: display_name={inc2.display_name!r}, version={inc2.version}")
        if inc2.version != 2:
            failures.append(f"expected version 2 after confirming the update, got {inc2.version}")
        if inc2.display_name != "SMR 3D":
            failures.append(f"display_name must be carried forward across a version bump, got {inc2.display_name!r}")

        # And on a fresh re-fetch from disk (not just the in-memory
        # objects just returned) -- the real round-trip through
        # increment.json.
        display = store.get_increment_for_display("Rename Identity Check", inc.name)
        print(f"Re-fetched from disk: display_name={display.display_name!r}, version={display.version}")
        if display.display_name != "SMR 3D":
            failures.append(f"display_name must round-trip through increment.json, got {display.display_name!r}")


def test_title_bar_shows_display_name_across_version_switch(failures):
    """DataViewPage's title bar must show display_name both on initial
    load (__init__) AND after switching versions via the combo
    (_apply_new_increment) -- a prior bug had the latter path regress
    to the full raw Record Name (300-550+ chars for real client files)
    even though __init__ correctly showed the short display_name.
    """
    QApplication.instance() or QApplication([])
    with tempfile.TemporaryDirectory() as tmp:
        store = MockDataStore(ProjectStore(base_dir=Path(tmp) / "data"))
        store.add_project("Title Bar Check")
        prepared = store.prepare_new_increment("Title Bar Check", FIXTURE)
        inc = store.finalize_new_increment(prepared, "Short Label")
        display = store.get_increment_for_display("Title Bar Check", inc.name)

        page = DataViewPage("Title Bar Check", display, store, on_back=lambda: None)
        print(f"Initial title: {page.title_label.text()!r}")
        if "Short Label" not in page.title_label.text():
            failures.append(f"initial title should show display_name 'Short Label', got {page.title_label.text()!r}")
        if page.title_label.toolTip() != inc.name:
            failures.append(f"initial title tooltip should be the full real name, got {page.title_label.toolTip()!r}")

        # Simulate a version switch: upload the same file again as a
        # new version, then push that new Increment through the same
        # path _on_row_version_changed/_apply_new_increment uses.
        store.confirm_update("Title Bar Check", inc.name, FIXTURE, store.simulate_comparison("Title Bar Check", inc.name, FIXTURE))
        v2_display = store.get_increment_for_display("Title Bar Check", inc.name, version=2)
        page._apply_new_increment(v2_display)
        print(f"After version switch: {page.title_label.text()!r}")
        if "Short Label" not in page.title_label.text():
            failures.append(f"after a version switch, title should still show display_name 'Short Label', got {page.title_label.text()!r} -- regressed to the raw Record Name")
        if page.title_label.toolTip() != inc.name:
            failures.append(f"after a version switch, title tooltip should still be the full real name, got {page.title_label.toolTip()!r}")


def test_sanitization(failures):
    cases = [
        ("  SMR 3D  ", "SMR 3D"),
        ("A" * 60, "A" * DISPLAY_NAME_MAX_LENGTH),
        ('bad/name:with*illegal<chars>?|"', "badnamewithillegalchars"),
        ("   ", None),
        ("", None),
        (None, None),
        ("multi   space    collapse", "multi space collapse"),
    ]
    for raw, expected in cases:
        got = _sanitize_display_name(raw)
        print(f"  _sanitize_display_name({raw!r}) -> {got!r} (expected {expected!r})")
        if got != expected:
            failures.append(f"_sanitize_display_name({raw!r}) expected {expected!r}, got {got!r}")


def test_blank_falls_back_to_derived_name(failures):
    with tempfile.TemporaryDirectory() as tmp:
        store = MockDataStore(ProjectStore(base_dir=Path(tmp) / "data"))
        store.add_project("Blank Name Check")
        prepared = store.prepare_new_increment("Blank Name Check", FIXTURE)
        derived = prepared["derived_display_name"]
        expected_derived = excel_reader.derive_short_display_name(prepared["record_name"])
        if derived != expected_derived:
            failures.append(f"prepare_new_increment()'s derived_display_name should match derive_short_display_name() directly, got {derived!r} vs {expected_derived!r}")

        # Blank string (user cleared the field) -> falls back.
        inc = store.finalize_new_increment(prepared, "")
        print(f"Blank display_name -> falls back to: {inc.display_name!r}")
        if inc.display_name != derived:
            failures.append(f"a blank display_name should fall back to the derived name {derived!r}, got {inc.display_name!r}")


def test_older_increment_without_display_name_still_works(failures):
    """An increment created with the OLD create_increment() signature
    (no display_name argument at all, like every increment created
    before this fix) must still work: IncrementRecord.display_name
    defaults to None, and ui.mock_data.Increment.display_name still
    resolves to a real, non-blank fallback.
    """
    with tempfile.TemporaryDirectory() as tmp:
        store = ProjectStore(base_dir=Path(tmp) / "data")
        store.create_project("Legacy Project")
        # No display_name kwarg at all -- simulates an increment created
        # by code that predates this field.
        store.create_increment("Legacy Project", "INC 1 - Legacy Item", FIXTURE)
        record = store.get_increment("Legacy Project", "INC 1 - Legacy Item")
        print(f"Legacy increment record.display_name: {record.display_name!r}")
        if record.display_name is not None:
            failures.append(f"a legacy increment's display_name should be None, got {record.display_name!r}")

        mock_store = MockDataStore(store)
        display = mock_store.get_increment_for_display("Legacy Project", "INC 1 - Legacy Item")
        print(f"Legacy increment's resolved display_name: {display.display_name!r}")
        if not display.display_name:
            failures.append("a legacy increment's resolved display_name must never be blank/None")


def main():
    failures = []
    test_display_name_never_affects_slug(failures)
    test_renamed_increment_still_matches_and_versions(failures)
    test_title_bar_shows_display_name_across_version_switch(failures)
    test_sanitization(failures)
    test_blank_falls_back_to_derived_name(failures)
    test_older_increment_without_display_name_still_works(failures)

    print("\n" + "=" * 70)
    if failures:
        print("RESULT: FAIL")
        for f in failures:
            print(" -", f)
        raise SystemExit(1)
    else:
        print(
            "RESULT: PASS -- a custom display_name never affects the folder slug, a renamed increment still "
            "exact-matches and versions correctly on re-upload, carries forward across version bumps, "
            "sanitizes correctly (trim/cap/illegal chars), falls back to the derived name when blank, and an "
            "older increment with no display_name at all still works."
        )


if __name__ == "__main__":
    main()
