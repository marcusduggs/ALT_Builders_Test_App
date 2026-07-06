"""
End-to-end validation of the real (non-mock) upload -> review -> confirm ->
display loop at (Index, Stage) status granularity, using the actual
sample_increment.xlsm fixture and a temp-directory ProjectStore (never
touches the real ~/SubmissionAppData).

Exercises ui.mock_data.MockDataStore exactly the way the UI does -- status
marks are set via store.set_stage_status(), the same call
ui/pages/data_view_page.py's click handler makes, not by overwriting
status.json directly -- just without Qt in the loop, so failures are easy
to pinpoint. See main() for the scenarios this walks through (matching
the task's Step 4; a full screenshot-driven walkthrough of the real UI,
including actual clicks, is a separate manual script since this one has
no Qt event loop).

Reflects the later decision (see core.status_tracker's module docstring
and core.excel_reader.raw_status_by_index) to trust the raw X/1 marker
directly as Done/Open on upload, rather than leaving every required
stage unset until Rey's team acts in the app: a first upload now seeds
every required stage immediately, and store.set_stage_status() is used
here specifically to prove a manual override still wins and still
carries forward across a re-upload.

Run directly: `python tests/test_end_to_end_flow.py`
"""

import os
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core import excel_reader
from core.project_store import ProjectStore
from ui.mock_data import MockDataStore

FIXTURE = os.path.join(os.path.dirname(__file__), "fixtures", "sample_increment.xlsm")
PROJECT_NAME = "E2E Test Hospital - Validation Project"


def main():
    failures = []

    with tempfile.TemporaryDirectory() as tmp:
        project_store = ProjectStore(base_dir=Path(tmp) / "data")
        store = MockDataStore(project_store)

        # ------------------------------------------------------------
        # 4.1 -- first upload of a brand-new increment
        # ------------------------------------------------------------
        print("=" * 70)
        print("4.1 -- Upload sample_increment.xlsm as a new increment")
        print("=" * 70)

        store.add_project(PROJECT_NAME, "/tmp/e2e", 10)
        increment = store.add_new_increment(PROJECT_NAME, FIXTURE)
        print(f"Created increment {increment.name!r}, version {increment.version}")

        if increment.name != "INC 2 - Foundation and Underground Utilities":
            failures.append(f"4.1: increment name should come from A-Project Info's Record Name, got {increment.name!r}")
        if increment.version != 1:
            failures.append(f"4.1: first upload should be version 1, got {increment.version}")

        display = store.get_increment_for_display(PROJECT_NAME, increment.name)
        total_required_stages = sum(len(row["required_stages"]) for row in display.all_data)
        print(
            f"Total items: {display.total_count}, total required stages: {total_required_stages}, "
            f"needing status: {display.needs_status_stage_count} stage(s) across {display.needs_status_row_count} item(s), "
            f"flagged as changed: {display.changed_count}"
        )

        if display.total_count == 0:
            failures.append("4.1: All Data view has zero rows -- parsing did not actually run")
        if display.needs_status_stage_count != 0:
            failures.append(
                f"4.1: every required stage should already resolve from the file's raw X/1 marker "
                f"on a first upload -- needs_status_stage_count should be 0, got "
                f"{display.needs_status_stage_count} of {total_required_stages}"
            )
        if display.needs_status_row_count != 0:
            failures.append(
                f"4.1: no item should show needing status on a first upload, got "
                f"{display.needs_status_row_count}"
            )
        if display.changed_count != 0:
            failures.append(f"4.1: a first upload shouldn't flag anything as 'recently added', got {display.changed_count}")

        expected_raw_status = excel_reader.raw_status_by_index(
            excel_reader.parse_workbook(FIXTURE)["records"]
        )
        status_after_v1 = project_store.load_status(PROJECT_NAME, increment.name)
        if status_after_v1 != expected_raw_status:
            failures.append(
                "4.1: status.json after a first upload should exactly match "
                "raw_status_by_index() on the same file (fixture has no fill-color "
                "marks, so raw value alone should fully determine it)"
            )

        # ------------------------------------------------------------
        # 4.2 -- explicitly override a couple of already-auto-seeded
        # (index, stage) pairs via store.set_stage_status -- the exact
        # call the Data View's click handler makes, not a direct
        # status.json overwrite -- proving a deliberate human mark still
        # wins over whatever the raw file value auto-derived
        # ------------------------------------------------------------
        print("\n" + "=" * 70)
        print("4.2 -- Override specific (index, stage) pairs via set_stage_status")
        print("=" * 70)

        rows_by_index = {row["index"]: row for row in display.all_data}
        for idx in ("B-C1", "B-F1"):
            if not rows_by_index[idx]["required_stages"]:
                failures.append(f"4.2: fixture assumption broken -- {idx} has no required stages to mark")

        if failures:
            print("RESULT: FAIL (fixture assumptions broken, stopping early)")
            for f in failures:
                print(" -", f)
            raise SystemExit(1)

        bc1_stage = rows_by_index["B-C1"]["required_stages"][0]
        bf1_stage = rows_by_index["B-F1"]["required_stages"][0]
        # Flip whatever the raw value already auto-seeded, so this
        # provably an override, not a no-op re-assertion of the same value.
        bc1_override = "Open" if rows_by_index["B-C1"]["stage_status"].get(bc1_stage) == "Done" else "Done"
        bf1_override = "Open" if rows_by_index["B-F1"]["stage_status"].get(bf1_stage) == "Done" else "Done"
        print(f"Overriding B-C1 stage {bc1_stage} -> {bc1_override}, B-F1 stage {bf1_stage} -> {bf1_override}")

        store.set_stage_status(PROJECT_NAME, increment.name, "B-C1", bc1_stage, bc1_override)
        store.set_stage_status(PROJECT_NAME, increment.name, "B-F1", bf1_stage, bf1_override)

        status_after_marks = project_store.load_status(PROJECT_NAME, increment.name)
        if status_after_marks["B-C1"][bc1_stage] != bc1_override:
            failures.append(f"4.2: B-C1 stage {bc1_stage} should be overridden to {bc1_override}")
        if status_after_marks["B-F1"][bf1_stage] != bf1_override:
            failures.append(f"4.2: B-F1 stage {bf1_stage} should be overridden to {bf1_override}")
        # Every other (index, stage) pair's auto-seeded value should be
        # completely untouched by these two overrides.
        expected_after_marks = {
            index: {
                stage: (
                    bc1_override
                    if (index, stage) == ("B-C1", bc1_stage)
                    else bf1_override if (index, stage) == ("B-F1", bf1_stage) else value
                )
                for stage, value in stages.items()
            }
            for index, stages in expected_raw_status.items()
        }
        if status_after_marks != expected_after_marks:
            failures.append(
                f"4.2: only the two overridden (index, stage) pairs should differ from the "
                f"original auto-seeded status.json, got {status_after_marks}"
            )

        # ------------------------------------------------------------
        # 4.3 -- re-upload the SAME file as a new version: the two
        # overrides should carry forward untouched (not reset back to
        # the file's raw value), and everything else should be
        # unaffected -- an identical file introduces no new (index,
        # stage) requirements, so there's nothing left for
        # apply_file_derived_seed to seed here at all
        # ------------------------------------------------------------
        print("\n" + "=" * 70)
        print("4.3 -- Re-upload the identical file as a new version")
        print("=" * 70)

        result = store.simulate_comparison(PROJECT_NAME, increment.name, FIXTURE)
        print(f"has_changes: {result.has_changes} (expect False -- identical file, no new items)")
        if result.has_changes:
            failures.append(
                f"4.3: an identical file must show 'no changes detected' "
                f"(added={result.added_items}, removed={result.removed_items}, "
                f"anomalies={result.column_anomalies}, needs_status={result.needs_status_items})"
            )
        if result.carried_over_status != expected_after_marks:
            failures.append("4.3: carried_over_status should be exactly the post-override status, unchanged")

        increment = store.confirm_update(PROJECT_NAME, increment.name, FIXTURE, result)
        print(f"Confirmed -- now version {increment.version}")
        if increment.version != 2:
            failures.append(f"4.3: confirming should bump to version 2, got {increment.version}")

        status_after_confirm = project_store.load_status(PROJECT_NAME, increment.name)
        if status_after_confirm != expected_after_marks:
            failures.append(
                "4.3: status.json after confirming should still hold the overrides "
                "(carried forward, not reset back to the raw file value)"
            )

        display_v2 = store.get_increment_for_display(PROJECT_NAME, increment.name)
        rows_by_index_v2 = {row["index"]: row for row in display_v2.all_data}

        if rows_by_index_v2["B-C1"]["stage_status"].get(bc1_stage) != bc1_override:
            failures.append(f"4.3: B-C1 stage {bc1_stage} should still show the override {bc1_override!r}")
        if rows_by_index_v2["B-F1"]["stage_status"].get(bf1_stage) != bf1_override:
            failures.append(f"4.3: B-F1 stage {bf1_stage} should still show the override {bf1_override!r}")

        if display_v2.changed_count != 0:
            failures.append(
                f"4.3: v1 and v2 are byte-identical, nothing should be flagged 'recently added', "
                f"got {display_v2.changed_count}"
            )

        print(
            f"needing status after re-upload: {display_v2.needs_status_stage_count} (expected 0 -- "
            f"an identical file introduces no new required stages to leave unseeded)"
        )
        if display_v2.needs_status_stage_count != 0:
            failures.append(
                f"4.3: re-uploading an identical file shouldn't leave anything needing status, "
                f"got {display_v2.needs_status_stage_count}"
            )

    print("\n" + "=" * 70)
    print("SUMMARY")
    print("=" * 70)
    if failures:
        print("RESULT: FAIL")
        for f in failures:
            print(" -", f)
        raise SystemExit(1)
    else:
        print("RESULT: PASS")


if __name__ == "__main__":
    main()
