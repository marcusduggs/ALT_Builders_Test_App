"""
Regression test for a direct client correction: the "Stage N Required
(Select with X)" column on grid sheets (B-Tests, C-On-Site Special
Inspections, D-Off-Site Special Inspections) only marks that a stage
APPLIES to an item -- it is not a completion signal. The real completion
signal is whether the rightmost "HCAI/FDD Use (Initial/Date)" column has
anything written in it (filled -> Done, blank -> Open), regardless of
whether the Required marker is "x" or "1".

core.excel_reader.raw_status_by_index() previously treated the Required
marker itself as the completion signal ("x" -> Done, "1" -> Open) -- a
PRIOR deliberate design choice, confirmed wrong by the client. Fixed by
reading a new BlueSheetRecord.hcai_fdd_use field (parsed from
GRID_HCAI_FDD_USE_COL_OFFSET = stage_last_col + 6) instead, while still
using the Required marker unchanged to decide WHICH stages are required.

Four things this file locks in, each confirmed against real HMR/SMR
fixture data (tests/fixtures/hmr_smr_projects/), not synthetic data:

1. The disagreement case that motivated the fix -- a real "x"-marked,
   HCAI/FDD-Use-blank row now reports Open, not Done.
2. The agreement case -- a real HCAI/FDD-Use-FILLED row still reports
   Done (confirming the fix isn't just "always Open").
3. The structural limitation the investigation surfaced and the client
   asked to have flagged, not silently decided: HCAI/FDD Use is ONE
   cell per physical row, not one per stage, so a row required for
   multiple stages gets the IDENTICAL Done/Open status for all of them.
4. VCR/milestone (F-Cons Verif) status, required-stage determination,
   stage-count detection, and Sum Data row generation are all
   completely unaffected by this fix -- it only touches grid-sheet
   completion status.

Also locks in HMR 10A's current totals (TIO_H240649-38-00 HMR 10A.xlsm)
as a KNOWN-BUGGY snapshot, not a correctness guarantee -- see
test_hmr10a_known_duplicate_index_bug_snapshot()'s own docstring.

Run directly: `python tests/test_hcai_fdd_use_status.py`
"""

import os
import sys
import tempfile
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core import excel_reader
from core.excel_reader import BlueSheetRecord
from core.project_store import ProjectStore
from ui.mock_data import MockDataStore

FIXDIR = os.path.join(os.path.dirname(__file__), "fixtures", "hmr_smr_projects")


def _load_display(fixture_name, project_name):
    with tempfile.TemporaryDirectory() as tmp:
        store = MockDataStore(ProjectStore(base_dir=Path(tmp) / "data"))
        store.add_project(project_name)
        increment = store.add_new_increment(project_name, os.path.join(FIXDIR, fixture_name))
        return store.get_increment_for_display(project_name, increment.name)


def _all_data_row(display, index):
    matches = [r for r in display.all_data if r["index"] == index]
    assert len(matches) == 1, f"expected exactly one All Data row for {index!r}, found {len(matches)}"
    return matches[0]


def test_required_marker_no_longer_drives_completion(failures):
    """Disagreement case: SMR 3D 2022 Delta 1's B-F1 has Stage 1 Required
    = "x" (raw cell, confirmed during investigation) and a BLANK HCAI/FDD
    Use cell. Before the fix this reported Done (driven by the "x");
    after the fix it must report Open.
    """
    display = _load_display("TIO H231182-38-00 SMR 3D 2022 Delta 1.xlsm", "Disagreement Case")
    row = _all_data_row(display, "B-F1")
    if row["required_stages"] != [1]:
        failures.append(f"B-F1 required_stages changed: expected [1], got {row['required_stages']}")
    if row["stage_status"] != {1: "Open"}:
        failures.append(f"B-F1 stage_status: expected {{1: 'Open'}} (HCAI/FDD Use blank), got {row['stage_status']}")


def test_hcai_fdd_use_filled_still_reports_done(failures):
    """Agreement case: SMR 2D's B-F1 has Stage 1 Required = "x" AND a
    real, filled-in HCAI/FDD Use cell ("BJS 9/25/2024", confirmed during
    investigation -- one of only 7 such cells found across all 15 real
    fixtures checked). Must still report Done -- the fix changes WHICH
    column drives completion, not a blanket flip to Open.
    """
    display = _load_display("TIO H220369-38-00_SMR 2D.xlsm", "Agreement Case")
    row = _all_data_row(display, "B-F1")
    if row["required_stages"] != [1]:
        failures.append(f"SMR 2D B-F1 required_stages changed: expected [1], got {row['required_stages']}")
    if row["stage_status"] != {1: "Done"}:
        failures.append(f"SMR 2D B-F1 stage_status: expected {{1: 'Done'}} (HCAI/FDD Use filled), got {row['stage_status']}")


def test_multi_required_stage_shares_one_hcai_cell(failures):
    """Structural limitation, flagged rather than silently decided: SMR
    4's B-F1 is required for 3 stages (confirmed real: [1, 2, 3]) but
    HCAI/FDD Use is a single cell per row, not one per stage -- so all 3
    required stages must come back with the IDENTICAL status (all Open
    here, since that row's HCAI/FDD Use is blank). There is no way, from
    this column alone, to mark stage 1 done but stage 2 not -- this test
    pins down that the app behaves that way deliberately, not as an
    undetected bug.
    """
    display = _load_display("TIO H220881-38-00_SMR 4.xlsm", "Multi-Stage Limitation")
    row = _all_data_row(display, "B-F1")
    if row["required_stages"] != [1, 2, 3]:
        failures.append(f"SMR 4 B-F1 required_stages changed: expected [1, 2, 3], got {row['required_stages']}")
    if row["stage_status"] != {1: "Open", 2: "Open", 3: "Open"}:
        failures.append(
            f"SMR 4 B-F1 stage_status: expected all 3 stages Open (one shared HCAI/FDD Use cell), "
            f"got {row['stage_status']}"
        )


def test_vcr_milestone_completely_unaffected(failures):
    """F-Cons Verif (VCR/milestone) items have their own, separate
    completion rule (vcr_marker_status(), unrelated to HCAI/FDD Use --
    grid records never even populate hcai_fdd_use at all). SMR 4C's F-3
    (confirmed during the earlier blank-row-exclusion fix to be the
    single real, non-blank F-3 entry) must still report vcr='Open' and
    required_stages=[] exactly as before this fix.
    """
    display = _load_display("TIO S251652-38-00 SMR 4C.xlsm", "VCR Unaffected Check")
    sum_data_f3 = [r for r in display.sum_data if r["index"] == "F-3"]
    if len(sum_data_f3) != 1:
        failures.append(f"SMR 4C F-3: expected exactly 1 sum_data row, got {len(sum_data_f3)}")
        return
    row = sum_data_f3[0]
    if row["required_stages"] != []:
        failures.append(f"SMR 4C F-3 required_stages changed: expected [], got {row['required_stages']}")
    if row["vcr"] != "Open":
        failures.append(f"SMR 4C F-3 vcr status changed: expected 'Open', got {row['vcr']!r}")


def test_stage_count_and_sum_data_row_generation_unaffected(failures):
    """Stage-count detection and Sum Data row expansion depend on the
    "Stage N Required" marker (required_stages), never on HCAI/FDD Use
    -- confirmed unaffected on SMR 4C: stage_count=3, 12 Sum Data rows,
    17 expanded on-screen (Index, Stage) rows, matching the values
    confirmed in the prior (blank-row-exclusion) fix's own verification.
    """
    display = _load_display("TIO S251652-38-00 SMR 4C.xlsm", "Row Generation Unaffected Check")
    if display.stage_count != 3:
        failures.append(f"SMR 4C stage_count changed: expected 3, got {display.stage_count}")
    if len(display.sum_data) != 12:
        failures.append(f"SMR 4C sum_data row count changed: expected 12, got {len(display.sum_data)}")
    expanded = excel_reader.expand_sum_data_rows(display.sum_data)
    if len(expanded) != 17:
        failures.append(f"SMR 4C expand_sum_data_rows() row count changed: expected 17, got {len(expanded)}")


def test_hmr10a_known_duplicate_index_bug_snapshot(failures):
    """NOT a correctness guarantee -- a snapshot of HMR 10A's CURRENT,
    KNOWN-BUGGY behavior, so this test starts failing (loudly, not
    silently) the moment core.excel_reader.required_stages_by_index()'s
    separate, already-identified last-write-wins duplicate-index bug is
    fixed in a later round, forcing this expected value to be revisited
    rather than left asserting a stale number forever.

    HMR 10A has 39 duplicate grid indices (one, B-C13, spans 9 physical
    rows) -- confirmed during this fix's own verification. Because
    required_stages_by_index() collapses each duplicate index's required
    stages to only its LAST physical row's list (and hands that same
    collapsed list back to every row sharing the index), the live app
    currently tracks only 49 of the real 130 required grid-stage cells
    in this file (confirmed via direct raw-cell count vs. live app
    count during this fix's own verification) -- entirely a pre-existing
    required_stages_by_index() artifact, unrelated to and unaffected by
    this HCAI/FDD-Use completion-source fix. All 49 are "Open" today
    (HMR 10A's real HCAI/FDD Use column is entirely blank in every row
    that currently makes it through the collapse) -- zero Done.
    """
    display = _load_display("TIO_H240649-38-00 HMR 10A.xlsm", "HMR 10A Known Bug Snapshot")
    done = open_count = 0
    for row in display.all_data:
        for stage in row["required_stages"]:
            status = row["stage_status"].get(stage)
            if status == "Done":
                done += 1
            elif status == "Open":
                open_count += 1
    if (done, open_count) != (0, 49):
        failures.append(
            f"HMR 10A grid Done/Open snapshot changed: expected (Done=0, Open=49), got (Done={done}, Open={open_count}). "
            f"If this changed because required_stages_by_index()'s duplicate-index bug was just fixed, "
            f"UPDATE this expected value -- do not just silence the failure."
        )


def test_raw_status_by_index_unit_semantics(failures):
    """Direct unit-level pin of core.excel_reader.raw_status_by_index()'s
    exact semantics on synthetic records, isolated from any real file --
    complements the real-fixture tests above with the edge cases they
    don't happen to exercise:
      - whitespace-only HCAI/FDD Use counts as blank (-> Open), not filled
      - a non-"x"/"1" marker ("NA") is still excluded entirely (same as
        before this fix -- stays unresolved, not auto-assigned a status)
      - a record with no index is skipped
    """
    records = [
        BlueSheetRecord(
            source_sheet="B-Tests", source_row=1, index="B-WS1",
            description="d", approval_agency="a", opaa_info=None,
            stage_values={1: "x"}, hcai_fdd_use="   ",  # whitespace-only
        ),
        BlueSheetRecord(
            source_sheet="B-Tests", source_row=2, index="B-NA1",
            description="d", approval_agency="a", opaa_info=None,
            stage_values={1: "NA"}, hcai_fdd_use="BJS 1/1/2025",
        ),
        BlueSheetRecord(
            source_sheet="B-Tests", source_row=3, index=None,
            description="d", approval_agency="a", opaa_info=None,
            stage_values={1: "x"}, hcai_fdd_use="BJS 1/1/2025",
        ),
    ]
    result = excel_reader.raw_status_by_index(records)
    if result.get("B-WS1") != {1: "Open"}:
        failures.append(f"whitespace-only HCAI/FDD Use should count as blank -> Open, got {result.get('B-WS1')}")
    if "B-NA1" in result:
        failures.append(f"'NA'-marked stage should stay excluded entirely (unresolved), got {result.get('B-NA1')}")
    if None in result or "None" in result:
        failures.append("a record with no index should never appear as a key in the result")


def main():
    failures = []

    # test_raw_status_by_index_unit_semantics is pure synthetic-record
    # logic (no file I/O at all) -- always runs, even if the real
    # fixtures aren't available locally. Every OTHER test function here
    # loads a real HMR/SMR file via FIXDIR, so those are skipped
    # together (not individually) when that folder is absent.
    if not os.path.isdir(FIXDIR):
        print(f"SKIP: {FIXDIR!r} not present (real HMR/SMR fixtures not available in this checkout) "
              f"-- running the fixture-independent unit test only")
    else:
        test_required_marker_no_longer_drives_completion(failures)
        test_hcai_fdd_use_filled_still_reports_done(failures)
        test_multi_required_stage_shares_one_hcai_cell(failures)
        test_vcr_milestone_completely_unaffected(failures)
        test_stage_count_and_sum_data_row_generation_unaffected(failures)
        test_hmr10a_known_duplicate_index_bug_snapshot(failures)
    test_raw_status_by_index_unit_semantics(failures)

    print("=" * 70)
    if failures:
        print("RESULT: FAIL")
        for f in failures:
            print(" -", f)
        raise SystemExit(1)
    else:
        print(
            "RESULT: PASS -- grid-row completion status now correctly comes from HCAI/FDD Use "
            "(not the Stage N Required marker) on real disagreement and agreement cases, the "
            "multi-required-stage/single-HCAI-cell limitation behaves as documented, VCR/milestone "
            "status and Sum Data row generation are unaffected, and HMR 10A's known "
            "required_stages_by_index()-bug-affected totals are pinned as a snapshot, not a guarantee."
        )


if __name__ == "__main__":
    main()
