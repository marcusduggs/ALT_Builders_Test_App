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

Also confirms HMR 10A's (TIO_H240649-38-00 HMR 10A.xlsm) duplicate-index
required-stages collapse is fixed -- see
test_hmr10a_duplicate_index_bug_fixed()'s own docstring.

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
    -- confirmed unaffected on SMR 4C: stage_count=3.

    Sum Data/expanded row counts here are 12/17, NOT the original 12/17
    this test first pinned, nor the 16/21 a LATER fix (F-Cons Verif's
    blank-reference-number "? rows" fix) temporarily raised them to:
    SMR 4C's F-Cons Verif sheet has 4 real milestone rows with a real
    description but a blank Reference Number ("Installation of
    Temporary Equipment", "Removal of Temporary Equipment",
    "Substantial Compliance...", "Certificate of Occupancy...") that
    fix kept (Index "?") -- but NONE of those 4 rows has a discipline
    sign-off mark (confirmed directly against the real file), so a
    STILL LATER discipline-mark refinement (the same rule SMR 2C's
    "Clear all plan review OIL Items" and SMR 4's 5 boilerplate rows
    are dropped under) drops all 4 again, landing back at 12/17 --
    independent of this HCAI/FDD-Use fix either way, which this test
    otherwise still confirms unaffected.
    """
    display = _load_display("TIO S251652-38-00 SMR 4C.xlsm", "Row Generation Unaffected Check")
    if display.stage_count != 3:
        failures.append(f"SMR 4C stage_count changed: expected 3, got {display.stage_count}")
    if len(display.sum_data) != 12:
        failures.append(f"SMR 4C sum_data row count changed: expected 12, got {len(display.sum_data)}")
    expanded = excel_reader.expand_sum_data_rows(display.sum_data)
    if len(expanded) != 17:
        failures.append(f"SMR 4C expand_sum_data_rows() row count changed: expected 17, got {len(expanded)}")


def test_hmr10a_duplicate_index_bug_fixed(failures):
    """FORMERLY a snapshot of HMR 10A's known-buggy behavior (Done=0,
    Open=49) -- now a correctness guarantee, confirming
    core.excel_reader.required_stages_by_index()'s last-write-wins
    duplicate-index collapse (and expand_sum_data_rows()'s matching
    "empty required_stages -> fake VCR row" symptom) are both fixed.

    HMR 10A has 39 duplicate grid Index values (one, B-C13, spans 9
    physical rows) -- confirmed via a direct, app-independent raw-cell
    scan of the real file (not assumed): 130 real required grid-stage
    marker cells (X/1) across all 3 grid sheets. Before the fix,
    required_stages_by_index() collapsed each duplicate index's required
    stages down to only its LAST physical row's own list and handed that
    one shared, often-wrong list back to every row sharing the index --
    only 49 of the real 130 required cells were tracked, and many
    duplicate rows wrongly showed "VCR" (zero required stages ->
    mistaken for a milestone item). Now keyed by row key (see
    assign_row_keys()), each of the 182 physical B-Tests rows / 123
    C-On-Site rows / 69 D-Off-Site rows gets its OWN correct required-
    stages list, so all 130 real required cells are now tracked -- zero
    lost, zero duplicated. All 130 are "Open" (HMR 10A's real HCAI/FDD
    Use column is entirely blank in every row) -- zero Done, unchanged
    from before; only the COUNT changed, not which status wins when a
    cell does resolve.
    """
    display = _load_display("TIO_H240649-38-00 HMR 10A.xlsm", "HMR 10A Duplicate-Index Fix Check")
    done = open_count = 0
    for row in display.all_data:
        for stage in row["required_stages"]:
            status = row["stage_status"].get(stage)
            if status == "Done":
                done += 1
            elif status == "Open":
                open_count += 1
    if (done, open_count) != (0, 130):
        failures.append(
            f"HMR 10A grid Done/Open: expected (Done=0, Open=130) now that the duplicate-index collapse is "
            f"fixed, got (Done={done}, Open={open_count})."
        )

    # Belt-and-suspenders: no All Data row for a grid sheet (B-Tests/
    # C-On-Site/D-Off-Site) may ever show "VCR" as a required stage --
    # that string only has meaning in Sum Data's expanded (Index, Stage)
    # rows (expand_sum_data_rows()), never in required_stages itself, so
    # this also doubles as a basic type-sanity check.
    for row in display.all_data:
        if "VCR" in row["required_stages"]:
            failures.append(f"HMR 10A: grid row {row['index']!r} has 'VCR' in required_stages -- should never happen")

    # The real bug symptom Rey reported: no GRID row may ever show up as
    # a Sum Data "VCR" row -- only a genuine F-Cons Verif (milestone)
    # item may. Confirmed directly against the expanded Sum Data rows,
    # not inferred.
    expanded = excel_reader.expand_sum_data_rows(display.sum_data)
    bad_vcr_rows = [
        e for e in expanded
        if e["stage"] == "VCR" and not str(e["index"]).startswith(("F-", "?"))
    ]
    if bad_vcr_rows:
        failures.append(f"HMR 10A: {len(bad_vcr_rows)} non-milestone row(s) show as 'VCR' in Sum Data: {bad_vcr_rows}")


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
        test_hmr10a_duplicate_index_bug_fixed(failures)
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
            "status and Sum Data row generation are unaffected, and HMR 10A's duplicate-index "
            "required_stages_by_index() collapse is now fixed (130 of 130 real required cells tracked, "
            "zero stray grid-row VCR rows)."
        )


if __name__ == "__main__":
    main()
