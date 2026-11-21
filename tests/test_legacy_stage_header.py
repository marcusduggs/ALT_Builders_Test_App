"""
Regression test for Bug B: core.excel_reader.detect_stage_count()'s
header regex only matched the newer "Stage N Required" column header.
Two real client files (tests/fixtures/hmr_smr_projects/ -- real HMR/SMR
TIO files, not synthetic data) use an OLDER HCAI template whose single,
un-numbered stage column reads "REQUIRED (Select with "X")" instead --
the regex silently matched zero columns, detect_stage_count() silently
returned 0, and every real required-stage marker in the sheet vanished
from Sum Data with no error at all (confirmed: all 296 real test/
inspection items disappeared this way for TIO S220804-38-00_SMR
2C_220705.xlsm specifically).

Investigation confirmed the old template's "REQUIRED" column carries
the exact same "X"/blank marker semantics as the new template's numbered
columns, and both affected real files' own Stage No. -> Stage Name
table (A-Project Info) is genuinely empty -- these are legitimately
single-stage projects under an older naming convention, not a different
concept needing separate handling. Fixed by recognizing BOTH header
forms as equivalent (a file using the legacy form always detects as
exactly 1 stage, since it's never numbered); a file matching NEITHER
form now raises ValueError instead of silently returning 0.

Run directly: `python tests/test_legacy_stage_header.py`
"""

import os
import sys
import tempfile
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import openpyxl
import pandas as pd

from core import excel_reader
from core.project_store import ProjectStore
from ui.mock_data import MockDataStore

FIXDIR = os.path.join(os.path.dirname(__file__), "fixtures", "hmr_smr_projects")
# The real file whose complete data loss (296 items) was directly
# confirmed during investigation.
SMR_2C = os.path.join(FIXDIR, "TIO S220804-38-00_SMR 2C_220705.xlsm")
# Confirmed to use the same legacy header during investigation, but
# untestable for Bug A until that fix landed (its Record Name is long
# enough to have crashed add_new_increment() outright before then).
SMR_2D = os.path.join(FIXDIR, "TIO H220369-38-00_SMR 2D.xlsm")
# An existing, already-tracked fixture using the NEWER "Stage N
# Required" template, 42 real stage columns -- confirms this fix didn't
# change behavior for the format that already worked.
NEW_TEMPLATE_SAMPLE = os.path.join(os.path.dirname(__file__), "fixtures", "sample_increment.xlsm")


def main():
    failures = []

    if not os.path.isdir(FIXDIR):
        print(f"SKIP: {FIXDIR!r} not present (real HMR/SMR fixtures not available in this checkout)")
        return

    # ------------------------------------------------------------
    # 1 -- TIO S220804-38-00_SMR 2C_220705.xlsm: the file with directly
    # confirmed 296-item data loss. Must now detect exactly 1 stage (not
    # 0, not raise) and recover real Sum Data rows instead of none.
    # ------------------------------------------------------------
    print("=" * 70)
    print("1 -- SMR 2C: previously silent-zero, directly confirmed 296-item data loss")
    print("=" * 70)
    parsed = excel_reader.normalize_workbook(SMR_2C)
    print(f"  stage_count: {parsed['stage_count']}")
    print(f"  All Data rows (real catalog items, unaffected by this bug): {len(parsed['all_data'])}")
    print(f"  Sum Data rows (previously 0 -- now the real required items): {len(parsed['sum_data'])}")
    if parsed["stage_count"] != 1:
        failures.append(f"1: SMR 2C should detect exactly 1 (legacy, un-numbered) stage, got {parsed['stage_count']}")
    # 297, not the original 296: a later fix stopped silently dropping
    # F-Cons Verif rows with a real description but a blank Reference
    # Number AND at least one discipline sign-off mark (see
    # core.excel_reader.parse_milestone_sheet's "?" index rule) -- SMR
    # 2C has exactly 1 such real row ("Construction Final", which has
    # discipline marks), now correctly included in All Data too
    # (build_all_data() includes every record unconditionally),
    # independent of this legacy-header fix. ("Clear all plan review
    # Outstanding Items List (OIL) Items" has NO discipline marks at
    # all and is excluded by a later discipline-mark refinement -- it
    # was briefly included too, at 298, between those two fixes.)
    if len(parsed["all_data"]) != 297:
        failures.append(f"1: fixture assumption broken -- expected 297 real catalog items, got {len(parsed['all_data'])}")
    if len(parsed["sum_data"]) == 0:
        failures.append("1: Sum Data is still empty -- the silent-zero bug has not been fixed for this file")

    # ------------------------------------------------------------
    # 2 -- TIO H220369-38-00_SMR 2D.xlsm: same legacy header, now
    # testable end to end since Bug A's fix no longer blocks it.
    # ------------------------------------------------------------
    print("\n" + "=" * 70)
    print("2 -- SMR 2D: same legacy template, previously untestable due to Bug A")
    print("=" * 70)
    parsed_2d = excel_reader.normalize_workbook(SMR_2D)
    print(f"  stage_count: {parsed_2d['stage_count']}")
    print(f"  All Data rows: {len(parsed_2d['all_data'])}")
    print(f"  Sum Data rows: {len(parsed_2d['sum_data'])}")
    if parsed_2d["stage_count"] != 1:
        failures.append(f"2: SMR 2D should detect exactly 1 (legacy, un-numbered) stage, got {parsed_2d['stage_count']}")
    if len(parsed_2d["sum_data"]) == 0:
        failures.append("2: Sum Data is still empty for SMR 2D")

    # ------------------------------------------------------------
    # 3 -- the legacy column's raw marker actually round-trips into a
    # real Done/Open status, not a blank/garbage value -- confirms the
    # fix recovers genuine data, not just a non-zero row count.
    # ------------------------------------------------------------
    print("\n" + "=" * 70)
    print("3 -- Recovered status values are real Done/Open, not blank/garbage")
    print("=" * 70)
    sum_data = parsed["sum_data"]
    statuses = set(sum_data["Stage 1"])
    print(f"  Distinct 'Stage 1' status values recovered for SMR 2C: {statuses}")
    # Grid-row values only: a milestone (F-Cons Verif) row's "Stage 1"
    # is always genuinely blank (these items track completion via VCR,
    # not the stage grid -- see required_stages_by_index()'s docstring)
    # and build_sum_data()'s raw DataFrame coerces that blank to NaN,
    # not None, once the column also holds a real status string --
    # confirmed as a pre-existing pandas quirk, not something this
    # fix/the legacy-header fix introduced: the same {..., nan} shape
    # already shows up today on SMR 3D 2022 Delta 1's own "Stage 1"
    # column, which has nothing to do with either fix. SMR 2C just
    # didn't have ANY milestone rows in Sum Data before a later fix (see
    # parse_milestone_sheet()'s "?" index rule) recovered its 2 real
    # blank-Reference-Number ones, so this is the first time SMR 2C's
    # own column happens to exercise it.
    #
    # Tightened, not loosened: every non-blank value must be EXACTLY
    # "Done"/"Open" (not just "is a string"), every blank must be a
    # REAL pandas NaN (not just "is a float", which would also accept a
    # garbage numeric value like 3.5), AND the NaN rows must actually
    # BE the milestone ones (Index "F-N" or "?") -- not some grid row
    # that unexpectedly lost its real marker. Checked both directions
    # so neither case can silently swap with the other.
    is_milestone_index = lambda idx: str(idx).startswith("F-") or idx == "?"
    bad_values = [v for v in statuses if v not in ("Done", "Open") and not pd.isna(v)]
    if bad_values:
        failures.append(f"3: found a status value that's neither 'Done'/'Open' nor a real NaN: {bad_values}")

    blank_rows = sum_data[sum_data["Stage 1"].isna()]
    non_milestone_blank = [idx for idx in blank_rows["Index"] if not is_milestone_index(idx)]
    if non_milestone_blank:
        failures.append(f"3: a GRID row's 'Stage 1' is blank -- expected only milestone rows to be, got {non_milestone_blank}")

    done_open_rows = sum_data[sum_data["Stage 1"].isin(["Done", "Open"])]
    milestone_with_real_status = [idx for idx in done_open_rows["Index"] if is_milestone_index(idx)]
    if milestone_with_real_status:
        failures.append(f"3: a milestone row has a real 'Stage 1' Done/Open value -- should always be blank, got {milestone_with_real_status}")

    if done_open_rows.empty:
        failures.append(f"3: expected at least one real 'Done'/'Open' status recovered, got {statuses}")

    # Restored, not just tightened: confirm this raw-DataFrame NaN can
    # never actually reach the LIVE Sum Data view/export -- the real
    # path (ui.mock_data.get_increment_for_display -> expand_sum_data_rows)
    # never reads build_sum_data()'s "Stage 1" column at all for a
    # milestone row; it only ever emits a single "VCR" row for one,
    # using the VCR column's own (always-resolved, never-blank)
    # vcr_marker_status() -- confirmed directly here, not assumed.
    with tempfile.TemporaryDirectory() as tmp:
        store = MockDataStore(ProjectStore(base_dir=Path(tmp) / "data"))
        store.add_project("Legacy Header Live Check")
        inc = store.add_new_increment("Legacy Header Live Check", SMR_2C)
        display = store.get_increment_for_display("Legacy Header Live Check", inc.name)
        expanded = excel_reader.expand_sum_data_rows(display.sum_data)
        milestone_numbered_rows = [e for e in expanded if e["index"] in ("?",) or str(e["index"]).startswith("F-")
                                    if e["stage"] != "VCR"]
        if milestone_numbered_rows:
            failures.append(f"3: a milestone row produced a numbered-stage row in the LIVE expansion (should only ever be 'VCR'): {milestone_numbered_rows}")
        live_statuses = {e["status"] for e in expanded if e["stage"] == "VCR"}
        if any(v not in ("Done", "Open", "NA") for v in live_statuses):
            failures.append(f"3: a live VCR row's status is something other than a resolved Done/Open/NA value: {live_statuses}")
        print(f"  Live VCR statuses for SMR 2C's recovered milestone rows: {live_statuses} (never blank/NaN)")

    # ------------------------------------------------------------
    # 4 -- the NEWER template (already-tracked fixture, 42 real stage
    # columns) is completely unaffected by widening the regex -- still
    # detects all 42, not 1, and the legacy pattern never matches a
    # "Stage N Required" column by accident.
    # ------------------------------------------------------------
    print("\n" + "=" * 70)
    print("4 -- Newer template (sample_increment.xlsm, 42 stages) unaffected")
    print("=" * 70)
    parsed_new = excel_reader.normalize_workbook(NEW_TEMPLATE_SAMPLE)
    print(f"  stage_count: {parsed_new['stage_count']}")
    if parsed_new["stage_count"] != 42:
        failures.append(f"4: sample_increment.xlsm should still detect 42 stages, got {parsed_new['stage_count']}")

    # ------------------------------------------------------------
    # 5 -- a genuinely unrecognized header (neither form) must now
    # raise ValueError, not silently return 0. No real fixture on hand
    # represents a third unknown template, so this is a constructed
    # worksheet -- it exercises the safety-net fallback path the two
    # real files above don't (both of them DO match the legacy
    # pattern); that path can only be tested with a header that matches
    # neither known form.
    # ------------------------------------------------------------
    print("\n" + "=" * 70)
    print("5 -- Truly unrecognized header raises, instead of silently returning 0")
    print("=" * 70)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.cell(row=5, column=excel_reader.GRID_INDEX_COL, value=excel_reader.GRID_HEADER_LABEL.title())
    ws.cell(row=5, column=excel_reader.GRID_STAGE_FIRST_COL, value="Some Completely Different Column Header")
    ws.cell(row=5, column=excel_reader.GRID_STAGE_FIRST_COL + 1, value="TESTS")
    try:
        count = excel_reader.detect_stage_count(ws, header_row=5)
        failures.append(f"5: expected ValueError for an unrecognized header, got a silent return of {count}")
    except ValueError as e:
        print(f"  Raised as expected: {e}")

    print("\n" + "=" * 70)
    if failures:
        print("RESULT: FAIL")
        for f in failures:
            print(" -", f)
        raise SystemExit(1)
    else:
        print("RESULT: PASS -- both real legacy-template files (SMR 2C, SMR 2D) now correctly detect 1 stage "
              "and recover real Done/Open Sum Data rows instead of silently losing all data; the newer "
              "42-stage template is completely unaffected; a truly unrecognized header now raises instead "
              "of silently returning 0.")


if __name__ == "__main__":
    main()
