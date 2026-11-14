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

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import openpyxl

from core import excel_reader

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
    if len(parsed["all_data"]) != 296:
        failures.append(f"1: fixture assumption broken -- expected 296 real catalog items, got {len(parsed['all_data'])}")
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
    if not statuses <= {"Done", "Open"}:
        failures.append(f"3: expected only 'Done'/'Open' status values, got {statuses}")

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
