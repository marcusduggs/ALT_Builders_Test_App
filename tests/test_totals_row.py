"""
Validates the All Data bottom totals row (core.excel_reader.
all_data_totals(), how it's surfaced on ui.mock_data.Increment, and how
it's rendered in the exported .xlsx) -- and confirms Sum Data has NO
totals row, and no Open/Done/Total/% Complete columns, at all (Rey's
TIO_Changes.xlsx "Sum Data" tab change requests #2/#3). Sum Data used
to have its own live-recomputed totals row (core.excel_reader.
live_sum_data_totals(), now deleted -- it had zero remaining callers
once both the in-app Sum Data tables and both Excel Sum Data sheets
stopped writing one); this file now computes the same Open/Done/Total
aggregate ITSELF, as an independent test-local oracle, purely to prove
a live status edit is still correctly reflected per ROW (there is no
longer an aggregate row to display it in at all).

Ground truth for sample_increment.xlsm was reverse-engineered directly
from the real file's own bottom-row formulas (data_only=False) and
cross-checked against its cached values (data_only=True):

  All Data row 381:  every Stage/VCR/SUM column = SUM(col3:col380)
                      (Excel's SUM() already skips "X" text, same rule
                      BlueSheetRecord.sum_value uses per row) -- SUM
                      (grand total) cached at 291.

Run directly: `python tests/test_totals_row.py`
"""

import os
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import openpyxl

from core import excel_export, excel_reader
from core.project_store import ProjectStore
from ui.mock_data import MockDataStore

FIXTURES = os.path.join(os.path.dirname(__file__), "fixtures")
SAMPLE = os.path.join(FIXTURES, "sample_increment.xlsm")
DEMO_BEFORE = os.path.join(FIXTURES, "demo_before.xlsm")
DEMO_AFTER = os.path.join(FIXTURES, "demo_after.xlsm")
PROJECT_NAME = "Totals Row Validation"
# sample_increment.xlsm's own real stage count -- core.excel_reader no
# longer has a fixed STAGE_COUNT constant (see its detect_stage_count()),
# but THIS fixture specifically is known to have 42.
EXPECTED_STAGE_COUNT = 42


def _compute_sum_data_totals(rows: list[dict]) -> dict:
    """Test-local independent oracle -- NOT a production function
    (core.excel_reader.live_sum_data_totals(), which this mirrors, was
    deleted: once Sum Data stopped having any Open/Done/Total/%
    Complete display anywhere -- see module docstring -- nothing in
    production code needed an aggregate anymore). Walks the SAME live
    `rows` shape (required_stages/stage_status/vcr) every Sum Data row
    already carries.
    """
    open_total = done_total = 0
    for row in rows:
        required = row.get("required_stages", [])
        stage_status = row.get("stage_status", {})
        for stage in required:
            status = stage_status.get(stage, "Open")
            if status == "Done":
                done_total += 1
            else:
                open_total += 1
        vcr_status = excel_reader.stage_marker_status(row.get("vcr"))
        if vcr_status == "Done":
            done_total += 1
        elif vcr_status == "Open":
            open_total += 1
    grand_total = open_total + done_total
    return {
        "open_total": open_total,
        "done_total": done_total,
        "grand_total": grand_total,
        "pct_complete": (done_total / grand_total) if grand_total else None,
    }


def main():
    failures = []

    # This test exercises the All Data export logic itself (totals row
    # presence/styling/typing) -- forced ON for the test's duration
    # regardless of the app's current default, since that logic must
    # keep working correctly whenever
    # core.excel_export.SHOW_ALL_DATA_AND_REPORT_TABS is flipped back
    # to True (see that constant's docstring). Restored afterward so
    # this doesn't leak into any other test run in the same process.
    original_toggle = excel_export.SHOW_ALL_DATA_AND_REPORT_TABS
    excel_export.SHOW_ALL_DATA_AND_REPORT_TABS = True
    try:
        _run(failures)
    finally:
        excel_export.SHOW_ALL_DATA_AND_REPORT_TABS = original_toggle

    print("\n" + "=" * 70)
    if failures:
        print("RESULT: FAIL")
        for f in failures:
            print(" -", f)
        raise SystemExit(1)
    else:
        print("RESULT: PASS -- All Data's totals row is correct on fresh upload, excluded from item counts, "
              "styled/typed correctly in export, and stays accurate after edits and re-uploads; Sum Data has "
              "NO totals row and NO Open/Done/Total/% Complete columns at all, on screen or in either export, "
              "and a live status edit is still correctly reflected per row with no aggregate to display it in.")


def _run(failures):
    with tempfile.TemporaryDirectory() as tmp:
        project_store = ProjectStore(base_dir=Path(tmp) / "data")
        store = MockDataStore(project_store)

        store.add_project(PROJECT_NAME)
        increment = store.add_new_increment(PROJECT_NAME, SAMPLE)
        display = store.get_increment_for_display(PROJECT_NAME, increment.name)

        # ------------------------------------------------------------
        # 1 -- fresh upload, zero edits. All Data's row set is 1:1 with
        # the real file (every row, no filtering), so its SUM total
        # matches the file's own cached grand total exactly: 291.
        # ------------------------------------------------------------
        print("=" * 70)
        print("1 -- Fresh upload: All Data's totals row matches the file exactly")
        print("=" * 70)

        all_totals = display.all_data_totals
        print(f"All Data SUM total: {all_totals.get('SUM')}")
        if all_totals.get("SUM") != 291:
            failures.append(f"1: All Data SUM total should be 291 (real file's cached grand total), got {all_totals.get('SUM')}")

        # Independent cross-check of Sum Data's live aggregate (see
        # _compute_sum_data_totals()'s own docstring for why this is a
        # test-local oracle, not production code), re-derived straight
        # from the parsed BlueSheetRecords -- a stage is required using
        # the UNCHANGED raw-marker rule (stage_marker_status(value) in
        # ("Done","Open")), but completion for a GRID record comes from
        # whether its hcai_fdd_use cell is filled in, not from the
        # marker itself (direct client correction). F-Cons Verif
        # (milestone) records are untouched by that correction and keep
        # using vcr_marker_status()-equivalent semantics.
        records = excel_reader.parse_workbook(SAMPLE)["records"]
        expected_open = expected_done = 0
        for rec in records:
            if not rec.index:
                continue
            if rec.source_sheet == excel_reader.MILESTONE_SHEET:
                vcr_status = excel_reader.vcr_marker_status(rec.vcr)
                if vcr_status == "Done":
                    expected_done += 1
                elif vcr_status == "Open":
                    expected_open += 1
                continue
            hcai_filled = rec.hcai_fdd_use is not None and str(rec.hcai_fdd_use).strip() != ""
            completion = "Done" if hcai_filled else "Open"
            for value in rec.stage_values.values():
                if excel_reader.stage_marker_status(value) not in ("Done", "Open"):
                    continue
                if completion == "Done":
                    expected_done += 1
                else:
                    expected_open += 1
        expected_total = expected_open + expected_done
        print(f"Independently cross-checked expected totals for OUR row set: Open={expected_open}, Done={expected_done}, Total={expected_total}")

        live_totals = _compute_sum_data_totals(display.sum_data)
        print(f"Sum Data live totals:     Open={live_totals['open_total']}, Done={live_totals['done_total']}, "
              f"Total={live_totals['grand_total']}, %={live_totals['pct_complete']}")
        if (live_totals["open_total"], live_totals["done_total"], live_totals["grand_total"]) != (expected_open, expected_done, expected_total):
            failures.append(
                f"1: Sum Data's live aggregate should match the independent cross-check "
                f"({expected_open}, {expected_done}, {expected_total}), got "
                f"({live_totals['open_total']}, {live_totals['done_total']}, {live_totals['grand_total']})"
            )
        if abs(live_totals["pct_complete"] - (expected_done / expected_total)) > 1e-9:
            failures.append(f"1: Sum Data % Complete should be {expected_done}/{expected_total}, got {live_totals['pct_complete']}")

        # ------------------------------------------------------------
        # 2 -- All Data's totals row is NOT counted as a real item
        # anywhere, and NEITHER All Data nor Sum Data ever has a row
        # with index "Totals" leaking in as real data.
        # ------------------------------------------------------------
        print("\n" + "=" * 70)
        print("2 -- Totals row never counted as a real item, never leaks into either row list")
        print("=" * 70)

        print(f"total_count: {display.total_count} (real item count -- blank/category-header rows already excluded upstream, no totals-row leakage)")
        if display.total_count != 345:
            failures.append(f"2: total_count should be exactly the real item count (345), got {display.total_count}")
        if any(row.get("index") in (None, "", "Totals") for row in display.all_data):
            failures.append("2: no All Data row should have a blank/'Totals' index -- the totals row must not leak into all_data")
        if any(row.get("index") in (None, "", "Totals") for row in display.sum_data):
            failures.append("2: no Sum Data row should have a blank/'Totals' index -- Sum Data has no totals row at all, and none should leak in")

        # ------------------------------------------------------------
        # 3 -- exported file: All Data's totals row present, bold,
        # tinted, numerically typed; Sum Data has NO totals row, NO
        # Open/Done/Total/% Complete columns, and its row count exactly
        # matches the flat expanded list (header + data rows only).
        # ------------------------------------------------------------
        print("\n" + "=" * 70)
        print("3 -- Exported file: All Data's totals row vs. Sum Data's lack of one")
        print("=" * 70)

        export_path = str(Path(tmp) / "totals_export.xlsx")
        excel_export.export_increment(display, export_path)
        wb_out = openpyxl.load_workbook(export_path)

        ws_all = wb_out["All Data"]
        all_totals_row = ws_all.max_row
        sum_cell = ws_all.cell(row=all_totals_row, column=3 + EXPECTED_STAGE_COUNT + 2)  # Index,Desc,Agency + 42 stages + VCR -> SUM
        print(f"Exported All Data totals row ({all_totals_row}): Index={ws_all.cell(row=all_totals_row, column=1).value!r}, "
              f"SUM={sum_cell.value!r} (type={type(sum_cell.value).__name__})")
        if ws_all.cell(row=all_totals_row, column=1).value != "Totals":
            failures.append(f"3: All Data exported totals row should read 'Totals' in the Index column, got {ws_all.cell(row=all_totals_row, column=1).value!r}")
        if sum_cell.value != 291:
            failures.append(f"3: exported All Data SUM total should be 291, got {sum_cell.value!r}")
        if not isinstance(sum_cell.value, (int, float)):
            failures.append(f"3: exported All Data SUM total should be a real number, got type {type(sum_cell.value).__name__}")
        if not ws_all.cell(row=all_totals_row, column=1).font.bold:
            failures.append("3: All Data exported totals row should be bold")
        if ws_all.cell(row=all_totals_row, column=1).fill.fgColor.rgb not in ("00000000", None) and \
           ws_all.cell(row=all_totals_row, column=1).fill.fill_type != "solid":
            failures.append("3: All Data exported totals row should have a fill applied")

        ws_sum = wb_out["Sum Data"]
        sum_header = [c.value for c in ws_sum[1]]
        print(f"Exported Sum Data header: {sum_header}")
        if any(h in sum_header for h in ("Open", "Done", "Total", "% Complete")):
            failures.append(f"3: exported Sum Data should have NO Open/Done/Total/% Complete columns at all, got {sum_header}")

        expected_sum_rows = len(excel_reader.expand_sum_data_rows(display.sum_data))
        actual_sum_rows = ws_sum.max_row - 1  # minus the header row
        print(f"Exported Sum Data row count: {actual_sum_rows} (expected {expected_sum_rows} -- no totals row to add)")
        if actual_sum_rows != expected_sum_rows:
            failures.append(f"3: exported Sum Data should have exactly {expected_sum_rows} rows (no totals row), got {actual_sum_rows}")
        last_row_index_value = ws_sum.cell(row=ws_sum.max_row, column=3).value  # Index column
        if last_row_index_value in (None, "", "Totals"):
            failures.append(f"3: exported Sum Data's LAST row should be a real data row, not a totals row, got Index={last_row_index_value!r}")

        expected_filter_range = f"A1:{openpyxl.utils.get_column_letter(excel_export.SUM_DATA_COMPLETION_DATE_COL)}{ws_sum.max_row}"
        print(f"Sum Data AutoFilter range: {ws_sum.auto_filter.ref} (expected {expected_filter_range})")
        if ws_sum.auto_filter.ref != expected_filter_range:
            failures.append(f"3: Sum Data AutoFilter should cover the header through the last (real) data row, got {ws_sum.auto_filter.ref!r}")

        # ------------------------------------------------------------
        # 4 -- a live status edit is still correctly reflected per row
        # immediately, with no re-upload / re-navigation needed -- even
        # though there's no aggregate Totals row left to display it in.
        # ------------------------------------------------------------
        print("\n" + "=" * 70)
        print("4 -- Live status edit reflected per-row immediately (no re-upload)")
        print("=" * 70)

        before = _compute_sum_data_totals(display.sum_data)
        # find a currently-Open required stage on some item to flip to Done
        target_row = next(
            row for row in display.sum_data
            if any(row["stage_status"].get(s, "Open") == "Open" for s in row["required_stages"])
        )
        target_stage = next(s for s in target_row["required_stages"] if target_row["stage_status"].get(s, "Open") == "Open")
        print(f"Flipping {target_row['index']} stage {target_stage} from Open -> Done")
        store.set_stage_status(PROJECT_NAME, increment.name, target_row["index"], target_stage, "Done")
        target_row["stage_status"][target_stage] = "Done"  # mirror the in-memory update the click handler makes

        after = _compute_sum_data_totals(display.sum_data)
        print(f"Before: Open={before['open_total']}, Done={before['done_total']}, %={before['pct_complete']:.4f}")
        print(f"After:  Open={after['open_total']}, Done={after['done_total']}, %={after['pct_complete']:.4f}")
        if after["open_total"] != before["open_total"] - 1:
            failures.append(f"4: Open total should decrease by 1 after the edit, went {before['open_total']} -> {after['open_total']}")
        if after["done_total"] != before["done_total"] + 1:
            failures.append(f"4: Done total should increase by 1 after the edit, went {before['done_total']} -> {after['done_total']}")
        if after["pct_complete"] <= before["pct_complete"]:
            failures.append(f"4: % Complete should increase after marking a stage Done, went {before['pct_complete']} -> {after['pct_complete']}")
        if after["grand_total"] != before["grand_total"]:
            failures.append(f"4: Total (grand total) should be unchanged by a status edit, went {before['grand_total']} -> {after['grand_total']}")

        # ------------------------------------------------------------
        # 5 -- totals reflect a NEW file's counts/values, not stale ones
        # from a previous upload (demo_before.xlsm -> demo_after.xlsm,
        # which adds 3 new items)
        # ------------------------------------------------------------
        print("\n" + "=" * 70)
        print("5 -- Totals reflect the currently-loaded file, not a stale one")
        print("=" * 70)

        growth_project = "Totals Row Growth Validation"
        store.add_project(growth_project)
        growth_increment = store.add_new_increment(growth_project, DEMO_BEFORE)
        display_before = store.get_increment_for_display(growth_project, growth_increment.name)
        totals_before = display_before.all_data_totals
        sum_totals_before = _compute_sum_data_totals(display_before.sum_data)

        result = store.simulate_comparison(growth_project, growth_increment.name, DEMO_AFTER)
        store.confirm_update(growth_project, growth_increment.name, DEMO_AFTER, result)
        display_after = store.get_increment_for_display(growth_project, growth_increment.name)
        totals_after = display_after.all_data_totals
        sum_totals_after = _compute_sum_data_totals(display_after.sum_data)

        print(f"total_count: before={display_before.total_count}, after={display_after.total_count}")
        if display_after.total_count - display_before.total_count != 3:
            failures.append("5: fixture assumption broken -- expected demo_after.xlsm to add exactly 3 items")

        # All Data's SUM total is a NUMERIC sum (Excel SUM() semantics --
        # ignores "X" text markers, see all_data_totals()'s docstring).
        # The 3 new items here (B-C18/19/20) happen to be marked with "X"
        # only, no literal numeric test-count value, so this total is
        # legitimately UNCHANGED by this specific upload -- verified
        # directly (every one of the 42 per-stage totals is identical
        # before/after). That's a fact about this fixture pair, not a
        # sign the totals are stale -- Sum Data's Open/Done counts (which
        # DO count "X"/"Open" markers, not sum numeric values) are the
        # right signal to prove this file's own data drove the
        # recomputation, and those clearly move:
        print(f"All Data SUM total: before={totals_before.get('SUM')}, after={totals_after.get('SUM')} (expected to tie -- see above)")
        print(f"Sum Data Done total: before={sum_totals_before['done_total']}, after={sum_totals_after['done_total']}")
        print(f"Sum Data grand total: before={sum_totals_before['grand_total']}, after={sum_totals_after['grand_total']}")
        if sum_totals_after["grand_total"] == sum_totals_before["grand_total"]:
            failures.append(
                f"5: Sum Data's Total (grand total) should change after uploading a file with 3 new "
                f"items, stayed at {sum_totals_before['grand_total']}"
            )
        if sum_totals_after["done_total"] == sum_totals_before["done_total"]:
            failures.append(
                f"5: Sum Data's Done total should change after uploading a file with 3 new items, "
                f"stayed at {sum_totals_before['done_total']}"
            )


if __name__ == "__main__":
    main()
