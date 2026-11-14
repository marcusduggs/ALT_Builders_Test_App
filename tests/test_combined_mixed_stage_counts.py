"""
Regression test for a real reported crash: combining INC 2 (sample_increment.xlsm,
42 stages) with INC 8 (inc8_real_2stage.xlsm, 2 stages) via View Combined Data
raised `KeyError: 3` in ui.pages.combined_data_view_page._build_sum_data_tab's
per-increment subtotal row, at `totals["stage_open_counts"][stage]`.

Root cause (historical): Sum Data used to be WIDE -- one row per item, with a
fixed "Stage 1".."Stage N" column range sized to the sheet's union max stage
count (here 42, from INC 2). core.combined_export.build_combined_sum_data_rows()
computed each increment's own "subtotal_totals" with THAT increment's own
(possibly narrower) stage_count, so INC 8's stage_open_counts dict only ever
had keys {1, 2} -- but the Sum Data tab indexed it across the full union range
and crashed on stage 3.

Sum Data is now long-format (core.excel_reader.expand_sum_data_rows): one row
per (Index, Stage) actually required for THAT item, with no fixed stage-column
grid at all -- there's nothing for a narrower increment's rows to misalign
against, so this exact bug class can no longer occur by construction. This
test instead confirms the fix's INTENT still holds in the new shape:
combining these two increments doesn't crash, INC 8's own rows never claim a
stage past its own real 2, and every per-increment subtotal / overall Totals
number is correct -- both on screen (the exact widget construction that used
to crash) and in the export.

Run directly: `python tests/test_combined_mixed_stage_counts.py`
"""

import os
import sys
import tempfile
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import openpyxl
from PySide6.QtWidgets import QApplication

from core import combined_export, excel_export
from core.excel_reader import derive_short_display_name, expand_sum_data_rows, live_sum_data_totals
from core.project_store import ProjectStore
from ui.mock_data import MockDataStore, build_combined_view
from ui.pages.combined_data_view_page import CombinedDataViewPage

FIXTURES = os.path.join(os.path.dirname(__file__), "fixtures")
INC8_2STAGE = os.path.join(FIXTURES, "inc8_real_2stage.xlsm")
SAMPLE_42STAGE = os.path.join(FIXTURES, "sample_increment.xlsm")
PROJECT_NAME = "Mixed Stage Count Validation"


def main():
    failures = []
    with tempfile.TemporaryDirectory() as tmp:
        store = MockDataStore(ProjectStore(base_dir=Path(tmp) / "data"))
        store.add_project(PROJECT_NAME)

        inc8 = store.add_new_increment(PROJECT_NAME, INC8_2STAGE)
        inc_sample = store.add_new_increment(PROJECT_NAME, SAMPLE_42STAGE)
        d8 = store.get_increment_for_display(PROJECT_NAME, inc8.name)
        dsample = store.get_increment_for_display(PROJECT_NAME, inc_sample.name)

        print(f"INC 8 ({d8.name}): stage_count={d8.stage_count}")
        print(f"INC 2 ({dsample.name}): stage_count={dsample.stage_count}")
        if d8.stage_count >= dsample.stage_count:
            failures.append(
                f"setup: fixture assumption broken -- expected INC 8's stage_count ({d8.stage_count}) "
                f"< INC 2's ({dsample.stage_count})"
            )

        # ------------------------------------------------------------
        # 1 -- build_combined_view() itself must not raise, and its union
        # stage_count must be the WIDER increment's, not the narrower one's
        # (still meaningful for All Data, which stays wide/unchanged).
        # ------------------------------------------------------------
        view = build_combined_view([dsample, d8])
        print(f"Combined view stage_count (union): {view.stage_count}")
        if view.stage_count != dsample.stage_count:
            failures.append(
                f"1: expected combined stage_count to be INC 2's {dsample.stage_count} (the union max), "
                f"got {view.stage_count}"
            )

        d8_entries = expand_sum_data_rows(d8.sum_data)
        dsample_entries = expand_sum_data_rows(dsample.sum_data)
        d8_stages = {entry["stage"] for entry in d8_entries if entry["stage"] != "VCR"}
        print(f"INC 8's own real Stage # values (from its Sum Data rows): {sorted(d8_stages)}")
        if any(isinstance(s, int) and s > 2 for s in d8_stages):
            failures.append(
                f"1: fixture/behavior assumption broken -- INC 8 only has 2 real stages, but its "
                f"expanded Sum Data rows claim a stage number past that: {sorted(d8_stages)}"
            )

        # ------------------------------------------------------------
        # 2 -- instantiate the REAL CombinedDataViewPage (the actual
        # crash site: _build_sum_data_tab runs from __init__) and read
        # its actual QStandardItemModel back -- confirms construction
        # doesn't raise, the right NUMBER of rows render (both
        # increments' own (Index, Stage) rows + 2 subtotal rows + 1
        # grand-totals row), and the subtotal/grand-total VALUES are
        # correct -- not just that nothing crashed.
        # ------------------------------------------------------------
        print("\n" + "=" * 70)
        print("2 -- CombinedDataViewPage's real Sum Data rows (the crash site)")
        print("=" * 70)

        app = QApplication.instance() or QApplication(sys.argv)
        try:
            page = CombinedDataViewPage(PROJECT_NAME, view, store, on_back=lambda: None)
        except KeyError as e:
            failures.append(
                f"2: constructing CombinedDataViewPage raised KeyError({e}) -- a narrower increment's "
                "own data should never be indexed against a wider union range in the long-format layout"
            )
            page = None

        if page is not None:
            model = page.sum_data_table.model()
            expected_row_count = len(d8_entries) + len(dsample_entries) + 2 + 1  # +2 subtotals, +1 Totals
            print(f"Rendered Sum Data row count: {model.rowCount()} (expected {expected_row_count})")
            if model.rowCount() != expected_row_count:
                failures.append(
                    f"2: expected {expected_row_count} rendered Sum Data rows "
                    f"({len(d8_entries)} + {len(dsample_entries)} real + 2 subtotal + 1 Totals), "
                    f"got {model.rowCount()}"
                )

            def _find_row(label: str, col: int = 1):
                # model.data(), not model.item() -- page.sum_data_table.model()
                # is now a QSortFilterProxyModel (the Sum Data filter bar's
                # proxy, see ui.widgets.sum_data_filter_bar), which has no
                # .item() (that's QStandardItemModel-specific); .data()
                # works identically on both.
                for r in range(model.rowCount()):
                    if model.data(model.index(r, col)) == label:
                        return r
                return None

            # Preview Subtotal labels use the SHORT display name (see
            # core.excel_reader.derive_short_display_name), a labels-
            # only exception to preview/export parity -- the EXPORTED
            # label (checked in step 3 below) still uses the full name
            # unchanged.
            inc8_subtotal_row = _find_row(f"{derive_short_display_name(d8.name)} — Subtotal")
            inc2_subtotal_row = _find_row(f"{derive_short_display_name(dsample.name)} — Subtotal")
            totals_row = _find_row("Totals")
            if inc8_subtotal_row is None or inc2_subtotal_row is None or totals_row is None:
                failures.append("2: expected a subtotal row for each increment plus one overall Totals row")
            else:
                # Project#,Increment,Index,Description,Agency,Stage#,Status,Notes,CompletionDate,Open,Done,Total,%
                OPEN_COL, DONE_COL, TOTAL_COL = 9, 10, 11
                d8_totals = live_sum_data_totals(d8.sum_data, d8.stage_count)
                dsample_totals = live_sum_data_totals(dsample.sum_data, dsample.stage_count)
                combined_totals = combined_export.combine_sum_data_totals([dsample, d8])

                inc8_rendered_total = int(model.data(model.index(inc8_subtotal_row, TOTAL_COL)))
                inc2_rendered_total = int(model.data(model.index(inc2_subtotal_row, TOTAL_COL)))
                grand_rendered_total = int(model.data(model.index(totals_row, TOTAL_COL)))
                print(f"INC 8 subtotal Total: {inc8_rendered_total} (expected {d8_totals['grand_total']})")
                print(f"INC 2 subtotal Total: {inc2_rendered_total} (expected {dsample_totals['grand_total']})")
                print(f"Overall Totals row Total: {grand_rendered_total} (expected {combined_totals['grand_total']})")
                if inc8_rendered_total != d8_totals["grand_total"]:
                    failures.append(
                        f"2: INC 8's rendered subtotal Total should be {d8_totals['grand_total']}, got {inc8_rendered_total}"
                    )
                if inc2_rendered_total != dsample_totals["grand_total"]:
                    failures.append(
                        f"2: INC 2's rendered subtotal Total should be {dsample_totals['grand_total']}, got {inc2_rendered_total}"
                    )
                if grand_rendered_total != combined_totals["grand_total"]:
                    failures.append(
                        f"2: overall rendered Totals row Total should be {combined_totals['grand_total']}, got {grand_rendered_total}"
                    )

        # ------------------------------------------------------------
        # 3 -- the sibling EXPORT path (core.combined_export) -- confirm
        # it still produces the right values end to end against the real
        # .xlsx it writes, using the new long-format column constants.
        # ------------------------------------------------------------
        print("\n" + "=" * 70)
        print("3 -- combined EXPORT path (core.combined_export)")
        print("=" * 70)

        export_path = str(Path(tmp) / "combined_mixed.xlsx")
        combined_export.export_combined_report([dsample, d8], export_path)  # must not raise

        wb = openpyxl.load_workbook(export_path)
        ws = wb["Sum Data"]
        increment_col = [
            ws.cell(row=r, column=excel_export.SUM_DATA_INCREMENT_COL).value for r in range(1, ws.max_row + 1)
        ]
        inc8_label = f"{d8.name} — Subtotal"
        inc2_label = f"{dsample.name} — Subtotal"
        if inc8_label not in increment_col or inc2_label not in increment_col:
            failures.append("3: expected both increments' subtotal rows in the exported Sum Data sheet")
        else:
            inc8_row = increment_col.index(inc8_label) + 1
            inc2_row = increment_col.index(inc2_label) + 1
            d8_totals = live_sum_data_totals(d8.sum_data, d8.stage_count)
            dsample_totals = live_sum_data_totals(dsample.sum_data, dsample.stage_count)
            inc8_exported_total = ws.cell(row=inc8_row, column=excel_export.SUM_DATA_TOTAL_COL).value
            inc2_exported_total = ws.cell(row=inc2_row, column=excel_export.SUM_DATA_TOTAL_COL).value
            print(f"Exported INC 8 subtotal Total: {inc8_exported_total} (expected {d8_totals['grand_total']})")
            print(f"Exported INC 2 subtotal Total: {inc2_exported_total} (expected {dsample_totals['grand_total']})")
            if inc8_exported_total != d8_totals["grand_total"]:
                failures.append(f"3: exported INC 8 subtotal Total should be {d8_totals['grand_total']}, got {inc8_exported_total}")
            if inc2_exported_total != dsample_totals["grand_total"]:
                failures.append(f"3: exported INC 2 subtotal Total should be {dsample_totals['grand_total']}, got {inc2_exported_total}")

            combined_totals = combined_export.combine_sum_data_totals([dsample, d8])
            totals_row = ws.max_row
            exported_grand_total = ws.cell(row=totals_row, column=excel_export.SUM_DATA_TOTAL_COL).value
            print(f"Exported overall Totals row Total: {exported_grand_total} (expected {combined_totals['grand_total']})")
            if exported_grand_total != combined_totals["grand_total"]:
                failures.append(
                    f"3: exported overall Totals Total should be {combined_totals['grand_total']}, got {exported_grand_total}"
                )

    print("\n" + "=" * 70)
    if failures:
        print("RESULT: FAIL")
        for f in failures:
            print(" -", f)
        raise SystemExit(1)
    else:
        print("RESULT: PASS -- combining a 2-stage and a 42-stage increment no longer has a wide "
              "stage-column grid to crash against at all; Sum Data's long (Index, Stage) row format "
              "renders and exports correctly for both increments, with correct per-increment subtotal "
              "and overall Totals numbers.")


if __name__ == "__main__":
    main()
