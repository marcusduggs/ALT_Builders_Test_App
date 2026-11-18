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
from core.excel_reader import expand_sum_data_rows
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
        # doesn't raise and the right NUMBER of rows render. Sum Data
        # has NO Subtotal/Totals rows at all anymore (Rey's
        # TIO_Changes.xlsx "Sum Data" tab change request #3), so the
        # expected count is simply both increments' own real rows, with
        # nothing else mixed in -- a stronger, simpler guarantee than
        # the subtotal-value checks this test used to also make (there
        # is no longer a subtotal/grand-total VALUE to compute for Sum
        # Data at all).
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
            expected_row_count = len(d8_entries) + len(dsample_entries)
            print(f"Rendered Sum Data row count: {model.rowCount()} (expected {expected_row_count})")
            if model.rowCount() != expected_row_count:
                failures.append(
                    f"2: expected {expected_row_count} rendered Sum Data rows "
                    f"({len(d8_entries)} + {len(dsample_entries)}, no Subtotal/Totals rows), "
                    f"got {model.rowCount()}"
                )

            # Stage # column (index 5): INC 8's own rows must never show
            # a stage number past its real 2 -- the exact symptom the
            # original crash's root cause would have produced.
            STAGE_COL = 5
            INCREMENT_COL = 1
            bad_stage_rows = [
                r for r in range(model.rowCount())
                if model.data(model.index(r, INCREMENT_COL)) == d8.display_name
                and str(model.data(model.index(r, STAGE_COL))).isdigit()
                and int(model.data(model.index(r, STAGE_COL))) > 2
            ]
            if bad_stage_rows:
                failures.append(f"2: found INC 8 row(s) claiming a stage past its real 2: rows {bad_stage_rows}")

        # ------------------------------------------------------------
        # 3 -- the sibling EXPORT path (core.combined_export) -- confirm
        # it still produces the right rows end to end against the real
        # .xlsx it writes: no crash, a flat list only (no Subtotal/
        # Totals rows), exact row count, same stage-range guarantee as
        # step 2.
        # ------------------------------------------------------------
        print("\n" + "=" * 70)
        print("3 -- combined EXPORT path (core.combined_export)")
        print("=" * 70)

        export_path = str(Path(tmp) / "combined_mixed.xlsx")
        combined_export.export_combined_report([dsample, d8], export_path)  # must not raise

        wb = openpyxl.load_workbook(export_path)
        ws = wb["Sum Data"]
        increment_col = [
            ws.cell(row=r, column=excel_export.SUM_DATA_INCREMENT_COL).value for r in range(2, ws.max_row + 1)
        ]
        exported_row_count = len(increment_col)
        expected_exported_rows = len(d8_entries) + len(dsample_entries)
        print(f"Exported Sum Data row count: {exported_row_count} (expected {expected_exported_rows})")
        if exported_row_count != expected_exported_rows:
            failures.append(
                f"3: expected exactly {expected_exported_rows} exported Sum Data rows (no Subtotal/Totals "
                f"rows), got {exported_row_count}"
            )
        if any("Subtotal" in str(v) or v == "Totals" for v in increment_col):
            failures.append("3: exported Sum Data sheet must have NO Subtotal/Totals rows at all")

        STAGE_COL_EXPORT = excel_export.SUM_DATA_STAGE_COL
        bad_exported_stage_rows = [
            r for r in range(2, ws.max_row + 1)
            if ws.cell(row=r, column=excel_export.SUM_DATA_INCREMENT_COL).value == d8.display_name
            and isinstance(ws.cell(row=r, column=STAGE_COL_EXPORT).value, int)
            and ws.cell(row=r, column=STAGE_COL_EXPORT).value > 2
        ]
        if bad_exported_stage_rows:
            failures.append(f"3: found exported INC 8 row(s) claiming a stage past its real 2: rows {bad_exported_stage_rows}")

    print("\n" + "=" * 70)
    if failures:
        print("RESULT: FAIL")
        for f in failures:
            print(" -", f)
        raise SystemExit(1)
    else:
        print("RESULT: PASS -- combining a 2-stage and a 42-stage increment no longer has a wide "
              "stage-column grid to crash against at all; Sum Data's long (Index, Stage) row format "
              "renders and exports correctly for both increments as a flat list (no Subtotal/Totals "
              "rows), with no row ever claiming a stage past its own increment's real stage count.")


if __name__ == "__main__":
    main()
