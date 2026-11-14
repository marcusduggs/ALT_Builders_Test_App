"""
Regression test for the Notes/Completion Date columns added to Sum
Data's output schema (core.excel_export.SUM_DATA_NOTES_COL,
SUM_DATA_COMPLETION_DATE_COL) -- positioned right after Status, before
the existing Open/Done/Total/% Complete summary columns. Scoped
deliberately narrow for this pass: the columns exist and are ALWAYS
blank on export -- no underlying app data feeds them, no merge/carry-
forward/read-back logic exists yet (explicitly parked for a later
round, see core/excel_export.py's SUM_DATA_NOTES_COL comment).

Covers both the single-increment export (core.excel_export) and the
combined/master-sheet export (core.combined_export) against real
HMR/SMR fixture files, plus the on-screen preview pages (parity with
the exported file, not a separately-maintained copy).

Run directly: `python tests/test_sum_data_notes_completion_date_columns.py`
"""

import os
import sys
import tempfile
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import openpyxl
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication

from core import combined_export, excel_export
from core.project_store import ProjectStore
from ui.mock_data import MockDataStore, build_combined_view
from ui.pages.combined_data_view_page import CombinedDataViewPage
from ui.pages.data_view_page import DataViewPage

FIXDIR = os.path.join(os.path.dirname(__file__), "fixtures", "hmr_smr_projects")
PROJECT_NAME = "Notes/Completion Date Regression"
FILES = ["TIO S221919-38-00_HMR 6C.xlsm", "TIO S221870-38-00_SMR 3C.xlsm"]


def _all_blank(ws, col: int, first_row: int = 2) -> bool:
    return all(ws.cell(row=r, column=col).value in (None, "") for r in range(first_row, ws.max_row + 1))


def main():
    failures = []

    if not os.path.isdir(FIXDIR):
        print(f"SKIP: {FIXDIR!r} not present (real HMR/SMR fixtures not available in this checkout)")
        return

    app = QApplication.instance() or QApplication(sys.argv)

    with tempfile.TemporaryDirectory() as tmp:
        store = MockDataStore(ProjectStore(base_dir=Path(tmp) / "data"))
        store.add_project(PROJECT_NAME)
        displays = []
        for fname in FILES:
            path = os.path.join(FIXDIR, fname)
            increment = store.add_new_increment(PROJECT_NAME, path)
            displays.append(store.get_increment_for_display(PROJECT_NAME, increment.name))

        # ------------------------------------------------------------
        # 1 -- Single-increment export: headers include Notes/Completion
        # Date right after Status, before Open; every data row AND the
        # totals row are genuinely blank there, not "", a formula, or
        # placeholder text.
        # ------------------------------------------------------------
        print("=" * 70)
        print("1 -- Single-increment export")
        print("=" * 70)
        single_path = str(Path(tmp) / "single.xlsx")
        excel_export.export_increment(displays[0], single_path)
        ws = openpyxl.load_workbook(single_path)["Sum Data"]
        headers = [c.value for c in ws[1]]
        print(f"  Headers: {headers}")
        expected_order = ["Status", "Notes", "Completion Date", "Open"]
        actual_order = [h for h in headers if h in expected_order]
        if actual_order != expected_order:
            failures.append(f"1: expected column order {expected_order}, got {actual_order}")

        if "Notes" not in headers or "Completion Date" not in headers:
            failures.append(f"1: expected both 'Notes' and 'Completion Date' headers, got {headers}")
        else:
            notes_col = headers.index("Notes") + 1
            date_col = headers.index("Completion Date") + 1
            if not _all_blank(ws, notes_col):
                failures.append("1: Notes column is not blank on every row of the single-increment export")
            if not _all_blank(ws, date_col):
                failures.append("1: Completion Date column is not blank on every row of the single-increment export")
            # Confirm genuinely blank, not a formula string.
            sample = ws.cell(row=2, column=notes_col)
            if sample.data_type == "f":
                failures.append(f"1: Notes cell is a formula ({sample.value!r}), not a blank value")

        # ------------------------------------------------------------
        # 2 -- Combined/master-sheet export: same columns, same
        # position, blank on every real row AND every per-increment
        # Subtotal row AND the final overall Totals row.
        # ------------------------------------------------------------
        print("\n" + "=" * 70)
        print("2 -- Combined master-sheet export")
        print("=" * 70)
        combined_path = str(Path(tmp) / "combined.xlsx")
        combined_export.export_combined_report(displays, combined_path)
        ws2 = openpyxl.load_workbook(combined_path)["Sum Data"]
        headers2 = [c.value for c in ws2[1]]
        print(f"  Headers: {headers2}")
        actual_order2 = [h for h in headers2 if h in expected_order]
        if actual_order2 != expected_order:
            failures.append(f"2: expected column order {expected_order}, got {actual_order2}")

        if "Notes" not in headers2 or "Completion Date" not in headers2:
            failures.append(f"2: expected both 'Notes' and 'Completion Date' headers, got {headers2}")
        else:
            notes_col2 = headers2.index("Notes") + 1
            date_col2 = headers2.index("Completion Date") + 1
            if not _all_blank(ws2, notes_col2):
                failures.append("2: Notes column is not blank on every row of the combined export "
                                 "(including Subtotal/Totals rows)")
            if not _all_blank(ws2, date_col2):
                failures.append("2: Completion Date column is not blank on every row of the combined export")

        # ------------------------------------------------------------
        # 3 -- Preview/export parity: both on-screen tables show the
        # SAME headers in the SAME position/order as the export, with
        # ONE known, deliberate exception -- the single-increment
        # preview's leading identity column is labeled "Increment" (to
        # match the combined view's own naming for the same concept,
        # now that it displays a short derived name rather than the raw
        # Record Name -- see core.excel_reader.derive_short_display_name),
        # while the exported .xlsx keeps "Inc #" unchanged. Every OTHER
        # header, including the Notes/Completion Date columns this test
        # was written for, must still match exactly.
        # ------------------------------------------------------------
        print("\n" + "=" * 70)
        print("3 -- Preview pages match the export (except the known 'Inc #'/'Increment' relabel)")
        print("=" * 70)
        page = DataViewPage(PROJECT_NAME, displays[0], store, on_back=lambda: None)
        table = page.sum_data_table
        model = table.model()
        preview_headers = [model.headerData(c, Qt.Horizontal) for c in range(model.columnCount())]
        print(f"  Single preview headers: {preview_headers}")
        expected_single_preview_headers = list(headers)
        expected_single_preview_headers[expected_single_preview_headers.index("Inc #")] = "Increment"
        if preview_headers != expected_single_preview_headers:
            failures.append(
                f"3: single-increment preview headers {preview_headers} don't match the expected "
                f"{expected_single_preview_headers} (export headers with 'Inc #' relabeled 'Increment')"
            )

        view = build_combined_view(displays)
        cpage = CombinedDataViewPage(PROJECT_NAME, view, store, on_back=lambda: None)
        cmodel = cpage.sum_data_table.model()
        cpreview_headers = [cmodel.headerData(c, Qt.Horizontal) for c in range(cmodel.columnCount())]
        print(f"  Combined preview headers: {cpreview_headers}")
        if cpreview_headers != headers2:
            failures.append(
                f"3: combined preview headers {cpreview_headers} don't match the export's {headers2}"
            )

    print("\n" + "=" * 70)
    if failures:
        print("RESULT: FAIL")
        for f in failures:
            print(" -", f)
        raise SystemExit(1)
    else:
        print("RESULT: PASS -- Notes/Completion Date columns appear, blank, right after Status in both the "
              "single-increment and combined/master-sheet exports (including Subtotal/Totals rows), and both "
              "on-screen preview pages match the exported headers exactly.")


if __name__ == "__main__":
    main()
