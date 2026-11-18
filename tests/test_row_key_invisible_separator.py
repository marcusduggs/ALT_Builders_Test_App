"""
Proves core.excel_reader.assign_row_keys()'s disambiguating suffix
(U+2063 INVISIBLE SEPARATOR + a 1-based occurrence number, appended to
EVERY occurrence of a duplicated Index -- including the first; see
that function's docstring for why no occurrence keeps the bare Index)
never reaches anything a user actually looks at: the in-app All Data/
Sum Data model, a single-increment export, a combined export, and the
Sum Data filter dropdown's values.

Uses TIO_H240649-38-00 HMR 10A.xlsm -- the one real fixture confirmed
to have duplicate grid Index values (39 of them), so this is a real
exercise of the suffix, not a synthetic stand-in.

Separately confirms the ONE place the suffix is deliberately allowed to
exist: status.json's own internal top-level keys, which are never
rendered to a user and are not what this test calls "reaching
user-visible output." That file (if a human ever opened it directly in
a text/hex editor) WOULD contain the raw separator bytes for every
occurrence of a duplicated row -- this test says so explicitly rather
than claiming the separator never appears in any file at all.

Run directly: `python tests/test_row_key_invisible_separator.py`
"""

import json
import os
import sys
import tempfile
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import openpyxl
from PySide6.QtWidgets import QApplication

from core import excel_export, excel_reader
from core.project_store import ProjectStore
from ui.mock_data import MockDataStore, build_combined_view
from ui.pages.combined_data_view_page import CombinedDataViewPage
from ui.pages.data_view_page import DataViewPage

FIXDIR = os.path.join(os.path.dirname(__file__), "fixtures", "hmr_smr_projects")
HMR_10A = os.path.join(FIXDIR, "TIO_H240649-38-00 HMR 10A.xlsm")
SEPARATOR = "⁣"


def _assert_no_separator(failures, label, values):
    bad = [v for v in values if v is not None and SEPARATOR in str(v)]
    if bad:
        failures.append(f"{label}: found the row-key separator in user-visible text: {bad[:5]}")


def main():
    failures = []

    if not os.path.isdir(FIXDIR):
        print(f"SKIP: {FIXDIR!r} not present (real HMR/SMR fixtures not available in this checkout)")
        return

    app = QApplication.instance() or QApplication(sys.argv)

    with tempfile.TemporaryDirectory() as tmp:
        store = MockDataStore(ProjectStore(base_dir=Path(tmp) / "data"))
        store.add_project("HMR10A Separator Check")
        inc = store.add_new_increment("HMR10A Separator Check", HMR_10A)
        display = store.get_increment_for_display("HMR10A Separator Check", inc.name)

        # Sanity: this file really does exercise the suffix -- if it
        # didn't, every check below would trivially pass for the wrong
        # reason (nothing to leak).
        row_keys_used = {r["row_key"] for r in display.all_data}
        suffixed = [k for k in row_keys_used if SEPARATOR in k]
        print(f"Distinct row keys: {len(row_keys_used)}, suffixed (duplicate-Index) ones: {len(suffixed)}")
        if not suffixed:
            failures.append("fixture assumption broken -- expected at least one suffixed row key on HMR 10A, found none")

        # ------------------------------------------------------------
        # 1 -- in-app model: All Data / Sum Data "index" fields
        # ------------------------------------------------------------
        _assert_no_separator(failures, "all_data['index']", [r["index"] for r in display.all_data])
        _assert_no_separator(failures, "sum_data['index']", [r["index"] for r in display.sum_data])
        expanded = excel_reader.expand_sum_data_rows(display.sum_data)
        _assert_no_separator(failures, "expand_sum_data_rows()['index']", [e["index"] for e in expanded])

        # ------------------------------------------------------------
        # 2 -- the real DataViewPage model + its Sum Data filter
        # dropdown's actual option text (not a re-derivation -- the
        # exact same QComboBox the user would open)
        # ------------------------------------------------------------
        page = DataViewPage("HMR10A Separator Check", display, store, on_back=lambda: None)
        model = page.table.model()
        all_data_index_values = [model.data(model.index(r, 0)) for r in range(len(display.all_data))]
        _assert_no_separator(failures, "DataViewPage All Data table column 0", all_data_index_values)

        # The filter bar's own combo dict is keyed by column index;
        # SINGLE_INCREMENT_SUM_DATA_FILTER_COLUMNS pairs each column
        # with its label -- find "Index"'s column directly from that.
        from ui.pages.data_view_page import SINGLE_INCREMENT_SUM_DATA_FILTER_COLUMNS
        index_col = next(col for col, label in SINGLE_INCREMENT_SUM_DATA_FILTER_COLUMNS if label == "Index")
        index_combo = page.sum_data_filter_bar._column_combos[index_col]
        dropdown_values = [index_combo.itemText(i) for i in range(index_combo.count())]
        print(f"Sum Data 'Index' filter dropdown has {len(dropdown_values)} options")
        _assert_no_separator(failures, "Sum Data Index filter dropdown", dropdown_values)

        # ------------------------------------------------------------
        # 3 -- single-increment export (.xlsx) -- read the REAL written
        # cells back with openpyxl, not the in-memory row dicts again
        # ------------------------------------------------------------
        single_path = Path(tmp) / "hmr10a_single.xlsx"
        excel_export.export_increment(display, single_path)
        wb = openpyxl.load_workbook(single_path)

        # All Data/Report sheets are currently gated off
        # (excel_export.SHOW_ALL_DATA_AND_REPORT_TABS=False) -- only
        # check the All Data sheet if it's actually there, so this test
        # doesn't silently assume a product decision that's already
        # been made elsewhere.
        if "All Data" in wb.sheetnames:
            all_data_ws = wb["All Data"]
            all_data_index_cells = [all_data_ws.cell(row=r, column=1).value for r in range(2, all_data_ws.max_row + 1)]
            _assert_no_separator(failures, "exported All Data sheet, column A", all_data_index_cells)
        else:
            print("  (All Data sheet not present in export -- SHOW_ALL_DATA_AND_REPORT_TABS is False, skipping that check)")

        sum_data_ws = wb["Sum Data"]
        sum_data_index_cells = [sum_data_ws.cell(row=r, column=3).value for r in range(2, sum_data_ws.max_row + 1)]
        _assert_no_separator(failures, "exported Sum Data sheet, column C", sum_data_index_cells)

        # ------------------------------------------------------------
        # 4 -- combined export + combined on-screen view (same fixture
        # combined with itself under a second increment name is
        # unnecessary -- one increment alone already has internal
        # duplicates to exercise)
        # ------------------------------------------------------------
        combined_view = build_combined_view([display])
        CombinedDataViewPage("HMR10A Separator Check", combined_view, store, on_back=lambda: None)
        app.processEvents()

        combined_path = Path(tmp) / "hmr10a_combined.xlsx"
        from core import combined_export
        combined_export.export_combined_report(combined_view.increments, combined_path)
        wb_combined = openpyxl.load_workbook(combined_path)

        if "All Data" in wb_combined.sheetnames:
            combined_all_data_ws = wb_combined["All Data"]
            combined_all_data_index_cells = [
                combined_all_data_ws.cell(row=r, column=2).value for r in range(2, combined_all_data_ws.max_row + 1)
            ]
            _assert_no_separator(failures, "combined exported All Data sheet, column B", combined_all_data_index_cells)
        else:
            print("  (combined All Data sheet not present in export -- SHOW_ALL_DATA_AND_REPORT_TABS is False, skipping)")

        combined_sum_data_ws = wb_combined["Sum Data"]
        combined_sum_data_index_cells = [
            combined_sum_data_ws.cell(row=r, column=3).value for r in range(2, combined_sum_data_ws.max_row + 1)
        ]
        _assert_no_separator(failures, "combined exported Sum Data sheet, column C", combined_sum_data_index_cells)

        # ------------------------------------------------------------
        # 5 -- status.json round-trip: the separator is EXPECTED to
        # appear here (it's the internal storage key, not a display
        # surface) -- confirmed explicitly, both directions.
        # ------------------------------------------------------------
        increment_record = store.store.get_increment("HMR10A Separator Check", inc.name)
        status_path = store.store._status_json_path(
            store.store.get_project("HMR10A Separator Check").slug, increment_record.slug
        )
        raw_text = status_path.read_text(encoding="utf-8")
        raw_keys = list(json.loads(raw_text).keys())
        suffixed_on_disk = [k for k in raw_keys if SEPARATOR in k]
        print(f"status.json: {len(raw_keys)} top-level keys, {len(suffixed_on_disk)} carry the separator")
        if not suffixed_on_disk:
            failures.append(
                "status.json should contain at least one row-key-suffixed entry for HMR 10A's duplicate "
                "indexes -- found none, meaning either the fixture assumption changed or row keys aren't "
                "actually reaching storage"
            )
        # Round-trip: load_status() must hand back the EXACT same key
        # set (as Python strings, separator intact) that's on disk.
        loaded = store.store.load_status("HMR10A Separator Check", inc.name)
        if set(loaded.keys()) != set(raw_keys):
            failures.append(
                f"load_status() keys don't match the raw file's keys -- got {set(loaded.keys()) ^ set(raw_keys)} "
                f"as the symmetric difference"
            )

    print("\n" + "=" * 70)
    if failures:
        print("RESULT: FAIL")
        for f in failures:
            print(" -", f)
        raise SystemExit(1)
    else:
        print(
            "RESULT: PASS -- the row-key disambiguating separator never appears in the in-app All Data/Sum "
            "Data model, the Sum Data Index filter dropdown, or either exported .xlsx (single-increment or "
            "combined); it DOES appear, as designed, in status.json's own internal keys for HMR 10A's "
            "duplicate-Index rows, and load_status()/save_status() round-trip those keys exactly."
        )


if __name__ == "__main__":
    main()
