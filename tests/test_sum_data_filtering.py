"""
Regression test for Sum Data's column filtering, in the style of the
client's own reference sheet -- in-app filter bar (ui.widgets.
sum_data_filter_bar, wired into DataViewPage and CombinedDataViewPage)
and real Excel AutoFilter on both exports.

Sum Data has NO Subtotal/Totals rows and NO Open/Done/Total/% Complete
columns at all anymore (Rey's TIO_Changes.xlsx "Sum Data" tab change
requests #2/#3) -- the summary-row recalculation this test used to
cover (ui.widgets.sum_data_filter_bar.SummaryAwareFilterProxyModel,
mark_summary_row(), SUMMARY_ROW_ROLE/SUMMARY_IS_TOTALS_ROLE) was
deleted along with those rows/columns; there is nothing left to
recalculate. This file now covers what's left:

  1. In-app filtering: a column filter hides non-matching rows, never
     mutates the underlying Increment.sum_data / status.json.
  2. Filters reset to "All" on every tab rebuild (increment switch,
     version switch, or just switching back to the Sum Data tab --
     DataViewPage already rebuilds it on all three).
  3. Preview/export parity: the in-app filter bar's column set matches
     the exported header columns at the same index -- now a literal,
     unconditional match (no more "Increment" on screen / "Inc #" in
     the export split -- both are "Project Name / Increment #" now).
  4. Export AutoFilter is REAL -- ws.auto_filter.ref read back with
     openpyxl, not just "the arrows appear" -- covering the header
     through the LAST row (no Totals/Subtotal row to exclude anymore,
     single-increment or combined).
  5. Notes/Completion Date stay blank and within the filterable range.
  6. QFontMetrics check against the REAL loaded stylesheet for the
     filter-bar labels, so this feature doesn't reintroduce the
     truncated-header bug class (tests/test_sum_data_column_widths.py's
     same methodology, applied to the new UI).

Run directly: `python tests/test_sum_data_filtering.py`
"""

import os
import sys
import tempfile
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import openpyxl
from PySide6.QtCore import QModelIndex
from PySide6.QtGui import QFontMetrics
from PySide6.QtWidgets import QApplication, QLabel

from core import combined_export, excel_export
from core.excel_export import SUM_DATA_HEADERS
from core.project_store import ProjectStore
from ui.mock_data import MockDataStore, build_combined_view
from ui.pages.combined_data_view_page import CombinedDataViewPage
from ui.pages.data_view_page import SINGLE_INCREMENT_SUM_DATA_FILTER_COLUMNS, SUM_DATA_FILTER_COLUMNS, DataViewPage

FIXDIR = os.path.join(os.path.dirname(__file__), "fixtures", "hmr_smr_projects")
PROJECT_NAME = "Sum Data Filtering Regression"
# SMR 2D has a genuine Done/Open MIX (confirmed during the HCAI/FDD-Use
# fix's own verification), so filtering by Status here actually hides
# some rows while keeping others, rather than vacuously matching
# everything.
FILE_A = "TIO H220369-38-00_SMR 2D.xlsm"
FILE_B = "TIO H231182-38-00 SMR 3D 2022 Delta 1.xlsm"
FILE_C = "TIO S251652-38-00 SMR 4C.xlsm"


def _visible_rows(proxy) -> list[int]:
    source = proxy.sourceModel()
    return [r for r in range(source.rowCount()) if proxy.filterAcceptsRow(r, QModelIndex())]


def main():
    if not os.path.isdir(FIXDIR):
        print(f"SKIP: {FIXDIR!r} not present (real HMR/SMR fixtures not available in this checkout)")
        return

    failures = []
    app = QApplication.instance() or QApplication(sys.argv)
    style_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "ui", "style.qss")
    app.setStyleSheet(Path(style_path).read_text())

    with tempfile.TemporaryDirectory() as tmp:
        data_dir = Path(tmp) / "data"
        store = MockDataStore(ProjectStore(base_dir=data_dir))
        store.add_project(PROJECT_NAME)
        inc_a = store.add_new_increment(PROJECT_NAME, os.path.join(FIXDIR, FILE_A))
        inc_b = store.add_new_increment(PROJECT_NAME, os.path.join(FIXDIR, FILE_B))
        inc_c = store.add_new_increment(PROJECT_NAME, os.path.join(FIXDIR, FILE_C))
        display_a = store.get_increment_for_display(PROJECT_NAME, inc_a.name)
        display_b = store.get_increment_for_display(PROJECT_NAME, inc_b.name)
        display_c = store.get_increment_for_display(PROJECT_NAME, inc_c.name)

        # ------------------------------------------------------------
        # 1 -- in-app filtering: hides non-matching rows, never mutates
        # the underlying data
        # ------------------------------------------------------------
        print("=" * 70)
        print("1 -- In-app filtering hides rows without touching underlying data")
        print("=" * 70)

        page = DataViewPage(PROJECT_NAME, display_a, store, on_back=lambda: None)
        app.processEvents()
        proxy = page._sum_data_filter_proxy
        source_model = proxy.sourceModel()
        total_source_rows = source_model.rowCount()

        status_before = store.store.load_status(PROJECT_NAME, inc_a.name)

        # Pick a real Status value that appears on at least one real row.
        status_combo = page.sum_data_filter_bar._column_combos[6]
        real_status_values = [status_combo.itemText(i) for i in range(status_combo.count()) if status_combo.itemText(i) != "All"]
        if not real_status_values:
            failures.append("1: expected at least one real Status value in the filter dropdown")
        else:
            target_status = real_status_values[0]
            status_combo.setCurrentText(target_status)
            app.processEvents()

            visible_rows = _visible_rows(proxy)
            print(f"  Filtering Status={target_status!r}: {len(visible_rows)} of {total_source_rows} source rows pass")
            if len(visible_rows) >= total_source_rows:
                failures.append(
                    f"1: fixture assumption broken -- expected Status={target_status!r} to hide at least one "
                    f"row (this fixture has a genuine Done/Open mix), but all {total_source_rows} rows passed"
                )
            # Every visible row must actually have that Status value.
            for r in visible_rows:
                cell_status = source_model.data(source_model.index(r, 6))
                if cell_status != target_status:
                    failures.append(f"1: row {r} passed the Status={target_status!r} filter but its own Status is {cell_status!r}")

            # Filtering must not have touched status.json at all.
            status_combo.setCurrentText("All")
            app.processEvents()
            status_after = store.store.load_status(PROJECT_NAME, inc_a.name)
            if status_after != status_before:
                failures.append("1: filtering changed status.json on disk -- it must only change which rows are SHOWN")

        # ------------------------------------------------------------
        # 2 -- same check against CombinedDataViewPage (3 real fixtures)
        # plus the "Project #" filter, plus a filter matching nothing.
        # ------------------------------------------------------------
        print("\n" + "=" * 70)
        print("2 -- CombinedDataViewPage filtering (3 real fixtures)")
        print("=" * 70)

        view = build_combined_view([display_a, display_b, display_c])
        cpage = CombinedDataViewPage(PROJECT_NAME, view, store, on_back=lambda: None)
        app.processEvents()
        c_proxy = cpage._sum_data_filter_proxy
        c_model = c_proxy.sourceModel()
        c_total_rows = c_model.rowCount()

        c_status_combo = cpage.sum_data_filter_bar._column_combos[6]
        c_status_combo.setCurrentText("Done")
        app.processEvents()
        visible_done_rows = _visible_rows(c_proxy)
        print(f"  Filtering Status=Done: {len(visible_done_rows)} of {c_total_rows} rows pass")
        if not visible_done_rows or len(visible_done_rows) >= c_total_rows:
            failures.append(
                f"2: fixture assumption broken -- expected Status=Done to hide at least one row but match "
                f"at least one, got {len(visible_done_rows)} of {c_total_rows}"
            )
        for r in visible_done_rows:
            if c_model.data(c_model.index(r, 6)) != "Done":
                failures.append(f"2: row {r} passed the Status=Done filter but its own Status is not Done")
        c_status_combo.setCurrentText("All")
        app.processEvents()

        project_combo = cpage.sum_data_filter_bar._column_combos[0]
        project_values = [project_combo.itemText(i) for i in range(project_combo.count()) if project_combo.itemText(i) != "All"]
        target_project = project_values[0]
        project_combo.setCurrentText(target_project)
        app.processEvents()
        visible_project_rows = _visible_rows(c_proxy)
        print(f"  Filtering Project #={target_project!r}: {len(visible_project_rows)} of {c_total_rows} rows pass")
        for r in visible_project_rows:
            if c_model.data(c_model.index(r, 0)) != target_project:
                failures.append(f"2: row {r} passed the Project #={target_project!r} filter but its own Project # doesn't match")
        project_combo.setCurrentText("All")
        app.processEvents()

        index_combo = cpage.sum_data_filter_bar._column_combos[2]
        c_proxy.set_filters({2: "NO_SUCH_INDEX_VALUE_AT_ALL"})
        app.processEvents()
        visible_none = _visible_rows(c_proxy)
        print(f"  A filter matching nothing: {len(visible_none)} rows visible (expected 0)")
        if visible_none:
            failures.append(f"2: a filter matching nothing should show zero rows, got {len(visible_none)}")
        c_proxy.set_filters({})
        app.processEvents()

        # Filtering never touches status.json for any increment.
        for inc, name in [(inc_a, "SMR 2D"), (inc_b, "SMR 3D Delta 1"), (inc_c, "SMR 4C")]:
            before = store.store.load_status(PROJECT_NAME, inc.name)
            c_status_combo.setCurrentText("Done")
            app.processEvents()
            c_status_combo.setCurrentText("All")
            app.processEvents()
            after = store.store.load_status(PROJECT_NAME, inc.name)
            if before != after:
                failures.append(f"2: filtering changed status.json on disk for {name} -- it must only change which rows are SHOWN")

        # ------------------------------------------------------------
        # 3 -- filters reset to "All" on every Sum Data tab rebuild
        # ------------------------------------------------------------
        print("\n" + "=" * 70)
        print("3 -- Filters reset to 'All' on every Sum Data tab rebuild")
        print("=" * 70)

        if real_status_values:
            page.sum_data_filter_bar._column_combos[6].setCurrentText(real_status_values[0])
            app.processEvents()
            had_active_filter = page.sum_data_filter_bar.has_active_filters()
            page._refresh_sum_data_tab()  # same call _on_tab_changed makes on every switch back to this tab
            app.processEvents()
            reset_filter_active = page.sum_data_filter_bar.has_active_filters()
            print(f"  Had active filter before rebuild: {had_active_filter}; after rebuild: {reset_filter_active}")
            if not had_active_filter:
                failures.append("3: fixture assumption broken -- expected a filter to be active before the rebuild")
            if reset_filter_active:
                failures.append("3: rebuilding the Sum Data tab should reset filters to 'All', but one is still active")

        # ------------------------------------------------------------
        # 4 -- preview/export parity: filter bar columns match the
        # exported header columns at the same index, literally now --
        # no more "Increment" on screen / "Inc #" in the export split
        # (both are "Project Name / Increment #").
        # ------------------------------------------------------------
        print("\n" + "=" * 70)
        print("4 -- Filter bar columns match the exported header columns")
        print("=" * 70)

        def _check_parity(label_prefix, actual_columns):
            mismatches = [
                (column, combo_label, SUM_DATA_HEADERS[column])
                for column, combo_label in actual_columns
                if combo_label != SUM_DATA_HEADERS[column]
            ]
            if mismatches:
                failures.append(f"4: {label_prefix} filter label(s) don't match the exported header at the same column index: {mismatches}")

        single_columns_by_index = dict(SUM_DATA_FILTER_COLUMNS)
        single_filter_columns = [
            (col, single_columns_by_index[col]) for col in page.sum_data_filter_bar._column_combos
        ]
        print(f"  DataViewPage filter columns (actual): {single_filter_columns}")
        expected_single = set(SINGLE_INCREMENT_SUM_DATA_FILTER_COLUMNS)
        if set(single_filter_columns) != expected_single:
            failures.append(
                f"4: DataViewPage's actual filter columns {single_filter_columns} should equal "
                f"SINGLE_INCREMENT_SUM_DATA_FILTER_COLUMNS {sorted(expected_single)}"
            )
        if not set(single_filter_columns) <= set(SUM_DATA_FILTER_COLUMNS):
            failures.append("4: DataViewPage's filter columns should be a SUBSET of the full SUM_DATA_FILTER_COLUMNS")
        _check_parity("DataViewPage", single_filter_columns)

        combined_filter_columns = [
            (col, single_columns_by_index[col]) for col in cpage.sum_data_filter_bar._column_combos
        ]
        print(f"  CombinedDataViewPage filter columns (actual): {combined_filter_columns}")
        if set(combined_filter_columns) != set(SUM_DATA_FILTER_COLUMNS):
            failures.append(
                f"4: CombinedDataViewPage's actual filter columns {combined_filter_columns} should be all six "
                f"of SUM_DATA_FILTER_COLUMNS {SUM_DATA_FILTER_COLUMNS}"
            )
        _check_parity("CombinedDataViewPage", combined_filter_columns)

        # ------------------------------------------------------------
        # 5 -- Export AutoFilter is a REAL AutoFilter, covering the
        # header through the LAST row -- no Totals/Subtotal row to
        # exclude anymore, single-increment or combined.
        # ------------------------------------------------------------
        print("\n" + "=" * 70)
        print("5 -- Export AutoFilter range, read back with openpyxl")
        print("=" * 70)

        single_path = str(Path(tmp) / "single_export.xlsx")
        excel_export.export_increment(display_a, single_path)
        wb_single = openpyxl.load_workbook(single_path)
        ws_single = wb_single["Sum Data"]
        print(f"  Single-increment: max_row={ws_single.max_row}, auto_filter.ref={ws_single.auto_filter.ref!r}")
        expected_single_ref = f"A1:I{ws_single.max_row}"
        if ws_single.auto_filter.ref != expected_single_ref:
            failures.append(
                f"5: single-increment AutoFilter ref should be {expected_single_ref!r} (every row, no Totals "
                f"row to exclude), got {ws_single.auto_filter.ref!r}"
            )
        last_row_index_val = ws_single.cell(row=ws_single.max_row, column=3).value
        if last_row_index_val in (None, "", "Totals"):
            failures.append(f"5: fixture assumption broken -- expected the last row to be a real data row, got Index={last_row_index_val!r}")

        combined_path = str(Path(tmp) / "combined_export.xlsx")
        combined_export.export_combined_report([display_a, display_b, display_c], combined_path)
        wb_combined = openpyxl.load_workbook(combined_path)
        ws_combined = wb_combined["Sum Data"]
        print(f"  Combined (3 fixtures): max_row={ws_combined.max_row}, auto_filter.ref={ws_combined.auto_filter.ref!r}")
        expected_combined_ref = f"A1:I{ws_combined.max_row}"
        if ws_combined.auto_filter.ref != expected_combined_ref:
            failures.append(
                f"5: combined AutoFilter ref should be {expected_combined_ref!r} (every row -- no Subtotal/"
                f"Totals rows exist to work around anymore), got {ws_combined.auto_filter.ref!r}"
            )
        combined_increment_col = [ws_combined.cell(row=r, column=2).value for r in range(2, ws_combined.max_row + 1)]
        if any("Subtotal" in str(v) or v == "Totals" for v in combined_increment_col):
            failures.append("5: combined Sum Data export should have NO Subtotal/Totals rows at all")

        # ------------------------------------------------------------
        # 6 -- Notes/Completion Date stay blank and within the
        # filterable range
        # ------------------------------------------------------------
        print("\n" + "=" * 70)
        print("6 -- Notes/Completion Date stay blank, inside the AutoFilter range")
        print("=" * 70)

        notes_col, completion_date_col = 8, 9
        blank_violations = [
            r for r in range(2, ws_single.max_row + 1)
            if ws_single.cell(row=r, column=notes_col).value not in (None, "")
            or ws_single.cell(row=r, column=completion_date_col).value not in (None, "")
        ]
        if blank_violations:
            failures.append(f"6: Notes/Completion Date should be blank on every data row, found values on rows {blank_violations}")

        # ------------------------------------------------------------
        # 7 -- QFontMetrics check: filter bar labels aren't truncated
        # against the REAL loaded stylesheet (same bug class as the
        # Sum Data table's own header-width regressions)
        # ------------------------------------------------------------
        print("\n" + "=" * 70)
        print("7 -- Filter bar labels render at their full sizeHint width (no truncation)")
        print("=" * 70)

        for page_name, bar, expected_columns in [
            ("DataViewPage", page.sum_data_filter_bar, SINGLE_INCREMENT_SUM_DATA_FILTER_COLUMNS),
            ("CombinedDataViewPage", cpage.sum_data_filter_bar, SUM_DATA_FILTER_COLUMNS),
        ]:
            if set(bar._column_combos.keys()) != {c for c, _ in expected_columns}:
                failures.append(
                    f"7: {page_name}'s filter bar columns {sorted(bar._column_combos.keys())} don't match "
                    f"the expected set {sorted(c for c, _ in expected_columns)}"
                )
            caption_labels = {lbl.text(): lbl for lbl in bar.findChildren(QLabel) if lbl.text().endswith(":")}
            for column, label in expected_columns:
                caption_text = f"{label}:"
                caption = caption_labels.get(caption_text)
                if caption is None:
                    failures.append(f"7: {page_name} is missing its {caption_text!r} filter caption label entirely")
                    continue
                fm = QFontMetrics(caption.font())
                required_width = fm.horizontalAdvance(caption_text)
                actual_width = caption.sizeHint().width()
                status = "pass" if actual_width >= required_width else "FAIL"
                print(f"  {page_name} {caption_text!r}: required={required_width} actual_sizeHint={actual_width} [{status}]")
                if actual_width < required_width:
                    failures.append(
                        f"7: {page_name}'s {caption_text!r} caption sizeHint ({actual_width}) is narrower than "
                        f"its own text needs ({required_width}) -- would truncate"
                    )

    print("\n" + "=" * 70)
    if failures:
        print("RESULT: FAIL")
        for f in failures:
            print(" -", f)
        raise SystemExit(1)
    else:
        print(
            "RESULT: PASS -- Sum Data column filtering hides only matching rows without ever mutating "
            "underlying data, on both DataViewPage and CombinedDataViewPage; filters reset on tab rebuild; "
            "filter columns match the exported headers exactly (no more on-screen/export label split); "
            "export AutoFilter ranges cover every row (no Subtotal/Totals rows exist to exclude anymore); "
            "Notes/Completion Date stay blank; and no filter-bar label is truncated."
        )


if __name__ == "__main__":
    main()
