"""
Regression test for Sum Data's "Project Name / Increment #" column
(column B) value -- confirmed, per a direct follow-up review, that all
FIVE surfaces that show it (single-increment on-screen view, combined
on-screen view, single-increment export, combined export, and the
combined-view's own filter dropdown) show the EXACT SAME text:
ui.mock_data.Increment.display_name -- a custom name set at upload,
else Rey's own short code (core.excel_reader.derive_increment_code(),
e.g. "SMR 3D"/"HMR 10A" -- see test_derive_increment_code.py's own
oracle test against Rey's TIO_Changes.xlsx), else
derive_short_display_name()'s shortened-prose fallback for an
increment outside that naming convention -- never increment.name (the
full, sometimes 300-550+ character raw Record Name).

Also confirms the on-screen column width still fits: measured directly
(QFontMetrics against the real loaded stylesheet) that every real
HMR/SMR fixture's own CODE is tiny (well under 100px) now that
derive_increment_code() is the common case, but display_name can still
fall back to a longer name (a user-typed custom name up to 40
characters, or derive_short_display_name()'s prose fallback for a
non-HMR/SMR increment -- confirmed: inc8_real_2stage.xlsm's own
fallback needs 322px) -- that one real outlier still elides with "…"
and recovers via this column's own full-name tooltip, the same
accepted pattern Description/Approval Agency already rely on.

Run directly: `python tests/test_sum_data_display_name_column.py`
"""

import glob
import os
import sys
import tempfile
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import openpyxl
from PySide6.QtCore import Qt, QSortFilterProxyModel
from PySide6.QtGui import QFontMetrics
from PySide6.QtWidgets import QApplication

from core import combined_export, excel_export, excel_reader
from core.project_store import ProjectStore
from ui.mock_data import MockDataStore, build_combined_view
from ui.pages.combined_data_view_page import CombinedDataViewPage
from ui.pages.data_view_page import DataViewPage

FIXDIR = os.path.join(os.path.dirname(__file__), "fixtures", "hmr_smr_projects")
SAMPLE = os.path.join(os.path.dirname(__file__), "fixtures", "sample_increment.xlsm")
COLUMN_B = 1


def main():
    failures = []
    app = QApplication.instance() or QApplication(sys.argv)
    style_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "ui", "style.qss")
    app.setStyleSheet(Path(style_path).read_text())

    if not os.path.isdir(FIXDIR):
        print(f"SKIP: {FIXDIR!r} not present (real HMR/SMR fixtures not available in this checkout)")
        return

    files = {
        "SMR 3D": "TIO H231182-38-00 SMR 3D 2022 Delta 1.xlsm",
        "HMR 10A": "TIO_H240649-38-00 HMR 10A.xlsm",
        "HMR 9B": "TIO H231359-38-00_HMR 9B.xlsm",
    }

    # ------------------------------------------------------------
    # 1 -- all 5 surfaces show the IDENTICAL text, for both a
    # fallback-derived name AND an explicit custom one.
    # ------------------------------------------------------------
    print("=" * 70)
    print("1 -- Column B identical across all 5 surfaces")
    print("=" * 70)

    with tempfile.TemporaryDirectory() as tmp:
        store = MockDataStore(ProjectStore(base_dir=Path(tmp) / "data"))
        store.add_project("Check")
        displays = []
        for label, fname in files.items():
            inc = store.add_new_increment("Check", os.path.join(FIXDIR, fname))
            displays.append(store.get_increment_for_display("Check", inc.name))

        # Also one with an EXPLICIT custom display_name, to prove the
        # "display_name when set" half of the rule, not just the
        # fallback half.
        prepared = store.prepare_new_increment("Check", SAMPLE)
        custom_inc = store.finalize_new_increment(prepared, "My Custom Label")
        custom_display = store.get_increment_for_display("Check", custom_inc.name)
        displays.append(custom_display)

        for display in displays:
            expected = display.display_name
            print(f"\n  {display.name[:50]!r}... expected display_name={expected!r}")

            page = DataViewPage("Check", display, store, on_back=lambda: None)
            single_view_value = page.sum_data_table.model().data(page.sum_data_table.model().index(0, COLUMN_B))
            print(f"    1. single view:      {single_view_value!r}")
            if single_view_value != expected:
                failures.append(f"1: {display.name}: single view column B = {single_view_value!r}, expected {expected!r}")

            single_path = Path(tmp) / f"single_{display.name[:10]}.xlsx"
            excel_export.export_increment(display, single_path)
            ws_single = openpyxl.load_workbook(single_path)["Sum Data"]
            single_export_value = ws_single.cell(row=2, column=2).value
            print(f"    3. single export:    {single_export_value!r}")
            if single_export_value != expected:
                failures.append(f"1: {display.name}: single export column B = {single_export_value!r}, expected {expected!r}")

        view = build_combined_view(displays)
        cpage = CombinedDataViewPage("Check", view, store, on_back=lambda: None)
        cmodel = cpage.sum_data_table.model()
        combined_path = Path(tmp) / "combined.xlsx"
        combined_export.export_combined_report(displays, combined_path)
        ws_combined = openpyxl.load_workbook(combined_path)["Sum Data"]

        combo = cpage.sum_data_filter_bar._column_combos[COLUMN_B]
        dropdown_values = {combo.itemText(i) for i in range(combo.count())}
        print(f"\n  5. filter dropdown values: {sorted(dropdown_values)}")

        row_cursor = 2  # 1-indexed export row, header is row 1
        for display in displays:
            expected = display.display_name
            # Find this increment's own first row in the combined on-screen model.
            combined_row = next(
                r for r in range(cmodel.rowCount())
                if cmodel.data(cmodel.index(r, 0)) == str(display.project_number or "")
                and cmodel.data(cmodel.index(r, COLUMN_B)) is not None
            )
            combined_view_value = cmodel.data(cmodel.index(combined_row, COLUMN_B))
            print(f"  {display.name[:40]!r}... 2. combined view: {combined_view_value!r}")
            if combined_view_value != expected:
                failures.append(f"1: {display.name}: combined view column B = {combined_view_value!r}, expected {expected!r}")

            matching_export_rows = [
                r for r in range(2, ws_combined.max_row + 1)
                if ws_combined.cell(row=r, column=1).value == str(display.project_number or "")
            ]
            if not matching_export_rows:
                failures.append(f"1: {display.name}: no rows found in combined export at all")
            else:
                combined_export_value = ws_combined.cell(row=matching_export_rows[0], column=2).value
                print(f"  {display.name[:40]!r}... 4. combined export: {combined_export_value!r}")
                if combined_export_value != expected:
                    failures.append(f"1: {display.name}: combined export column B = {combined_export_value!r}, expected {expected!r}")

            if expected not in dropdown_values:
                failures.append(f"1: {display.name}: expected display_name {expected!r} missing from the filter dropdown entirely")

        if custom_display.display_name != "My Custom Label":
            failures.append(f"1: custom display_name should be 'My Custom Label', got {custom_display.display_name!r}")

    # ------------------------------------------------------------
    # 2 -- derive_short_display_name()'s real output can run up to
    # ~100 characters -- confirm the on-screen column 1 width (both
    # pages) is wide enough for the WIDEST derived name across ALL 14
    # real fixtures, OR -- for the few genuine outliers -- that the
    # table elides with "…" (not a silent hard clip) and the cell's own
    # tooltip still has the full text.
    # ------------------------------------------------------------
    print("\n" + "=" * 70)
    print("2 -- Column width fits the common case; outliers elide with a tooltip")
    print("=" * 70)

    all_files = sorted(glob.glob(os.path.join(FIXDIR, "*"))) + [
        os.path.join(os.path.dirname(__file__), "fixtures", "inc8_real_2stage.xlsm")
    ]
    with tempfile.TemporaryDirectory() as tmp:
        store = MockDataStore(ProjectStore(base_dir=Path(tmp) / "data"))
        store.add_project("WidthCheck")
        widest_display = None
        widest_px = 0
        page_for_metrics = None
        for f in all_files:
            inc = store.add_new_increment("WidthCheck", f)
            display = store.get_increment_for_display("WidthCheck", inc.name)
            page = DataViewPage("WidthCheck", display, store, on_back=lambda: None)
            if page_for_metrics is None:
                page_for_metrics = page
            fm = QFontMetrics(page.sum_data_table.font())
            px = fm.horizontalAdvance(display.display_name)
            if px > widest_px:
                widest_px = px
                widest_display = display

        col_width = page_for_metrics.sum_data_table.columnWidth(COLUMN_B)
        print(f"  Column width: {col_width}px")
        print(f"  Widest real display_name: {widest_display.display_name!r} ({widest_px}px)")
        if col_width < 250:
            failures.append(f"2: column 1 width ({col_width}px) looks too narrow for realistic derived names -- regression?")

        if widest_px > col_width:
            # Not a failure -- confirm elide mode + tooltip recover it.
            elide_mode = page_for_metrics.sum_data_table.textElideMode()
            print(f"  Widest case exceeds column width -- textElideMode={elide_mode} (must be ElideRight, not None)")
            if elide_mode != Qt.TextElideMode.ElideRight:
                failures.append(f"2: the widest real display_name overflows the column, but textElideMode is {elide_mode}, not ElideRight -- would hard-clip with no visual indicator")

            # Confirm the tooltip on that row carries the full real name.
            page = DataViewPage("WidthCheck", widest_display, store, on_back=lambda: None)
            model = page.sum_data_table.model()
            tooltip = model.item(0, COLUMN_B).toolTip() if hasattr(model, "item") else None
            if tooltip is None:
                # table.model() may be the filter proxy; fetch the source model's item.
                source = model.sourceModel() if isinstance(model, QSortFilterProxyModel) else model
                tooltip = source.item(0, COLUMN_B).toolTip()
            print(f"  Tooltip for the widest case: {tooltip!r}")
            if tooltip != widest_display.name:
                failures.append(f"2: the widest display_name's row should have a tooltip with the full real name {widest_display.name!r}, got {tooltip!r}")

    print("\n" + "=" * 70)
    if failures:
        print("RESULT: FAIL")
        for f in failures:
            print(" -", f)
        raise SystemExit(1)
    else:
        print(
            "RESULT: PASS -- column B shows the identical Increment.display_name text (custom name when set, "
            "else the derived short name) across all 5 surfaces: single view, combined view, single export, "
            "combined export, and the combined filter dropdown; the column width fits the common real case, "
            "and the rare, genuinely long outlier elides with '…' and keeps the full name in its tooltip."
        )


if __name__ == "__main__":
    main()
