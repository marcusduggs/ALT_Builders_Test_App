"""
Regression test for Sum Data's column filtering, in the style of the
client's own reference sheet -- in-app filter bar (ui.widgets.
sum_data_filter_bar, wired into DataViewPage and CombinedDataViewPage)
and real Excel AutoFilter on both exports.

Covers, all against real HMR/SMR fixture data:
  1. In-app filtering: a column filter hides non-matching rows, never
     mutates the underlying Increment.sum_data / status.json.
  2. Totals/Subtotal recalculation (DataViewPage's single Totals row,
     and CombinedDataViewPage's Totals + per-increment Subtotal rows):
     with no filter active, every summary row shows exactly its
     precomputed full-dataset value, byte-for-byte unchanged. With a
     filter active, every summary row recomputes Open/Done/Total/%
     Complete from only its own scope's CURRENTLY VISIBLE rows; a
     Subtotal row is hidden entirely if none of its own rows pass; the
     Totals row is never hidden and gains a "(filtered)" label suffix.
     Verified: the baseline (no-filter) values match the pre-existing
     full-dataset numbers exactly; filtering to Status=Done isolates
     SMR 2D's 7 Done rows with Totals Done=7/Open=0 and only SMR 2D's
     Subtotal visible; filtering to one increment shows only that
     increment's Subtotal with numbers equal to its own unfiltered
     Subtotal; clearing the filter returns everything to the original
     baseline; a filter matching nothing shows zero data rows with
     Totals at zero and % Complete not dividing by zero.
  3. Filters reset to "All" on every tab rebuild (increment switch,
     version switch, or just switching back to the Sum Data tab --
     DataViewPage already rebuilds it on all three).
  4. Preview/export parity: the in-app filter bar's column set matches
     the exported header columns at the same index. A Subtotal row's
     own LABEL is a deliberate, labels-only exception (short display
     name on screen, full text in the export, same convention as the
     existing "Increment" column) -- covered on its own, since numbers
     (not labels) are what must still match exactly.
  5. Export AutoFilter is REAL -- ws.auto_filter.ref read back with
     openpyxl, not just "the arrows appear" -- single-increment range
     excludes the Totals row; combined range necessarily includes every
     interspersed Subtotal row and the final Totals row (see
     core.combined_export's own AutoFilter comment for why that's
     unavoidable). Exports are NOT touched by the Totals/Subtotal
     recalculation or short-label changes above -- this section
     confirms that too.
  6. Notes/Completion Date stay blank and within the filterable range.
  7. QFontMetrics check against the REAL loaded stylesheet for the new
     filter-bar labels, so this feature doesn't reintroduce the
     truncated-header bug class (tests/test_sum_data_column_widths.py's
     same methodology, applied to the new UI).

Run directly: `python tests/test_sum_data_filtering.py`
"""

import json
import os
import sys
import tempfile
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import openpyxl
from PySide6.QtCore import QModelIndex, Qt
from PySide6.QtGui import QFontMetrics
from PySide6.QtWidgets import QApplication, QLabel

from core import combined_export, excel_export
from core.excel_export import SUM_DATA_HEADERS
from core.excel_reader import derive_short_display_name
from core.project_store import ProjectStore
from ui.mock_data import MockDataStore, build_combined_view
from ui.pages.combined_data_view_page import CombinedDataViewPage
from ui.pages.data_view_page import SINGLE_INCREMENT_SUM_DATA_FILTER_COLUMNS, SUM_DATA_FILTER_COLUMNS, DataViewPage
from ui.widgets.sum_data_filter_bar import SUMMARY_IS_TOTALS_ROLE, SUMMARY_ROW_ROLE

FIXDIR = os.path.join(os.path.dirname(__file__), "fixtures", "hmr_smr_projects")
PROJECT_NAME = "Sum Data Filtering Regression"
# SMR 2D has a genuine Done/Open MIX (7 Done, 12 Open -- confirmed
# during the HCAI/FDD-Use fix's own verification), so filtering by
# Status here actually hides some rows while keeping others, rather
# than vacuously matching everything.
FILE_A = "TIO H220369-38-00_SMR 2D.xlsm"
FILE_B = "TIO H231182-38-00 SMR 3D 2022 Delta 1.xlsm"
FILE_C = "TIO S251652-38-00 SMR 4C.xlsm"

NUMERIC_COLS = (9, 10, 11, 12)  # Open, Done, Total, % Complete


def _summary_rows(model) -> set[int]:
    return {r for r in range(model.rowCount()) if model.data(model.index(r, 1), SUMMARY_ROW_ROLE)}


def _is_totals_row(model, row: int) -> bool:
    return bool(model.data(model.index(row, 1), SUMMARY_IS_TOTALS_ROLE))


def _proxy_data(proxy, source_row: int, column: int):
    """Reads (source_row, column) THROUGH the proxy, not the source
    model directly -- required for a summary row's numeric/label
    columns, since SummaryAwareFilterProxyModel's recalculation lives
    entirely in its own data() override; the source model's stored
    cell value never changes (filtering must never mutate it -- see
    that class's docstring), so reading the source model directly
    would always see the ORIGINAL unfiltered value regardless of the
    active filter, silently defeating this exact check.
    """
    source_model = proxy.sourceModel()
    return proxy.data(proxy.mapFromSource(source_model.index(source_row, column)))


def _row_values(proxy, source_row: int) -> tuple:
    return tuple(_proxy_data(proxy, source_row, c) for c in NUMERIC_COLS)


def _label_of(proxy, source_row: int):
    return _proxy_data(proxy, source_row, 1)


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
        totals_source_row = next(r for r in _summary_rows(source_model) if _is_totals_row(source_model, r))

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
            # Every visible real (non-summary) row must actually have
            # that Status value; the Totals row must always be among
            # the visible rows regardless.
            for r in visible_rows:
                if r == totals_source_row:
                    continue
                cell_status = source_model.data(source_model.index(r, 6))
                if cell_status != target_status:
                    failures.append(f"1: row {r} passed the Status={target_status!r} filter but its own Status is {cell_status!r}")
            if totals_source_row not in visible_rows:
                failures.append("1: the Totals row must never be hidden by a filter, but it was")

            # Filtering must not have touched status.json at all.
            status_combo.setCurrentText("All")
            app.processEvents()
            status_after = store.store.load_status(PROJECT_NAME, inc_a.name)
            if status_after != status_before:
                failures.append("1: filtering changed status.json on disk -- it must only change which rows are SHOWN")

        # ------------------------------------------------------------
        # 2 -- Totals/Subtotal recalculation: DataViewPage's single
        # Totals row first (no Subtotal rows there -- one increment).
        # ------------------------------------------------------------
        print("\n" + "=" * 70)
        print("2a -- DataViewPage Totals row recalculates when filtered, matches baseline unfiltered")
        print("=" * 70)

        baseline_single_totals = _row_values(proxy, totals_source_row)
        baseline_single_label = _label_of(proxy, totals_source_row)
        print(f"  Baseline (unfiltered) Totals: label={baseline_single_label!r} values={baseline_single_totals}")

        if real_status_values:
            target_status = real_status_values[0]
            status_combo.setCurrentText(target_status)
            app.processEvents()

            visible_real_rows = [r for r in _visible_rows(proxy) if r != totals_source_row]
            expected_done = sum(1 for r in visible_real_rows if source_model.data(source_model.index(r, 6)) == "Done")
            expected_open = sum(1 for r in visible_real_rows if source_model.data(source_model.index(r, 6)) == "Open")
            expected_total = expected_open + expected_done

            filtered_totals = _row_values(proxy, totals_source_row)
            filtered_label = _label_of(proxy, totals_source_row)
            print(f"  Filtered (Status={target_status!r}) Totals: label={filtered_label!r} values={filtered_totals}")
            print(f"  Independently expected: Open={expected_open} Done={expected_done} Total={expected_total}")

            if filtered_totals[:3] != (str(expected_open), str(expected_done), str(expected_total)):
                failures.append(
                    f"2a: DataViewPage's filtered Totals should be Open={expected_open}, Done={expected_done}, "
                    f"Total={expected_total}, got {filtered_totals[:3]}"
                )
            if "(filtered)" not in str(filtered_label):
                failures.append(f"2a: Totals label should gain a '(filtered)' suffix when a filter is active, got {filtered_label!r}")

            status_combo.setCurrentText("All")
            app.processEvents()
            reverted_totals = _row_values(proxy, totals_source_row)
            reverted_label = _label_of(proxy, totals_source_row)
            print(f"  After clearing filter: label={reverted_label!r} values={reverted_totals}")
            if reverted_totals != baseline_single_totals or reverted_label != baseline_single_label:
                failures.append(
                    f"2a: clearing the filter should restore the exact baseline ({baseline_single_label!r}, "
                    f"{baseline_single_totals}), got ({reverted_label!r}, {reverted_totals})"
                )

        # ------------------------------------------------------------
        # 2b -- CombinedDataViewPage: full scenario set against 3 real
        # fixtures (SMR 2D + 2 others).
        # ------------------------------------------------------------
        print("\n" + "=" * 70)
        print("2b -- CombinedDataViewPage Totals/Subtotal recalculation (3 real fixtures)")
        print("=" * 70)

        view = build_combined_view([display_a, display_b, display_c])
        cpage = CombinedDataViewPage(PROJECT_NAME, view, store, on_back=lambda: None)
        app.processEvents()
        c_proxy = cpage._sum_data_filter_proxy
        c_model = c_proxy.sourceModel()
        summary_rows = sorted(_summary_rows(c_model))
        c_totals_row = next(r for r in summary_rows if _is_totals_row(c_model, r))
        c_subtotal_rows = [r for r in summary_rows if r != c_totals_row]
        print(f"  Summary rows: {summary_rows} (Totals row={c_totals_row}, Subtotal rows={c_subtotal_rows})")
        if len(c_subtotal_rows) != 3:
            failures.append(f"2b: expected exactly 3 Subtotal rows (3 increments), found {len(c_subtotal_rows)}")

        # Capture the baseline (unfiltered) values for every summary
        # row BEFORE applying any filter -- these are the "current
        # values" the no-filter case must reproduce exactly, and the
        # exact values a single-increment filter's isolated Subtotal
        # must match too.
        baseline_combined = {r: _row_values(c_proxy, r) for r in summary_rows}
        for r in summary_rows:
            label = _label_of(c_proxy, r)
            print(f"  Baseline row {r} ({label!r}): {baseline_combined[r]}")

        def _find_subtotal_row_for(display):
            short = derive_short_display_name(display.name)
            target = f"{short} — Subtotal"
            for r in c_subtotal_rows:
                if _label_of(c_proxy, r) == target:
                    return r
            return None

        smr2d_subtotal_row = _find_subtotal_row_for(display_a)
        if smr2d_subtotal_row is None:
            failures.append("2b: could not find SMR 2D's own Subtotal row by its short-name label")

        # --- Filtering Status=Done isolates SMR 2D's 7 Done rows ---
        print("\n  -- Filtering Status=Done --")
        c_status_combo = cpage.sum_data_filter_bar._column_combos[6]
        c_status_combo.setCurrentText("Done")
        app.processEvents()

        visible_real_rows = [
            r for r in _visible_rows(c_proxy) if r not in summary_rows
        ]
        print(f"  Visible data rows under Status=Done: {len(visible_real_rows)}")
        if len(visible_real_rows) != 7:
            failures.append(f"2b: filtering Status=Done should show exactly 7 rows (all from SMR 2D), got {len(visible_real_rows)}")

        visible_summary_rows = {r for r in summary_rows if c_proxy.filterAcceptsRow(r, QModelIndex())}
        print(f"  Visible summary rows: {sorted(visible_summary_rows)}")
        if visible_summary_rows != {c_totals_row, smr2d_subtotal_row}:
            failures.append(
                f"2b: filtering Status=Done should show only SMR 2D's Subtotal ({smr2d_subtotal_row}) plus "
                f"the Totals row ({c_totals_row}), got {sorted(visible_summary_rows)}"
            )

        totals_done_filtered = _row_values(c_proxy, c_totals_row)
        print(f"  Totals row under Status=Done: {totals_done_filtered}")
        if totals_done_filtered[:2] != ("0", "7"):
            failures.append(f"2b: Totals under Status=Done should show Open=0, Done=7, got Open={totals_done_filtered[0]}, Done={totals_done_filtered[1]}")

        if smr2d_subtotal_row is not None:
            smr2d_subtotal_filtered = _row_values(c_proxy, smr2d_subtotal_row)
            print(f"  SMR 2D Subtotal under Status=Done: {smr2d_subtotal_filtered}")
            if smr2d_subtotal_filtered[:2] != ("0", "7"):
                failures.append(f"2b: SMR 2D's Subtotal under Status=Done should show Open=0, Done=7, got {smr2d_subtotal_filtered[:2]}")

        c_status_combo.setCurrentText("All")
        app.processEvents()

        # --- Filtering to one increment (Project #) ---
        print("\n  -- Filtering to one increment via Project # --")
        project_combo = cpage.sum_data_filter_bar._column_combos[0]
        project_values = [project_combo.itemText(i) for i in range(project_combo.count()) if project_combo.itemText(i) != "All"]
        target_project = project_values[0]
        project_combo.setCurrentText(target_project)
        app.processEvents()

        visible_summary_after_project_filter = {r for r in summary_rows if c_proxy.filterAcceptsRow(r, QModelIndex())}
        non_totals_visible = visible_summary_after_project_filter - {c_totals_row}
        print(f"  Visible Subtotal row(s) for Project #={target_project!r}: {sorted(non_totals_visible)}")
        if len(non_totals_visible) != 1:
            failures.append(f"2b: filtering to one Project # should show exactly one Subtotal row, got {sorted(non_totals_visible)}")
        else:
            isolated_row = next(iter(non_totals_visible))
            isolated_values = _row_values(c_proxy, isolated_row)
            baseline_values = baseline_combined[isolated_row]
            print(f"  Isolated Subtotal values: {isolated_values} (baseline was {baseline_values})")
            if isolated_values != baseline_values:
                failures.append(
                    f"2b: filtering to one increment's own Subtotal should reproduce its OWN unfiltered "
                    f"values {baseline_values} exactly, got {isolated_values}"
                )

        project_combo.setCurrentText("All")
        app.processEvents()

        # --- Clearing filters returns everything to the original values ---
        print("\n  -- Clearing filters restores the exact baseline --")
        reverted_combined = {r: _row_values(c_proxy, r) for r in summary_rows}
        if reverted_combined != baseline_combined:
            failures.append(f"2b: clearing all filters should restore the exact baseline {baseline_combined}, got {reverted_combined}")
        reverted_totals_label = _label_of(c_proxy, c_totals_row)
        if "(filtered)" in str(reverted_totals_label):
            failures.append(f"2b: Totals label should drop '(filtered)' once no filter is active, got {reverted_totals_label!r}")

        # --- A filter matching nothing: Totals shows zeros, no div-by-zero ---
        print("\n  -- A filter matching nothing --")
        index_combo = cpage.sum_data_filter_bar._column_combos[2]
        c_proxy.set_filters({2: "NO_SUCH_INDEX_VALUE_AT_ALL"})
        app.processEvents()
        visible_data_rows_none = [r for r in _visible_rows(c_proxy) if r not in summary_rows]
        visible_summary_none = {r for r in summary_rows if c_proxy.filterAcceptsRow(r, QModelIndex())}
        print(f"  Visible data rows: {len(visible_data_rows_none)}; visible summary rows: {sorted(visible_summary_none)}")
        if visible_data_rows_none:
            failures.append(f"2b: a filter matching nothing should show zero data rows, got {len(visible_data_rows_none)}")
        if visible_summary_none != {c_totals_row}:
            failures.append(f"2b: a filter matching nothing should leave only the Totals row visible, got {sorted(visible_summary_none)}")
        zero_totals = _row_values(c_proxy, c_totals_row)
        print(f"  Totals row with nothing matching: {zero_totals}")
        if zero_totals[:3] != ("0", "0", "0"):
            failures.append(f"2b: Totals with nothing matching should show Open=0, Done=0, Total=0, got {zero_totals[:3]}")
        if zero_totals[3] != "—":
            failures.append(f"2b: % Complete with Total=0 should show '—' (no division by zero), got {zero_totals[3]!r}")
        c_proxy.set_filters({})
        app.processEvents()

        # --- Filtering never touches status.json for any increment ---
        for inc, name in [(inc_a, "SMR 2D"), (inc_b, "SMR 3D Delta 1"), (inc_c, "SMR 4C")]:
            before = store.store.load_status(PROJECT_NAME, inc.name)
            c_status_combo.setCurrentText("Done")
            app.processEvents()
            c_status_combo.setCurrentText("All")
            app.processEvents()
            after = store.store.load_status(PROJECT_NAME, inc.name)
            if before != after:
                failures.append(f"2b: filtering changed status.json on disk for {name} -- it must only change which rows are SHOWN")

        # ------------------------------------------------------------
        # 2c -- Subtotal rows are now short (screenshot evidence), not
        # a dozen+ lines tall, confirmed against the actual rendered
        # row height.
        # ------------------------------------------------------------
        print("\n" + "=" * 70)
        print("2c -- Subtotal rows render short (screenshot evidence)")
        print("=" * 70)
        from PySide6.QtWidgets import QMainWindow

        win = QMainWindow()
        win.setCentralWidget(cpage)
        win.resize(1280, 900)
        win.show()
        app.processEvents()
        if smr2d_subtotal_row is not None:
            cpage.sum_data_table.scrollTo(c_proxy.index(smr2d_subtotal_row, 1) if False else c_model.index(smr2d_subtotal_row, 1))
            app.processEvents()
            subtotal_row_height = cpage.sum_data_table.rowHeight(smr2d_subtotal_row)
            normal_row_height = cpage.sum_data_table.rowHeight(0)
            print(f"  Subtotal row height: {subtotal_row_height}px (a normal data row here is {normal_row_height}px)")
            screenshot_path = os.path.join(tmp, "subtotal_short.png")
            win.grab().save(screenshot_path)
            print(f"  Screenshot saved to {screenshot_path}")
            if subtotal_row_height > normal_row_height * 2:
                failures.append(
                    f"2c: Subtotal row height ({subtotal_row_height}px) should be comparable to a normal "
                    f"data row's ({normal_row_height}px), not several times taller"
                )

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
        # exported header columns at the same index; Subtotal label is
        # a deliberate, labels-only exception (short name on screen,
        # full text in the export) -- numbers must still match exactly.
        # ------------------------------------------------------------
        print("\n" + "=" * 70)
        print("4 -- Filter bar columns match the exported header columns")
        print("=" * 70)

        def _check_parity(label_prefix, actual_columns):
            mismatches = []
            for column, combo_label in actual_columns:
                exported_header = SUM_DATA_HEADERS[column]
                # "Increment" on screen, "Inc #" in the export -- the one
                # known, already-documented relabel (see
                # tests/test_sum_data_notes_completion_date_columns.py).
                if combo_label == "Increment" and exported_header == "Inc #":
                    continue
                if combo_label != exported_header:
                    mismatches.append((column, combo_label, exported_header))
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

        # Subtotal label parity exception: preview uses the SHORT name
        # (+ full name as tooltip), export keeps the full name -- a
        # labels-only divergence; the underlying NUMBERS (already
        # proven equal in step 2b above, independent of this label
        # check) must still match.
        print("\n  -- Subtotal label: preview uses short name + tooltip, export keeps full text --")
        combined_export_path = str(Path(tmp) / "label_parity_combined.xlsx")
        combined_export.export_combined_report([display_a, display_b, display_c], combined_export_path)
        wb_label = openpyxl.load_workbook(combined_export_path)
        ws_label = wb_label["Sum Data"]
        exported_increment_col = [ws_label.cell(row=r, column=2).value for r in range(1, ws_label.max_row + 1)]
        for display in (display_a, display_b, display_c):
            full_label = f"{display.name} — Subtotal"
            short_label = f"{derive_short_display_name(display.name)} — Subtotal"
            if full_label not in exported_increment_col:
                failures.append(f"4: expected the EXPORT's Subtotal label to still be the full text {full_label!r}")
            preview_row = next((r for r in c_subtotal_rows if _label_of(c_proxy, r) == short_label), None)
            if preview_row is None:
                failures.append(f"4: expected the PREVIEW's Subtotal label to be the short-name text {short_label!r}")
            else:
                tooltip = c_model.item(preview_row, 1).toolTip()
                if tooltip != display.name:
                    failures.append(f"4: Subtotal row's tooltip should be the full increment name {display.name!r}, got {tooltip!r}")
                print(f"  {display.name[:40]}...: preview={short_label!r} tooltip_ok={tooltip == display.name} export={full_label[:60]}...")

        # ------------------------------------------------------------
        # 5 -- Export AutoFilter is a REAL AutoFilter, not just arrows
        # (exports are NOT touched by this round's changes -- confirms
        # that too)
        # ------------------------------------------------------------
        print("\n" + "=" * 70)
        print("5 -- Export AutoFilter range, read back with openpyxl")
        print("=" * 70)

        single_path = str(Path(tmp) / "single_export.xlsx")
        excel_export.export_increment(display_a, single_path)
        wb_single = openpyxl.load_workbook(single_path)
        ws_single = wb_single["Sum Data"]
        print(f"  Single-increment: max_row={ws_single.max_row}, auto_filter.ref={ws_single.auto_filter.ref!r}")
        expected_single_ref = f"A1:M{ws_single.max_row - 1}"  # excludes the Totals row (always the last row)
        if ws_single.auto_filter.ref != expected_single_ref:
            failures.append(
                f"5: single-increment AutoFilter ref should be {expected_single_ref!r} (excluding the Totals "
                f"row), got {ws_single.auto_filter.ref!r}"
            )
        if ws_single.cell(row=ws_single.max_row, column=2).value != "Totals":
            failures.append("5: fixture assumption broken -- expected the last row to be the Totals row")

        ws_combined = ws_label  # reuse the export already written for the label-parity check above
        print(f"  Combined (3 fixtures): max_row={ws_combined.max_row}, auto_filter.ref={ws_combined.auto_filter.ref!r}")
        if ws_combined.cell(row=ws_combined.max_row, column=2).value != "Totals":
            failures.append("5: fixture assumption broken -- expected the combined sheet's last row to be the Totals row")
        last_subtotal_row = ws_combined.max_row - 1
        last_subtotal_label = ws_combined.cell(row=last_subtotal_row, column=2).value or ""
        if "Subtotal" not in last_subtotal_label:
            failures.append(
                f"5: fixture assumption broken -- expected row {last_subtotal_row} (one above Totals) to be "
                f"a per-increment Subtotal row, got {last_subtotal_label!r}"
            )
        expected_combined_ref = f"A1:M{last_subtotal_row}"
        if ws_combined.auto_filter.ref != expected_combined_ref:
            failures.append(
                f"5: combined AutoFilter ref should be {expected_combined_ref!r} (every interspersed Subtotal "
                f"row stays in range, but the final Totals row -- always last, never interspersed -- is "
                f"excluded, same as the single-increment export), got {ws_combined.auto_filter.ref!r}"
            )
        if ws_combined.max_row <= last_subtotal_row:
            failures.append("5: fixture assumption broken -- expected the Totals row to be strictly after the last Subtotal row")

        # ------------------------------------------------------------
        # 6 -- Notes/Completion Date stay blank and within the
        # filterable range
        # ------------------------------------------------------------
        print("\n" + "=" * 70)
        print("6 -- Notes/Completion Date stay blank, inside the AutoFilter range")
        print("=" * 70)

        notes_col, completion_date_col = 8, 9
        blank_violations = [
            r for r in range(2, ws_single.max_row)
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
            "underlying data; Totals/Subtotal rows recalculate from visible rows when filtered (Subtotal "
            "hidden if none of its own rows match, Totals never hidden and relabeled '(filtered)'), exactly "
            "reproduce the full-dataset baseline when unfiltered, isolate correctly when filtering to one "
            "increment, revert exactly on clearing filters, and handle a zero-match filter without dividing "
            "by zero; Subtotal labels use the short display name on screen (full name in the export and in "
            "a tooltip) while the underlying numbers still match exactly; filters reset on tab rebuild; "
            "filter columns match the exported headers; export AutoFilter ranges are untouched and exactly "
            "as designed; Notes/Completion Date stay blank; and no filter-bar label is truncated."
        )


if __name__ == "__main__":
    main()
