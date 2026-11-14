"""
Regression test: for EVERY one of the 13 real HMR/SMR fixtures,
combined into one CombinedDataViewPage, filtering down to that one
increment (via its own Project #) must leave exactly that increment's
Subtotal row visible, with its recalculated Open/Done/Total/%
Complete equal -- not approximately, exactly -- to that same
increment's own UNFILTERED, precomputed Subtotal (and the Totals row,
since with only one increment visible the overall total and that
increment's own total are the same thing).

Explicitly includes the two fixtures already known to have real
structural edge cases from earlier rounds of this feature:
  - HMR 10A (TIO_H240649-38-00 HMR 10A.xlsm): 39 duplicate grid
    indices (required_stages_by_index()'s own, separately-scoped
    collapsing bug -- NOT touched by this filtering feature).
  - SMR 4C (TIO S251652-38-00 SMR 4C.xlsm): a genuinely blank
    F-Cons Verif row excluded from Sum Data entirely (the earlier
    blank-row-exclusion fix).

Neither edge case is specific to filtering -- both affect what rows
exist/are counted BEFORE any filter is ever applied -- but the point
of this test is exactly that: confirm the filtering feature built on
top of whatever rows/counts those earlier fixes already produce
recalculates correctly for ALL 13 real files, not just the ones used
in the feature's own original 3-fixture verification.

Run directly: `python tests/test_sum_data_filter_all_fixtures.py`
"""

import glob
import os
import sys
import tempfile
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from PySide6.QtCore import QModelIndex
from PySide6.QtWidgets import QApplication

from core.project_store import ProjectStore
from ui.mock_data import MockDataStore, build_combined_view
from ui.pages.combined_data_view_page import CombinedDataViewPage
from ui.widgets.sum_data_filter_bar import SUMMARY_IS_TOTALS_ROLE, SUMMARY_ROW_ROLE

FIXDIR = os.path.join(os.path.dirname(__file__), "fixtures", "hmr_smr_projects")
PROJECT_NAME = "All-Fixtures Filter Regression"
NUMERIC_COLS = (9, 10, 11, 12)  # Open, Done, Total, % Complete
PROJECT_COL = 0


def _summary_rows(model) -> list[int]:
    return [r for r in range(model.rowCount()) if model.data(model.index(r, 1), SUMMARY_ROW_ROLE)]


def _is_totals_row(model, row: int) -> bool:
    return bool(model.data(model.index(row, 1), SUMMARY_IS_TOTALS_ROLE))


def _proxy_data(proxy, source_row: int, column: int):
    source_model = proxy.sourceModel()
    return proxy.data(proxy.mapFromSource(source_model.index(source_row, column)))


def _row_values(proxy, source_row: int) -> tuple:
    return tuple(_proxy_data(proxy, source_row, c) for c in NUMERIC_COLS)


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

    files = sorted(glob.glob(os.path.join(FIXDIR, "*")))
    print(f"Using all {len(files)} real fixtures")
    if len(files) != 13:
        failures.append(f"fixture assumption broken -- expected 13 real HMR/SMR files, found {len(files)}")

    required_names = ["HMR 10A", "SMR 4C"]
    for required in required_names:
        if not any(required.replace(" ", "") in os.path.basename(f).replace(" ", "").replace("_", "") for f in files):
            failures.append(f"fixture assumption broken -- expected a file matching {required!r} among the 13 fixtures")

    with tempfile.TemporaryDirectory() as tmp:
        store = MockDataStore(ProjectStore(base_dir=Path(tmp) / "data"))
        store.add_project(PROJECT_NAME)
        displays = []
        for f in files:
            inc = store.add_new_increment(PROJECT_NAME, f)
            displays.append(store.get_increment_for_display(PROJECT_NAME, inc.name))

        view = build_combined_view(displays)
        cpage = CombinedDataViewPage(PROJECT_NAME, view, store, on_back=lambda: None)
        app.processEvents()

        proxy = cpage._sum_data_filter_proxy
        model = proxy.sourceModel()
        summary_rows = _summary_rows(model)
        totals_row = next(r for r in summary_rows if _is_totals_row(model, r))
        subtotal_rows = [r for r in summary_rows if r != totals_row]
        print(f"Summary rows: {len(subtotal_rows)} Subtotal + 1 Totals (out of {model.rowCount()} total rows)")
        if len(subtotal_rows) != len(displays):
            failures.append(f"expected exactly {len(displays)} Subtotal rows (one per increment), found {len(subtotal_rows)}")

        # Baseline (unfiltered) Subtotal values and each increment's own
        # Project # -- captured once, up front, before any filter.
        baseline_by_subtotal_row: dict[int, tuple] = {r: _row_values(proxy, r) for r in subtotal_rows}

        project_combo = cpage.sum_data_filter_bar._column_combos[PROJECT_COL]
        project_values = [project_combo.itemText(i) for i in range(project_combo.count()) if project_combo.itemText(i) != "All"]
        print(f"Distinct Project # values: {len(project_values)}")
        if len(project_values) != len(displays):
            failures.append(
                f"expected {len(displays)} distinct Project # values (one per increment, each unique), "
                f"found {len(project_values)} -- can't reliably isolate one increment per iteration if "
                f"any two share a Project #"
            )
        if len(set(str(d.project_number) for d in displays)) != len(displays):
            failures.append("two or more of the 13 fixtures share the same Project #, which would make isolating 'this one increment' by Project # ambiguous")

        # Iterate `displays` directly and use EACH ONE'S OWN
        # project_number attribute -- NOT zip(displays, project_values),
        # which would silently pair display[i] with whatever Project #
        # happens to sort into position i in the filter combo's
        # alphabetically-sorted dropdown (a real bug caught while
        # writing this test: the printed increment-name/Project#
        # pairing was wrong, even though the pass/fail logic below
        # never actually depended on it -- only the print label did).
        for display in displays:
            project_number = str(display.project_number)
            short_name = display.name[:60]
            print(f"\n--- {short_name}... (Project #={project_number!r}) ---")

            project_combo.setCurrentText(project_number)
            app.processEvents()

            visible_data_rows = [r for r in _visible_rows(proxy) if r not in summary_rows]
            visible_summary = [r for r in summary_rows if proxy.filterAcceptsRow(r, QModelIndex())]
            visible_subtotals = [r for r in visible_summary if r != totals_row]

            if len(visible_subtotals) != 1:
                failures.append(
                    f"{short_name}: filtering to Project #={project_number!r} should isolate exactly one "
                    f"Subtotal row, found {len(visible_subtotals)} ({visible_subtotals})"
                )
                project_combo.setCurrentText("All")
                app.processEvents()
                continue

            isolated_row = visible_subtotals[0]
            baseline = baseline_by_subtotal_row[isolated_row]
            filtered_subtotal = _row_values(proxy, isolated_row)
            filtered_totals = _row_values(proxy, totals_row)

            print(f"  visible data rows: {len(visible_data_rows)}")
            print(f"  baseline (unfiltered) Subtotal:   {baseline}")
            print(f"  recalculated Subtotal (filtered): {filtered_subtotal}")
            print(f"  recalculated Totals (filtered):   {filtered_totals}")

            if filtered_subtotal != baseline:
                failures.append(
                    f"{short_name}: recalculated Subtotal {filtered_subtotal} != its own unfiltered "
                    f"baseline {baseline}"
                )
            if filtered_totals != baseline:
                failures.append(
                    f"{short_name}: with only this increment visible, the Totals row {filtered_totals} "
                    f"should also equal its baseline Subtotal {baseline}"
                )
            # Total (column index 2 within NUMERIC_COLS, i.e. the 3rd
            # value) must equal the number of visible data rows --
            # catches a recalculation that silently drops or
            # double-counts rows, independent of trusting the
            # baseline comparison alone.
            if int(filtered_subtotal[2]) != len(visible_data_rows):
                failures.append(
                    f"{short_name}: recalculated Total {filtered_subtotal[2]} should equal the "
                    f"{len(visible_data_rows)} actually-visible data rows"
                )

            project_combo.setCurrentText("All")
            app.processEvents()

        # Confirm clearing the last filter restored every Subtotal's
        # baseline exactly, not just the one checked last.
        reverted = {r: _row_values(proxy, r) for r in subtotal_rows}
        if reverted != baseline_by_subtotal_row:
            failures.append(f"after clearing the filter, Subtotal values should match the original baseline exactly; got {reverted}")

    print("\n" + "=" * 70)
    if failures:
        print("RESULT: FAIL")
        for f in failures:
            print(" -", f)
        raise SystemExit(1)
    else:
        print(
            "RESULT: PASS -- for all 13 real HMR/SMR fixtures (including HMR 10A's duplicate-index rows and "
            "SMR 4C's blank-row exclusion), filtering the combined view down to one increment isolates "
            "exactly that increment's Subtotal row, and its recalculated Open/Done/Total/% Complete exactly "
            "matches that same increment's own unfiltered, precomputed Subtotal -- no mismatches."
        )


if __name__ == "__main__":
    main()
