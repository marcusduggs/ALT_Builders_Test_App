"""
Regression test: for EVERY one of the 13 real HMR/SMR fixtures,
combined into one CombinedDataViewPage, filtering down to that one
increment (via its own Project #) must leave EXACTLY that increment's
own rows visible -- the right count, and every visible row genuinely
belonging to it -- with no other increment's rows leaking in or being
hidden/dropped.

Sum Data has no Subtotal/Totals rows to recalculate anymore (Rey's
TIO_Changes.xlsx "Sum Data" tab change request #3), so this is a
simpler, more direct guarantee than the version of this test that used
to check a recalculated Subtotal row's numbers -- there's no aggregate
row left to get wrong; the row SET itself is the only thing to verify.

Explicitly includes the two fixtures already known to have real
structural edge cases from earlier rounds of this feature:
  - HMR 10A (TIO_H240649-38-00 HMR 10A.xlsm): previously had a
    duplicate-grid-index collapse bug (required_stages_by_index()),
    now fixed -- not touched by this filtering feature either way.
  - SMR 4C (TIO S251652-38-00 SMR 4C.xlsm): a genuinely blank
    F-Cons Verif row excluded from Sum Data entirely (an earlier
    blank-row-exclusion fix).

Neither edge case is specific to filtering -- both affect what rows
exist/are counted BEFORE any filter is ever applied -- but the point
of this test is exactly that: confirm the filtering feature built on
top of whatever rows those earlier fixes already produce works
correctly for ALL 13 real files, not just a handful used elsewhere.

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

from core.excel_reader import expand_sum_data_rows
from core.project_store import ProjectStore
from ui.mock_data import MockDataStore, build_combined_view
from ui.pages.combined_data_view_page import CombinedDataViewPage

FIXDIR = os.path.join(os.path.dirname(__file__), "fixtures", "hmr_smr_projects")
PROJECT_NAME = "All-Fixtures Filter Regression"
PROJECT_COL = 0
INDEX_COL = 2
STAGE_COL = 5
STATUS_COL = 6


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
        total_rows = model.rowCount()
        expected_total_rows = sum(len(expand_sum_data_rows(d.sum_data)) for d in displays)
        print(f"Total Sum Data rows across all 13 increments: {total_rows} (expected {expected_total_rows})")
        if total_rows != expected_total_rows:
            failures.append(f"expected exactly {expected_total_rows} total rows (no Subtotal/Totals rows mixed in), got {total_rows}")

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
        # alphabetically-sorted dropdown.
        for display in displays:
            project_number = str(display.project_number)
            short_name = display.name[:60]
            print(f"\n--- {short_name}... (Project #={project_number!r}) ---")

            project_combo.setCurrentText(project_number)
            app.processEvents()

            visible_rows = _visible_rows(proxy)
            expected_expanded = expand_sum_data_rows(display.sum_data)
            expected_rows = len(expected_expanded)
            print(f"  visible rows: {len(visible_rows)} (expected {expected_rows})")
            if len(visible_rows) != expected_rows:
                failures.append(
                    f"{short_name}: filtering to Project #={project_number!r} should show exactly "
                    f"{expected_rows} rows (this increment's own full row count), got {len(visible_rows)}"
                )
            # Every visible row must actually belong to this Project #.
            bad_rows = [
                r for r in visible_rows if model.data(model.index(r, PROJECT_COL)) != project_number
            ]
            if bad_rows:
                failures.append(f"{short_name}: row(s) {bad_rows} passed the Project # filter but don't actually have that Project #")

            # Not just a count/Project# check -- the actual SET of
            # (Index, Stage #, Status) tuples visible must equal exactly
            # what expand_sum_data_rows() produces for THIS increment's
            # own sum_data, so a bug that shuffled another increment's
            # row content under this Project # label (while leaving the
            # Project # cell itself correct) would still be caught.
            expected_triples = sorted(
                (str(r["index"]), str(r["stage"]), r["status"]) for r in expected_expanded
            )
            visible_triples = sorted(
                (
                    str(model.data(model.index(r, INDEX_COL))),
                    str(model.data(model.index(r, STAGE_COL))),
                    model.data(model.index(r, STATUS_COL)),
                )
                for r in visible_rows
            )
            if visible_triples != expected_triples:
                missing = set(expected_triples) - set(visible_triples)
                extra = set(visible_triples) - set(expected_triples)
                failures.append(
                    f"{short_name}: visible (Index, Stage #, Status) rows under Project #={project_number!r} "
                    f"don't match expand_sum_data_rows()'s own result exactly -- missing={missing}, extra={extra}"
                )

            project_combo.setCurrentText("All")
            app.processEvents()

        # Confirm clearing the last filter restored full visibility.
        reverted_visible = _visible_rows(proxy)
        if len(reverted_visible) != total_rows:
            failures.append(f"after clearing the filter, all {total_rows} rows should be visible again, got {len(reverted_visible)}")

    print("\n" + "=" * 70)
    if failures:
        print("RESULT: FAIL")
        for f in failures:
            print(" -", f)
        raise SystemExit(1)
    else:
        print(
            "RESULT: PASS -- for all 13 real HMR/SMR fixtures (including HMR 10A's once-duplicate-index rows "
            "and SMR 4C's blank-row exclusion), filtering the combined Sum Data view down to one increment's "
            "Project # shows exactly that increment's own rows -- the right count, with no other increment's "
            "rows leaking in or being dropped -- and clearing the filter restores full visibility."
        )


if __name__ == "__main__":
    main()
