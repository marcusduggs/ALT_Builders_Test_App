"""
Regression test guarding against the Sum Data table's column widths
ever again being narrower than their own header text needs to render
without truncation -- this happened TWICE with hand-guessed pixel
values ("Completion Date" clipped to "ompletion Date"; then "%
Complete" clipped at its right edge), both times because the width was
picked without accounting for the REAL loaded stylesheet (bold header
weight + padding from ui/style.qss's QHeaderView::section rule).

Uses QHeaderView.sectionSizeHint() -- Qt's OWN authoritative size for a
header section, computed against whatever stylesheet/font/padding is
actually loaded -- rather than reconstructing that math by hand (a
third hand-guess would just risk a third version of the same bug).
Checks EVERY column in BOTH tables, not just the ones known to have
failed before, so a future header relabel or font change that
narrows the margin on some other column gets caught here first.

Run directly: `python tests/test_sum_data_column_widths.py`
"""

import os
import sys
import tempfile
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication

from core.project_store import ProjectStore
from ui.mock_data import MockDataStore, build_combined_view
from ui.pages.combined_data_view_page import CombinedDataViewPage
from ui.pages.data_view_page import DataViewPage

FIXTURE = os.path.join(os.path.dirname(__file__), "fixtures", "sample_increment.xlsm")
PROJECT_NAME = "Column Width Regression"


def _check_table(label, table, failures):
    header = table.horizontalHeader()
    model = table.model()
    print(f"\n{label}")
    for col in range(model.columnCount()):
        text = model.headerData(col, Qt.Horizontal)
        current = table.columnWidth(col)
        required = header.sectionSizeHint(col)
        status = "FAIL" if current < required else "pass"
        print(f"  col {col} {str(text)!r}: current={current}, required (sectionSizeHint)={required}  [{status}]")
        if current < required:
            failures.append(f"{label}: column {col} ({text!r}) width={current} < required={required}")


def main():
    failures = []

    app = QApplication.instance() or QApplication(sys.argv)
    style_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "ui", "style.qss")
    app.setStyleSheet(Path(style_path).read_text())  # the REAL stylesheet -- this is the whole point

    with tempfile.TemporaryDirectory() as tmp:
        store = MockDataStore(ProjectStore(base_dir=Path(tmp) / "data"))
        store.add_project(PROJECT_NAME)
        increment = store.add_new_increment(PROJECT_NAME, FIXTURE)
        display = store.get_increment_for_display(PROJECT_NAME, increment.name)

        page = DataViewPage(PROJECT_NAME, display, store, on_back=lambda: None)
        single_table = page.sum_data_table
        app.processEvents()
        _check_table("DataViewPage Sum Data table", single_table, failures)

        view = build_combined_view([display])
        cpage = CombinedDataViewPage(PROJECT_NAME, view, store, on_back=lambda: None)
        app.processEvents()
        _check_table("CombinedDataViewPage Sum Data table", cpage.sum_data_table, failures)

    print("\n" + "=" * 70)
    if failures:
        print("RESULT: FAIL")
        for f in failures:
            print(" -", f)
        raise SystemExit(1)
    else:
        print("RESULT: PASS -- every Sum Data column, in both DataViewPage and CombinedDataViewPage, is at "
              "least as wide as QHeaderView.sectionSizeHint() says it needs to be to render its own header "
              "text without truncation, against the REAL loaded stylesheet.")


if __name__ == "__main__":
    main()
