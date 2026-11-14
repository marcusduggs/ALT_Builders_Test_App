"""
Regression test for the Sum Data filter bar's width/popup/subtitle
layout bugs, confirmed via screenshot against all 13 real HMR/SMR
fixtures combined:

  - Uncapped QComboBox sizeHint() (Qt sizes a combo to fit its single
    longest item, with no cap) made the filter bar's own
    minimumSizeHint() 2050px wide (Increment combo alone: 636px;
    Approval Agency: 556px), which propagated straight up through
    the page and the window, forcing it wider than any normal laptop
    screen.
  - The dropdown popup defaulted to the closed combo's own width, so
    real long Index entries (e.g. "B-C14.1") were cut off in the open
    list, and with 140 distinct Index values and no visible-item cap,
    the popup ran off both the top and bottom of the screen.
  - CombinedDataViewPage's header subtitle comma-joined every selected
    increment's name, growing to 10+ wrapped lines with 13 real
    increments and pushing the table down the page.

A second, PREVIOUSLY-MASKED bug surfaced only once the filter bar
itself was fixed: DataViewPage's own page-title QLabel had no cap or
eliding at all, and a real HMR/SMR increment's already-shortened
display name plus " — Version N" still measured 794px wide in the
page's actual bold 18px title font -- alone enough to push the whole
header row (and so the window) past 1280px. Fixed with
ui.pages.data_view_page._ElidingLabel; covered here too.

Run directly: `python tests/test_sum_data_filter_bar_layout.py`
"""

import glob
import os
import sys
import tempfile
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from PySide6.QtGui import QFontMetrics
from PySide6.QtWidgets import QApplication, QMainWindow

from core.project_store import ProjectStore
from ui.mock_data import MockDataStore, build_combined_view
from ui.pages.combined_data_view_page import CombinedDataViewPage
from ui.pages.data_view_page import DataViewPage

FIXDIR = os.path.join(os.path.dirname(__file__), "fixtures", "hmr_smr_projects")
PROJECT_NAME = "Filter Bar Layout Regression"
MAX_WINDOW_WIDTH = 1280


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

    with tempfile.TemporaryDirectory() as tmp:
        store = MockDataStore(ProjectStore(base_dir=Path(tmp) / "data"))
        store.add_project(PROJECT_NAME)
        displays = []
        for f in files:
            inc = store.add_new_increment(PROJECT_NAME, f)
            displays.append(store.get_increment_for_display(PROJECT_NAME, inc.name))
        view = build_combined_view(displays)

        for resolution in [(1280, 800), (1920, 1080)]:
            w, h = resolution
            print("\n" + "=" * 70)
            print(f"Resolution {w}x{h}")
            print("=" * 70)

            pages = [
                ("CombinedDataViewPage", CombinedDataViewPage(PROJECT_NAME, view, store, on_back=lambda: None)),
                ("DataViewPage", DataViewPage(PROJECT_NAME, displays[0], store, on_back=lambda: None)),
            ]
            for page_name, page in pages:
                win = QMainWindow()
                win.setCentralWidget(page)
                win.resize(w, h)
                win.show()
                app.processEvents()

                min_width = win.minimumSizeHint().width()
                print(f"  {page_name}: win.minimumSizeHint().width()={min_width}")
                if min_width >= MAX_WINDOW_WIDTH:
                    failures.append(
                        f"{page_name} @ {w}x{h}: window minimumSizeHint width {min_width} should stay under "
                        f"{MAX_WINDOW_WIDTH}, same as a normal laptop screen"
                    )

                filter_bar = page.sum_data_filter_bar
                combos = filter_bar._column_combos
                # DataViewPage (one increment at a time) deliberately
                # shows only 4 filters -- Project #/Increment are
                # dropped there since every row shares the same value
                # (see ui.pages.data_view_page.
                # SINGLE_INCREMENT_SUM_DATA_FILTER_COLUMNS). Combined
                # keeps all 6.
                expected_count = 4 if page_name == "DataViewPage" else 6
                if len(combos) != expected_count:
                    failures.append(f"{page_name}: expected {expected_count} filter combos, found {len(combos)}")
                for col, combo in combos.items():
                    if not combo.isVisible():
                        failures.append(f"{page_name} @ {w}x{h}: filter combo for column {col} is not visible")
                    if combo.x() + combo.width() > win.width():
                        failures.append(
                            f"{page_name} @ {w}x{h}: filter combo for column {col} extends past the window's "
                            f"right edge (x={combo.x()}, width={combo.width()}, window width={win.width()})"
                        )

                # The Index combo (column 2) has the most distinct real
                # values (140 across these 13 fixtures) -- the popup
                # must widen to fit the longest one in full, and cap
                # its own visible rows (confirmed separately via
                # verticalScrollBar().maximum() > 0 -- i.e. there IS
                # more content than fits, so scrolling is possible,
                # rather than the popup just growing to show everything
                # unbounded).
                index_combo = combos[2]
                index_combo.showPopup()
                app.processEvents()
                popup = index_combo.view().window()
                if popup is None:
                    failures.append(f"{page_name} @ {w}x{h}: Index combo's popup window was never created")
                else:
                    metrics = QFontMetrics(index_combo.font())
                    longest_item = max(
                        (index_combo.itemText(i) for i in range(index_combo.count())),
                        key=metrics.horizontalAdvance,
                    )
                    required_width = metrics.horizontalAdvance(longest_item)
                    print(
                        f"    Index popup: width={popup.width()} height={popup.height()} "
                        f"longest_item={longest_item!r} required_text_width={required_width}"
                    )
                    if popup.width() < required_width:
                        failures.append(
                            f"{page_name} @ {w}x{h}: Index popup width {popup.width()} is narrower than its "
                            f"longest item {longest_item!r} needs ({required_width}) -- would be cut off"
                        )
                    screen_height = app.primaryScreen().geometry().height()
                    if popup.y() < 0 or (popup.y() + popup.height()) > screen_height + 50:
                        failures.append(
                            f"{page_name} @ {w}x{h}: Index popup (y={popup.y()}, height={popup.height()}) "
                            f"extends off-screen (screen height={screen_height})"
                        )
                    scrollbar = index_combo.view().verticalScrollBar()
                    if scrollbar.maximum() <= 0:
                        failures.append(
                            f"{page_name} @ {w}x{h}: expected the Index popup (140 distinct values) to need "
                            f"scrolling (maxVisibleItems should cap it), but its scrollbar has no range"
                        )
                index_combo.hidePopup()
                app.processEvents()

        # ------------------------------------------------------------
        # Combined header subtitle: capped to first 3 names + count,
        # full list still reachable via tooltip
        # ------------------------------------------------------------
        print("\n" + "=" * 70)
        print("Combined header subtitle")
        print("=" * 70)
        cpage = CombinedDataViewPage(PROJECT_NAME, view, store, on_back=lambda: None)
        app.processEvents()
        from PySide6.QtWidgets import QLabel

        subtitle = next(
            lbl for lbl in cpage.findChildren(QLabel)
            if lbl.objectName() == "hint" and lbl.text().startswith(PROJECT_NAME)
        )
        print(f"  Subtitle text: {subtitle.text()!r}")
        print(f"  Subtitle length: {len(subtitle.text())} chars (13 full names comma-joined would be 662+)")
        if len(subtitle.text()) > 300:
            failures.append(
                f"Combined subtitle is {len(subtitle.text())} chars long -- expected it capped to the first "
                f"3 names + 'and N more', not every one of the 13 increment names"
            )
        if "more" not in subtitle.text():
            failures.append("Combined subtitle with 13 increments should say '... and N more'")
        tooltip = subtitle.toolTip()
        full_names_present = sum(1 for inc in view.increments if inc.name in tooltip)
        if full_names_present != len(view.increments):
            failures.append(
                f"Combined subtitle's tooltip should still contain all {len(view.increments)} full increment "
                f"names, found {full_names_present}"
            )

        # ------------------------------------------------------------
        # Single-increment title: elided (single ellipsis, not doubled),
        # full name preserved as tooltip
        # ------------------------------------------------------------
        print("\n" + "=" * 70)
        print("Single-increment title eliding")
        print("=" * 70)
        dpage = DataViewPage(PROJECT_NAME, displays[0], store, on_back=lambda: None)
        win = QMainWindow()
        win.setCentralWidget(dpage)
        win.resize(1280, 800)
        win.show()
        app.processEvents()
        title_text = dpage.title_label.text()
        print(f"  Title label text: {title_text!r}")
        if "……" in title_text:
            failures.append(f"Title label has a doubled ellipsis: {title_text!r}")
        if dpage.title_label.toolTip() != displays[0].name:
            failures.append("Title label's tooltip should be the full, untruncated increment name")

    print("\n" + "=" * 70)
    if failures:
        print("RESULT: FAIL")
        for f in failures:
            print(" -", f)
        raise SystemExit(1)
    else:
        print(
            "RESULT: PASS -- Sum Data filter bar stays under 1280px wide with all 13 real fixtures combined, "
            "all 6 filters are visible and within the window at both 1280x800 and 1920x1080, the Index "
            "popup widens to fit its longest real item and caps/scrolls its visible rows without running "
            "off-screen, the combined header subtitle is capped with the full list still in its tooltip, "
            "and the single-increment title elides cleanly (no doubled ellipsis) with the full name in its "
            "own tooltip."
        )


if __name__ == "__main__":
    main()
