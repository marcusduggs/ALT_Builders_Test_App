"""
Reusable Sum Data column-filter bar, plus a proxy model that hides rows
not matching the active per-column filters -- used by both
ui.pages.data_view_page.DataViewPage and
ui.pages.combined_data_view_page.CombinedDataViewPage's Sum Data tables.

*** No per-column header dropdown arrows (client's reference sheet
style) -- confirmed not feasible with this app's table widgets, not
just assumed ***
QHeaderView inherits setIndexWidget()/indexWidget() from
QAbstractItemView, but empirically (tested directly against this app's
own QTableView/QHeaderView stack) calling it on a header produces a
widget with geometry QRect(0, 0, 0, 0) -- zero size, never actually
positioned over the section. QHeaderView overrides painting/geometry
for its own section-based layout and never lays out index widgets the
way a view's body does; there is no supported way to embed a working
interactive widget in a header section. A real Excel-style dropdown
arrow would require a hand-built QHeaderView subclass manually
positioning child QToolButtons per section (paintSection/resizeEvent/
sectionResized), which would ALSO have to stay in sync with this app's
existing FrozenTableView overlay (a second QTableView sharing the same
model, with its own separate header) -- real, ongoing fragility for a
compliance-reporting desktop app. A filter bar above the table is used
instead, per the explicit fallback instruction.

*** Filtering never touches the underlying data ***
QSortFilterProxyModel is a read-only VIEW transformation over its
source model -- it can only hide/show/reorder rows as seen through the
proxy; it has no mechanism to write back to the source model, and this
module never calls any source-model mutator. Increment.sum_data and
status.json are therefore unreachable from here by construction, not
just by convention.

*** Summary-row recalculation removed ***
Sum Data no longer has Subtotal or Totals rows at all (Rey's own
TIO_Changes.xlsx "Sum Data" tab change request #3: "Remove Subtotals,
and Total" -- Sum Data is now a flat list of data rows only, in both
the in-app view and the export). This module used to carry a
SUMMARY_ROW_ROLE-aware QSortFilterProxyModel subclass that recomputed a
summary row's Open/Done/Total/% Complete from its scope's currently-
visible rows -- entirely moot now that there is no summary row (and no
Open/Done/Total/% Complete columns at all -- change request #2). What
remains is exactly the plain per-column exact-match filtering every
row (summary or otherwise) always needed underneath that logic.
"""

from __future__ import annotations

from PySide6.QtCore import QModelIndex, QSortFilterProxyModel, Qt, Signal
from PySide6.QtGui import QFontMetrics
from PySide6.QtWidgets import QComboBox, QGridLayout, QLabel, QVBoxLayout, QWidget

ALL_VALUES_LABEL = "All"


class SumDataFilterProxyModel(QSortFilterProxyModel):
    """Hides a data row unless every active per-column filter matches it
    (AND across columns, exact text match per column).
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self._filters: dict[int, str] = {}

    def set_filters(self, filters: dict[int, str]) -> None:
        self._filters = dict(filters)
        self.invalidateFilter()

    def _row_matches_filters(self, source_row: int, source_parent=QModelIndex()) -> bool:
        model = self.sourceModel()
        for column, required_value in self._filters.items():
            cell_text = model.data(model.index(source_row, column, source_parent), Qt.DisplayRole)
            if cell_text != required_value:
                return False
        return True

    def filterAcceptsRow(self, source_row: int, source_parent) -> bool:
        if self.sourceModel() is None:
            return True
        return self._row_matches_filters(source_row, source_parent)


def _distinct_column_values(model, column: int) -> list[str]:
    """Every distinct non-blank value a column has across the FULL
    (unfiltered) dataset -- computed once when the filter bar is built,
    not re-narrowed as other filters change (simpler and more
    predictable than Excel's own cascading dropdowns).
    """
    values: set[str] = set()
    for row in range(model.rowCount()):
        text = model.data(model.index(row, column), Qt.DisplayRole)
        if text:
            values.add(str(text))
    return sorted(values)


# Each combo's own CLOSED-box width is capped here, independent of how
# long its longest item is -- confirmed necessary by measurement: with
# all 13 real HMR/SMR fixtures combined, the uncapped "Increment" combo
# alone sized itself to 636px and "Approval Agency" to 556px (Qt's
# default QComboBox sizeHint grows to fit the single longest item, with
# no cap), making the filter bar's own minimumSizeHint() 2050px wide --
# which propagated straight up to the whole page/window's
# minimumSizeHint(), forcing the window wider than any normal laptop
# screen. The closed box shows an elided ("...") version of its current
# selection when it's longer than this cap -- standard, automatic
# QComboBox behavior, not something this module has to implement.
_COMBO_MAX_WIDTH = 170
_COMBO_MIN_CONTENTS_LENGTH = 12
# Caps how many rows show at once in the OPEN dropdown before it
# scrolls -- confirmed necessary: the real Index combo across all 13
# fixtures has 140 distinct values, which (with no cap) sized its popup
# tall enough to run off both the top and bottom of the screen with no
# way to scroll back to items pushed out of view.
_COMBO_MAX_VISIBLE_ITEMS = 15
# How many filters share one row before wrapping to the next -- 6
# filters split 3-and-3 comfortably clears a 1280px-wide window even
# after accounting for labels, spacing, and margins (confirmed: two
# rows of 3 capped-170px combos + their labels come to roughly 700-750px
# each, well under 1280px), without needing a dynamic reflow layout for
# what is always exactly 6 fixed filters.
_FILTERS_PER_ROW = 3


class SumDataFilterBar(QWidget):
    """One QComboBox per filterable column, wrapped into rows of
    _FILTERS_PER_ROW via QGridLayout so all six filters stay reachable
    even on a narrow window (see module-level constants above for why
    both the per-combo width cap and the wrap are necessary, confirmed
    by direct measurement, not assumed). Labels are given no fixed
    width of their own -- QLabel's default sizeHint() is content-driven,
    so there's no truncation risk on the label side the way a fixed-
    width QHeaderView column has (confirmed in
    tests/test_sum_data_filtering.py's QFontMetrics check, against the
    real loaded stylesheet, not assumed).
    """

    filters_changed = Signal()

    def __init__(self, source_model, columns: list[tuple[int, str]], parent=None):
        """columns: [(column_index, header_label), ...] -- which Sum
        Data columns to offer a filter for, and what to label each one.
        """
        super().__init__(parent)
        self._column_combos: dict[int, QComboBox] = {}

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 4, 0, 4)
        outer.setSpacing(4)

        filter_caption = QLabel("Filter:")
        filter_caption.setObjectName("hint")
        outer.addWidget(filter_caption)

        grid = QGridLayout()
        grid.setSpacing(10)
        for position, (column, header_label) in enumerate(columns):
            row, pair_col = divmod(position, _FILTERS_PER_ROW)

            combo_label = QLabel(f"{header_label}:")
            grid.addWidget(combo_label, row, pair_col * 2)

            combo = QComboBox()
            combo.setSizeAdjustPolicy(QComboBox.AdjustToMinimumContentsLengthWithIcon)
            combo.setMinimumContentsLength(_COMBO_MIN_CONTENTS_LENGTH)
            combo.setMaximumWidth(_COMBO_MAX_WIDTH)
            combo.setMaxVisibleItems(_COMBO_MAX_VISIBLE_ITEMS)
            combo.addItem(ALL_VALUES_LABEL)
            values = _distinct_column_values(source_model, column)
            combo.addItems(values)
            _widen_popup_to_fit(combo, [ALL_VALUES_LABEL] + values)
            combo.currentIndexChanged.connect(lambda _index: self.filters_changed.emit())
            self._column_combos[column] = combo
            grid.addWidget(combo, row, pair_col * 2 + 1)

        grid.setColumnStretch(_FILTERS_PER_ROW * 2, 1)
        outer.addLayout(grid)

    def current_filters(self) -> dict[int, str]:
        return {
            column: combo.currentText()
            for column, combo in self._column_combos.items()
            if combo.currentText() != ALL_VALUES_LABEL
        }

    def has_active_filters(self) -> bool:
        return bool(self.current_filters())


def _widen_popup_to_fit(combo: QComboBox, items: list[str]) -> None:
    """Widens the OPEN dropdown's popup view (not the closed combo box
    itself, which stays capped at _COMBO_MAX_WIDTH) so every item shows
    in full, never cut off -- confirmed by direct measurement that
    QComboBox's popup otherwise defaults to the closed box's own width,
    which is exactly the reported bug (e.g. "B-C14.1"/"B-ME10" clipped
    in the open list). Verified empirically that combo.view().
    setMinimumWidth() is what actually controls the real popup window's
    width at showPopup() time -- not just the view's own internal
    sizeHint -- before relying on it here.
    """
    if not items:
        return
    metrics = QFontMetrics(combo.font())
    widest = max(items, key=metrics.horizontalAdvance)
    # +40px: scrollbar + the popup item delegate's own left/right
    # padding (QComboBox QAbstractItemView::item in ui/style.qss sets
    # 6px/10px padding) -- a plain text-width measurement alone clips
    # the last character or two without this allowance.
    combo.view().setMinimumWidth(metrics.horizontalAdvance(widest) + 40)


def apply_filtering(
    table,
    model,
    filter_columns: list[tuple[int, str]],
    status_label: QLabel,
) -> tuple[SumDataFilterBar, SumDataFilterProxyModel]:
    """Wires a fresh SumDataFilterBar + SumDataFilterProxyModel onto
    `table` (a FrozenTableView already built against the plain `model`),
    and keeps `status_label` updated with a "Showing X of Y rows" / "no
    filter active" readout. Returns both so the caller can add the
    filter bar to its own layout above the table.

    Built fresh every time the Sum Data tab is (re)built -- which
    DataViewPage already does on every increment/version switch AND
    every time the user switches back to the Sum Data tab at all (see
    that page's _on_tab_changed) -- so filters reset to "All" on any of
    those navigations with no separate reset logic needed; a stale
    filter from a previous increment/version/tab visit can never carry
    over into a freshly rebuilt table.
    """
    proxy = SumDataFilterProxyModel()
    proxy.setSourceModel(model)
    table.setModel(proxy)

    filter_bar = SumDataFilterBar(model, filter_columns)

    total_data_rows = model.rowCount()

    def _refresh():
        proxy.set_filters(filter_bar.current_filters())
        visible_data_rows = sum(
            1 for row in range(model.rowCount()) if proxy.filterAcceptsRow(row, QModelIndex())
        )
        if filter_bar.has_active_filters():
            status_label.setText(f"Filter active — showing {visible_data_rows} of {total_data_rows} rows")
            status_label.setObjectName("filterActiveIndicator")
        else:
            status_label.setText(f"Showing all {total_data_rows} rows")
            status_label.setObjectName("hint")
        status_label.style().unpolish(status_label)
        status_label.style().polish(status_label)

    filter_bar.filters_changed.connect(_refresh)
    _refresh()

    return filter_bar, proxy
