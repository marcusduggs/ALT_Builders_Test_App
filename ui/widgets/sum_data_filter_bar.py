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

*** Totals/Subtotal rows: follow the filter, with the Totals row
pinned visible and relabeled when they do ***
A Subtotal row is shown only if at least one of its own increment's
data rows passes the active filter(s); the overall Totals row is NEVER
hidden (even when literally nothing matches, so its own "0 of 0" can
still be read, rather than vanishing along with every real row). When
ANY filter is active, both kinds of summary row recompute Open/Done/
Total/% Complete from only their OWN scope's CURRENTLY VISIBLE rows --
a Subtotal's scope is its one increment's real rows; the Totals row's
scope is every real row across every increment. With no filter active,
both show exactly the precomputed full-dataset values, byte-for-byte
unchanged from before any filtering existed. The Totals row's own
label gains a "(filtered)" suffix whenever a filter is active, so the
number is never mistaken for the unfiltered total at a glance; a
Subtotal row's own label is left alone (its NUMBERS moving is the
signal there, and relabeling a row that might not even be visible
added no value here).

This supersedes an earlier design (ship with "always full dataset,
clearly labeled" instead) that this project's own author decided
against after live use -- recorded here as settled product behavior,
not as a remaining open question.

*** Filtering never touches the underlying data ***
QSortFilterProxyModel is a read-only VIEW transformation over its
source model -- it can only hide/show/reorder rows, and override what
data() REPORTS, as seen through the proxy; it has no mechanism to
write back to the source model, and this module never calls any
source-model mutator. The recalculated Totals/Subtotal figures
described above are computed fresh on every read inside data(), from
the source model's own already-rendered Status text -- never written
anywhere, including back into the source model itself. Increment.
sum_data and status.json are therefore unreachable from here by
construction, not just by convention.
"""

from __future__ import annotations

from typing import Any, Callable

from PySide6.QtCore import QModelIndex, QSortFilterProxyModel, Qt, Signal
from PySide6.QtGui import QFontMetrics
from PySide6.QtWidgets import QComboBox, QGridLayout, QLabel, QVBoxLayout, QWidget

# Custom item-data roles, all set on column 1 ("Increment"/"Inc #"
# position in every Sum Data table this app builds) of a summary
# (Totals or Subtotal) row. Callers must set all three together, via
# mark_summary_row(), BEFORE wrapping the source model in
# SummaryAwareFilterProxyModel.
SUMMARY_ROW_ROLE = Qt.UserRole + 100
# True only for the one overall Totals row (never for a per-increment
# Subtotal row) -- distinguishes "never hidden, gets the '(filtered)'
# label suffix" from "hidden when its own scope has zero visible rows,
# label never changes".
SUMMARY_IS_TOTALS_ROLE = Qt.UserRole + 101
# list[int] of source row numbers this summary row aggregates -- every
# real data row belonging to one increment, for a Subtotal row; every
# real data row in the whole table, for the Totals row.
SUMMARY_SCOPE_ROWS_ROLE = Qt.UserRole + 102

ALL_VALUES_LABEL = "All"

# Column positions shared by every Sum Data table this module serves
# (DataViewPage and CombinedDataViewPage both use this exact layout --
# see core.excel_export.SUM_DATA_HEADERS, which both preview tables
# mirror) -- used only to recompute a summary row's own numbers from
# its scope rows' Status text, not for the ordinary per-column filter
# matching above (which works on whatever columns the caller offers a
# filter for).
_STATUS_COLUMN = 6
_OPEN_COLUMN = 9
_DONE_COLUMN = 10
_TOTAL_COLUMN = 11
_PCT_COLUMN = 12


class SummaryAwareFilterProxyModel(QSortFilterProxyModel):
    """Hides a data row unless every active per-column filter matches
    it (AND across columns, exact text match per column). A Subtotal
    row (SUMMARY_ROW_ROLE, SUMMARY_IS_TOTALS_ROLE False) is hidden when
    NONE of its own SUMMARY_SCOPE_ROWS_ROLE rows pass that same check;
    the Totals row (SUMMARY_IS_TOTALS_ROLE True) is never hidden.

    data() is also overridden: whenever a filter is active, a summary
    row's Open/Done/Total/% Complete columns (and the Totals row's own
    label) are recomputed from only its scope's currently-visible rows
    instead of passing through the source model's precomputed,
    full-dataset value -- see the module docstring for the exact rule
    and why. See that same docstring for why this can never mutate the
    source model (or anything else) to do so.
    """

    def __init__(self, format_pct: Callable[[int, int], str], parent=None):
        super().__init__(parent)
        self._filters: dict[int, str] = {}
        self._format_pct = format_pct

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
        model = self.sourceModel()
        if model is None:
            return True
        marker_index = model.index(source_row, 1, source_parent)
        if model.data(marker_index, SUMMARY_ROW_ROLE):
            if model.data(marker_index, SUMMARY_IS_TOTALS_ROLE):
                return True
            if not self._filters:
                return True
            scope = model.data(marker_index, SUMMARY_SCOPE_ROWS_ROLE) or []
            return any(self._row_matches_filters(r, source_parent) for r in scope)
        return self._row_matches_filters(source_row, source_parent)

    def _visible_scope_counts(self, scope: list[int]) -> tuple[int, int]:
        model = self.sourceModel()
        open_count = done_count = 0
        for row in scope:
            if not self._row_matches_filters(row):
                continue
            status = model.data(model.index(row, _STATUS_COLUMN))
            if status == "Done":
                done_count += 1
            elif status == "Open":
                open_count += 1
        return open_count, done_count

    def data(self, index, role: int = Qt.DisplayRole):
        if role != Qt.DisplayRole or not self._filters:
            return super().data(index, role)
        model = self.sourceModel()
        source_index = self.mapToSource(index)
        row, column = source_index.row(), source_index.column()
        marker_index = model.index(row, 1)
        if not model.data(marker_index, SUMMARY_ROW_ROLE):
            return super().data(index, role)
        if column not in (1, _OPEN_COLUMN, _DONE_COLUMN, _TOTAL_COLUMN, _PCT_COLUMN):
            return super().data(index, role)

        is_totals = bool(model.data(marker_index, SUMMARY_IS_TOTALS_ROLE))
        if column == 1:
            base_label = super().data(index, role)
            return f"{base_label} (filtered)" if is_totals else base_label

        scope = model.data(marker_index, SUMMARY_SCOPE_ROWS_ROLE) or []
        open_count, done_count = self._visible_scope_counts(scope)
        total = open_count + done_count
        if column == _OPEN_COLUMN:
            return str(open_count)
        if column == _DONE_COLUMN:
            return str(done_count)
        if column == _TOTAL_COLUMN:
            return str(total)
        return self._format_pct(done_count, total)


def _distinct_column_values(model, column: int) -> list[str]:
    """Every distinct non-blank value a column has across the FULL
    (unfiltered) dataset, excluding summary rows -- computed once when
    the filter bar is built, not re-narrowed as other filters change
    (simpler and more predictable than Excel's own cascading dropdowns).
    """
    values: set[str] = set()
    for row in range(model.rowCount()):
        if model.data(model.index(row, 1), SUMMARY_ROW_ROLE):
            continue
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


def mark_summary_row(model, row: int, scope_rows: list[int], is_totals: bool) -> None:
    """Marks `row` (0-indexed) of `model` as a Totals/Subtotal summary
    row, for SummaryAwareFilterProxyModel: `scope_rows` is every real
    (non-summary) row this one aggregates -- one increment's own rows
    for a Subtotal, every real row in the table for the Totals row
    (`is_totals=True`, which also exempts it from ever being hidden and
    gives it the "(filtered)" label suffix once a filter is active).
    Also excluded from _distinct_column_values()'s filter-option lists.
    """
    marker_index = model.index(row, 1)
    model.setData(marker_index, True, SUMMARY_ROW_ROLE)
    model.setData(marker_index, is_totals, SUMMARY_IS_TOTALS_ROLE)
    model.setData(marker_index, list(scope_rows), SUMMARY_SCOPE_ROWS_ROLE)


def apply_filtering(
    table,
    model,
    filter_columns: list[tuple[int, str]],
    status_label: QLabel,
    format_pct: Callable[[int, int], str],
) -> tuple[SumDataFilterBar, SummaryAwareFilterProxyModel]:
    """Wires a fresh SumDataFilterBar + SummaryAwareFilterProxyModel
    onto `table` (a FrozenTableView already built against the plain
    `model`), and keeps `status_label` updated with a "Showing X of Y
    rows" / "no filter active" readout. Returns both so the caller can
    add the filter bar to its own layout above the table.

    `format_pct(done, total)` must be the SAME % Complete formatter
    (zero-total guard included) the caller's own static Totals/Subtotal
    rows already use -- passed in rather than reimplemented here so the
    filtered and unfiltered numbers can never drift to a different
    rounding/zero-division convention from each other.

    Built fresh every time the Sum Data tab is (re)built -- which
    DataViewPage already does on every increment/version switch AND
    every time the user switches back to the Sum Data tab at all (see
    that page's _on_tab_changed) -- so filters reset to "All" on any of
    those navigations with no separate reset logic needed; a stale
    filter from a previous increment/version/tab visit can never carry
    over into a freshly rebuilt table.
    """
    proxy = SummaryAwareFilterProxyModel(format_pct)
    proxy.setSourceModel(model)
    table.setModel(proxy)

    filter_bar = SumDataFilterBar(model, filter_columns)

    total_data_rows = sum(
        1 for row in range(model.rowCount()) if not model.data(model.index(row, 1), SUMMARY_ROW_ROLE)
    )

    def _refresh():
        proxy.set_filters(filter_bar.current_filters())
        visible_data_rows = sum(
            1
            for row in range(model.rowCount())
            if not model.data(model.index(row, 1), SUMMARY_ROW_ROLE) and proxy.filterAcceptsRow(row, QModelIndex())
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
