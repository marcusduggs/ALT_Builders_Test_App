"""
A QTableView with the first N columns pinned to the left edge while the
remaining columns scroll horizontally underneath -- the standard Qt
"frozen column" pattern: a second QTableView sharing the same model
overlays the left edge of the first one and never scrolls horizontally.
"""

from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QAbstractItemView, QTableView


class FrozenTableView(QTableView):
    def __init__(self, frozen_columns: int, parent=None):
        super().__init__(parent)
        self.frozen_columns = frozen_columns
        self.frozen_view = QTableView(self)
        self._setup_frozen_view()

        self.horizontalHeader().sectionResized.connect(self._sync_column_width)
        self.verticalHeader().sectionResized.connect(self._sync_row_height)

        vsb_main = self.verticalScrollBar()
        vsb_frozen = self.frozen_view.verticalScrollBar()
        vsb_main.valueChanged.connect(vsb_frozen.setValue)
        vsb_frozen.valueChanged.connect(vsb_main.setValue)

    def _setup_frozen_view(self):
        self.frozen_view.setFocusPolicy(Qt.NoFocus)
        self.frozen_view.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.frozen_view.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.frozen_view.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.frozen_view.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.frozen_view.setStyleSheet("QTableView { border: none; border-right: 2px solid #ccd0d9; }")
        # The overlay keeps its OWN horizontal header (showing just the
        # frozen columns) so those labels stay pinned too, instead of
        # scrolling away with the main header underneath. Its vertical
        # header is hidden unconditionally -- with it visible, Qt also
        # draws a top-left corner button that has no equivalent in the main
        # view once the main view's vertical header is hidden, which looks
        # like a stray unstyled black box.
        self.frozen_view.verticalHeader().hide()
        self.frozen_view.show()

    def setModel(self, model):
        super().setModel(model)
        self.frozen_view.setModel(model)
        self.frozen_view.setSelectionModel(self.selectionModel())
        self._apply_frozen_column_visibility()
        self._update_frozen_geometry()

    def setAlternatingRowColors(self, enabled):
        # frozen_view is a genuinely SEPARATE QTableView (sharing only the
        # model, not the outer view's own settings) -- without this
        # override, a caller's setAlternatingRowColors(True) on the outer
        # FrozenTableView would zebra-stripe the scrolling columns but
        # leave the pinned (frozen) columns plain white, an obviously
        # mismatched result.
        super().setAlternatingRowColors(enabled)
        self.frozen_view.setAlternatingRowColors(enabled)

    def _apply_frozen_column_visibility(self):
        if self.model() is None:
            return
        for col in range(self.model().columnCount()):
            self.frozen_view.setColumnHidden(col, col >= self.frozen_columns)

    def setColumnWidth(self, column, width):
        super().setColumnWidth(column, width)
        if column < self.frozen_columns:
            self.frozen_view.setColumnWidth(column, width)
        self._update_frozen_geometry()

    def _sync_column_width(self, column, old_size, new_size):
        if column < self.frozen_columns:
            self.frozen_view.setColumnWidth(column, new_size)
        self._update_frozen_geometry()

    def _sync_row_height(self, row, old_size, new_size):
        self.frozen_view.setRowHeight(row, new_size)

    def _update_frozen_geometry(self):
        if self.model() is None:
            return
        total_width = sum(self.columnWidth(c) for c in range(self.frozen_columns))
        header_height = self.horizontalHeader().height() if not self.horizontalHeader().isHidden() else 0
        # Starts at the very top (y = frameWidth), covering its own header
        # AND the data rows, so the frozen columns' header labels stay
        # pinned in place the same way the data cells do.
        self.frozen_view.setGeometry(
            self.frameWidth(),
            self.frameWidth(),
            total_width,
            header_height + self.viewport().height(),
        )

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._update_frozen_geometry()

    def moveCursor(self, cursor_action, modifiers):
        # Keep keyboard navigation from scrolling frozen columns out from
        # under the overlay -- same fix used in Qt's own frozen-column
        # example.
        current = super().moveCursor(cursor_action, modifiers)
        if (
            cursor_action == QAbstractItemView.MoveLeft
            and current.column() >= self.frozen_columns
            and self.visualRect(current).topLeft().x() < self.frozen_view.columnWidth(0)
        ):
            new_value = self.horizontalScrollBar().value() + self.visualRect(current).topLeft().x() - self.frozen_view.columnWidth(0)
            self.horizontalScrollBar().setValue(new_value)
        return current
