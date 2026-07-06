"""
Data view: All Data / Sum Data / Report for one increment, as tabs over
the SAME cached parse -- see ui.mock_data.MockDataStore.get_increment_for_display,
which loads all three together in one background pass. Switching tabs
never re-parses the workbook or shows a progress dialog; it just
re-renders from the Increment already sitting in memory (self.increment).

All Data tab: the first three columns (Index, Description, Approval
Agency) pinned while the 42 stage columns + VCR + SUM scroll underneath.
Required-stage cells (a real X/1 marker in the source file, see
core.excel_reader.required_stages_by_index) are interactive: click to
toggle Open/Done, saved immediately to status.json via
ui.mock_data.MockDataStore.set_stage_status -- no separate Save step, and
no re-parsing the workbook on every click (the in-memory Increment is
updated directly; set_stage_status is a cheap, instant local write).
Non-required cells remain exactly as before: blank, non-interactive.

Sum Data tab: same frozen-column layout, read-only. Each required stage
shows the LIVE status from the same in-memory stage_status dict the All
Data tab edits (ui.mock_data.MockDataStore's Increment.sum_data rows
literally share that dict object -- see _sum_data_rows() there), defaulting
to "Open" when unset (this tab's "current state" framing, distinct from
All Data's "needs status" tracking concept) -- plus a live % Complete
column. Rebuilt every time this tab becomes current (see
_on_tab_changed), so an edit made on All Data always shows up here
immediately without a background reload.

Report tab: a plain (non-frozen) table -- Approval Agency / Index /
Description / Total, grouped and Grand-Totaled exactly as
core.excel_reader.build_report() computes it. Static: doesn't depend on
status.json, so it's built once.

"Export to Excel" (top right of the header) writes all three tabs -- the
same cached Increment every tab renders from -- to one .xlsx via
core.excel_export.
"""

from __future__ import annotations

from PySide6.QtCore import QModelIndex, Qt, QTimer
from PySide6.QtGui import QColor, QFont, QIcon, QPainter, QPixmap, QStandardItem, QStandardItemModel
from PySide6.QtWidgets import (
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from core import excel_export
from ui.mock_data import STAGE_COUNT, Increment, MockDataStore
from ui.widgets.frozen_table import FrozenTableView
from ui.workers import run_with_progress

FROZEN_COLUMNS = 3
RECENTLY_ADDED_COLOR = QColor("#eaf7ee")
FLASH_COLOR = QColor("#fff3b0")

_STAGE_STATUS_COLORS = {
    "Done": QColor("#2e7d32"),
    "Open": QColor("#c62828"),
}
_STAGE_STATUS_LABELS = {
    "Done": "X",
    "Open": "1",
}


def _first_line(text) -> str:
    if not text:
        return ""
    return str(text).split("\n", 1)[0]


def _format_total(value) -> str:
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def _format_pct(done: int, total: int) -> str:
    if total == 0:
        return "—"
    return f"{round(100 * done / total)}%"


def _needs_status_icon() -> QIcon:
    size = 14
    pixmap = QPixmap(size, size)
    pixmap.fill(Qt.transparent)
    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.Antialiasing)
    painter.setBrush(QColor("#ffb703"))
    painter.setPen(Qt.NoPen)
    painter.drawEllipse(0, 0, size, size)
    painter.setPen(QColor("#3a2900"))
    font = QFont()
    font.setBold(True)
    font.setPointSize(9)
    painter.setFont(font)
    painter.drawText(pixmap.rect(), Qt.AlignCenter, "!")
    painter.end()
    return QIcon(pixmap)


def _stage_status_icon(status: str | None) -> QIcon:
    """A small colored badge -- "X" on green for Done, "1" on red for
    Open -- mirroring the same marks Rey's team uses in the source
    Excel file (see reys_fill_color_convention in project memory), so
    the in-app status reads the same way the paperwork does. A required
    stage that has never had a status set renders as a completely blank
    icon -- no ring, no placeholder mark -- same as a cell that isn't
    required at all; "needs status" is still surfaced at the row level
    (the "!" badge -- see _apply_row_badge) and in the footer counts, not
    on every individual cell. (The Sum Data tab never passes None here --
    it defaults an unset required stage to "Open" itself, see
    _build_sum_data_table.)
    """
    color = _STAGE_STATUS_COLORS.get(status)
    if color is None:
        return QIcon()

    size = 16
    pixmap = QPixmap(size, size)
    pixmap.fill(Qt.transparent)
    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.Antialiasing)
    painter.setBrush(color)
    painter.setPen(Qt.NoPen)
    painter.drawRoundedRect(1, 1, size - 2, size - 2, 3, 3)
    painter.setPen(QColor("#ffffff"))
    font = QFont()
    font.setBold(True)
    font.setPointSize(9)
    painter.setFont(font)
    painter.drawText(pixmap.rect(), Qt.AlignCenter, _STAGE_STATUS_LABELS[status])
    painter.end()
    return QIcon(pixmap)


def _stage_tooltip(stage: int, status: str | None) -> str:
    if status is None:
        return f"Stage {stage} — needs status. Click to mark Open."
    return f"Stage {stage} — {status}. Click to toggle."


class DataViewPage(QWidget):
    def __init__(self, project_name: str, increment: Increment, store: MockDataStore, on_back, parent=None):
        super().__init__(parent)
        self.project_name = project_name
        self.increment = increment
        self.store = store
        self.on_back = on_back

        outer = QVBoxLayout(self)
        outer.setContentsMargins(20, 20, 20, 20)
        outer.setSpacing(12)

        self.tabs = QTabWidget()
        self.tabs.addTab(self._build_all_data_tab(), "All Data")

        self.sum_data_container = QWidget()
        self.sum_data_layout = QVBoxLayout(self.sum_data_container)
        self.sum_data_layout.setContentsMargins(0, 0, 0, 0)
        self._refresh_sum_data_tab()
        self.tabs.addTab(self.sum_data_container, "Sum Data")

        self.tabs.addTab(self._build_report_tab(), "Report")

        self.tabs.currentChanged.connect(self._on_tab_changed)

        outer.addWidget(self._build_header())
        outer.addWidget(self.tabs, stretch=1)
        outer.addWidget(self._build_footer())

    def _on_tab_changed(self, index: int):
        # Re-render Sum Data from the cached Increment every time it
        # becomes current -- picks up any (index, stage) edits made on
        # the All Data tab since the last time this tab was shown,
        # without ever re-parsing the workbook or touching a background
        # thread (see module docstring).
        if self.tabs.widget(index) is self.sum_data_container:
            self._refresh_sum_data_tab()

    # ------------------------------------------------------------------
    # header + export
    # ------------------------------------------------------------------
    def _build_header(self) -> QWidget:
        row = QWidget()
        layout = QHBoxLayout(row)
        layout.setContentsMargins(0, 0, 0, 0)

        back_button = QPushButton("< Back to Increments")
        back_button.clicked.connect(lambda: self.on_back())

        title = QLabel(f"{self.increment.name} — Version {self.increment.version}")
        title.setObjectName("pageTitle")

        subtitle = QLabel(f"{self.project_name}  ·  Last updated {self.increment.last_updated}")
        subtitle.setObjectName("hint")

        title_box = QVBoxLayout()
        title_box.setSpacing(2)
        title_box.addWidget(title)
        title_box.addWidget(subtitle)

        export_button = QPushButton("Export to Excel")
        export_button.clicked.connect(self._on_export_clicked)

        layout.addWidget(back_button)
        layout.addSpacing(16)
        layout.addLayout(title_box)
        layout.addStretch(1)
        layout.addWidget(export_button)
        return row

    def _on_export_clicked(self):
        default_name = excel_export.default_filename(self.project_name, self.increment.name, self.increment.version)
        path, _ = QFileDialog.getSaveFileName(self, "Export to Excel", default_name, "Excel Workbook (*.xlsx)")
        if not path:
            return  # user cancelled -- nothing changes

        def on_finished(_result):
            QMessageBox.information(self, "Export Complete", f"Saved to:\n\n{path}")

        def on_error(exc):
            QMessageBox.critical(self, "Could Not Export", f"This file couldn't be saved:\n\n{exc}")

        run_with_progress(
            self, "Exporting to Excel...", excel_export.export_increment,
            on_finished, on_error, self.increment, path,
        )

    # ------------------------------------------------------------------
    # All Data tab
    # ------------------------------------------------------------------
    def _build_all_data_tab(self) -> FrozenTableView:
        headers = ["Index", "Description", "Approval Agency"]
        headers += [f"Stage {i}" for i in range(1, STAGE_COUNT + 1)]
        headers += ["VCR", "SUM"]

        model = QStandardItemModel(len(self.increment.all_data), len(headers))
        model.setHorizontalHeaderLabels(headers)

        for row_idx, row_data in enumerate(self.increment.all_data):
            is_recently_added = bool(row_data.get("recently_added"))

            index_item = QStandardItem(str(row_data["index"]))
            self._apply_row_badge(index_item, row_data)

            description_item = QStandardItem(_first_line(row_data.get("description")))
            description_item.setToolTip(row_data.get("description") or "")

            agency_item = QStandardItem(row_data.get("approval_agency") or "")

            items = [index_item, description_item, agency_item]

            required = set(row_data.get("required_stages", []))
            stage_status = row_data.get("stage_status", {})
            for stage in range(1, STAGE_COUNT + 1):
                stage_item = QStandardItem("")
                stage_item.setTextAlignment(Qt.AlignCenter)
                if stage in required:
                    status = stage_status.get(stage)
                    stage_item.setIcon(_stage_status_icon(status))
                    stage_item.setToolTip(_stage_tooltip(stage, status))
                items.append(stage_item)

            vcr_item = QStandardItem("" if row_data.get("vcr") is None else str(row_data["vcr"]))
            vcr_item.setTextAlignment(Qt.AlignCenter)
            items.append(vcr_item)

            sum_item = QStandardItem(str(row_data.get("sum", 0)))
            sum_item.setTextAlignment(Qt.AlignCenter)
            items.append(sum_item)

            for item in items:
                item.setEditable(False)
                if is_recently_added:
                    item.setBackground(RECENTLY_ADDED_COLOR)
            for col_idx, item in enumerate(items):
                model.setItem(row_idx, col_idx, item)

        table = FrozenTableView(frozen_columns=FROZEN_COLUMNS)
        table.setModel(model)
        table.setAlternatingRowColors(False)
        table.verticalHeader().hide()  # Index column already identifies each row
        table.setColumnWidth(0, 90)
        table.setColumnWidth(1, 260)
        table.setColumnWidth(2, 140)
        for col in range(FROZEN_COLUMNS, FROZEN_COLUMNS + STAGE_COUNT):
            table.setColumnWidth(col, 56)
        table.setColumnWidth(FROZEN_COLUMNS + STAGE_COUNT, 56)
        table.setColumnWidth(FROZEN_COLUMNS + STAGE_COUNT + 1, 56)
        table.horizontalHeader().setMinimumSectionSize(40)
        table.clicked.connect(self._on_cell_clicked)
        self.table = table
        return table

    # ------------------------------------------------------------------
    # status-setting interaction (All Data tab only -- Sum Data is read-only)
    # ------------------------------------------------------------------
    def _on_cell_clicked(self, model_index: QModelIndex):
        row, col = model_index.row(), model_index.column()
        stage = col - FROZEN_COLUMNS + 1
        if not (1 <= stage <= STAGE_COUNT):
            return  # Index/Description/Agency/VCR/SUM -- not a stage cell

        row_data = self.increment.all_data[row]
        required = row_data.get("required_stages", [])
        if stage not in required:
            return  # not required for this item -- stays blank, non-interactive

        stage_status = row_data.setdefault("stage_status", {})
        current = stage_status.get(stage)
        new_status = "Open" if current in (None, "Done") else "Done"

        # Persist immediately -- cheap local JSON write, no "Save" button.
        self.store.set_stage_status(self.project_name, self.increment.name, row_data["index"], stage, new_status)

        # Update in-memory state and just the affected cells -- no
        # workbook re-parse, no full-table rebuild. Increment.sum_data
        # rows share this exact stage_status dict object (see
        # ui.mock_data._sum_data_rows), so this write is already visible
        # there too the next time that tab is rendered.
        stage_status[stage] = new_status
        row_data["needs_status_stages"] = [s for s in required if s not in stage_status]
        row_data["needs_status"] = bool(row_data["needs_status_stages"])

        model = self.table.model()
        stage_item = model.item(row, col)
        stage_item.setIcon(_stage_status_icon(new_status))
        stage_item.setToolTip(_stage_tooltip(stage, new_status))
        self._flash(stage_item, self._resting_background(row_data))

        index_item = model.item(row, 0)
        self._apply_row_badge(index_item, row_data)

        self._refresh_footer()

    def _resting_background(self, row_data: dict) -> QColor:
        return RECENTLY_ADDED_COLOR if row_data.get("recently_added") else QColor(Qt.transparent)

    def _flash(self, item: QStandardItem, resting_color: QColor):
        """Brief color flash so a click's effect is unmistakable, even
        though the icon change alone is already visible.

        Reverts to an explicitly-computed resting color rather than
        whatever item.background() reads back at call time -- clicking
        the same cell twice within the flash window would otherwise
        capture the *still-flashing* color as "original" and get stuck
        showing it forever once that second timer fires.
        """
        item.setBackground(FLASH_COLOR)
        QTimer.singleShot(350, lambda: item.setBackground(resting_color))

    def _apply_row_badge(self, index_item: QStandardItem, row_data: dict):
        if row_data.get("needs_status"):
            index_item.setIcon(_needs_status_icon())
            stages = ", ".join(str(s) for s in row_data.get("needs_status_stages", []))
            index_item.setToolTip(f"Needs status for stage(s): {stages}")
        else:
            index_item.setIcon(QIcon())
            index_item.setToolTip("")

    # ------------------------------------------------------------------
    # Sum Data tab
    # ------------------------------------------------------------------
    def _refresh_sum_data_tab(self):
        while self.sum_data_layout.count():
            child = self.sum_data_layout.takeAt(0)
            if child.widget():
                child.widget().deleteLater()
        self.sum_data_layout.addWidget(self._build_sum_data_table())

    def _build_sum_data_table(self) -> FrozenTableView:
        headers = ["Index", "Description", "Approval Agency"]
        headers += [f"Stage {i}" for i in range(1, STAGE_COUNT + 1)]
        headers += ["VCR", "% Complete"]

        rows = self.increment.sum_data
        model = QStandardItemModel(len(rows), len(headers))
        model.setHorizontalHeaderLabels(headers)

        for row_idx, row_data in enumerate(rows):
            index_item = QStandardItem(str(row_data["index"]))
            description_item = QStandardItem(_first_line(row_data.get("description")))
            description_item.setToolTip(row_data.get("description") or "")
            agency_item = QStandardItem(row_data.get("approval_agency") or "")

            items = [index_item, description_item, agency_item]

            required = row_data.get("required_stages", [])
            stage_status = row_data.get("stage_status", {})
            done_count = 0
            for stage in range(1, STAGE_COUNT + 1):
                stage_item = QStandardItem("")
                stage_item.setTextAlignment(Qt.AlignCenter)
                if stage in required:
                    # Live current state, not "needs status" -- an unset
                    # required stage defaults to Open here (see module
                    # docstring's Sum Data section).
                    status = stage_status.get(stage, "Open")
                    if status == "Done":
                        done_count += 1
                    stage_item.setIcon(_stage_status_icon(status))
                    stage_item.setToolTip(f"Stage {stage} — {status}")
                items.append(stage_item)

            vcr_item = QStandardItem("" if row_data.get("vcr") is None else str(row_data["vcr"]))
            vcr_item.setTextAlignment(Qt.AlignCenter)
            items.append(vcr_item)

            pct_item = QStandardItem(_format_pct(done_count, len(required)))
            pct_item.setTextAlignment(Qt.AlignCenter)
            items.append(pct_item)

            for item in items:
                item.setEditable(False)
            for col_idx, item in enumerate(items):
                model.setItem(row_idx, col_idx, item)

        table = FrozenTableView(frozen_columns=FROZEN_COLUMNS)
        table.setModel(model)
        table.setAlternatingRowColors(False)
        table.verticalHeader().hide()
        table.setColumnWidth(0, 90)
        table.setColumnWidth(1, 260)
        table.setColumnWidth(2, 140)
        for col in range(FROZEN_COLUMNS, FROZEN_COLUMNS + STAGE_COUNT):
            table.setColumnWidth(col, 56)
        table.setColumnWidth(FROZEN_COLUMNS + STAGE_COUNT, 56)
        table.setColumnWidth(FROZEN_COLUMNS + STAGE_COUNT + 1, 80)
        table.horizontalHeader().setMinimumSectionSize(40)
        # Read-only view of the live status -- no click handler; the All
        # Data tab is the sole editing surface (see module docstring).
        return table

    # ------------------------------------------------------------------
    # Report tab
    # ------------------------------------------------------------------
    def _build_report_tab(self) -> QTableWidget:
        headers = ["Approval Agency", "Index", "Description", "Total"]
        rows = self.increment.report

        table = QTableWidget(len(rows), len(headers))
        table.setHorizontalHeaderLabels(headers)
        table.verticalHeader().hide()
        table.setEditTriggers(QTableWidget.NoEditTriggers)
        table.setAlternatingRowColors(False)
        table.setSelectionBehavior(QTableWidget.SelectRows)

        bold_font = QFont()
        bold_font.setBold(True)

        for row_idx, row_data in enumerate(rows):
            is_grand_total = row_data.get("approval_agency") == "Grand Total"
            values = [
                row_data.get("approval_agency") or "",
                row_data.get("index") or "",
                _first_line(row_data.get("description")),
                _format_total(row_data.get("total", 0)),
            ]
            for col_idx, value in enumerate(values):
                item = QTableWidgetItem(str(value))
                if col_idx == 3:
                    item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                if is_grand_total:
                    item.setFont(bold_font)
                table.setItem(row_idx, col_idx, item)

        table.setColumnWidth(0, 220)
        table.setColumnWidth(1, 90)
        table.setColumnWidth(2, 320)
        table.setColumnWidth(3, 100)
        table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        return table

    # ------------------------------------------------------------------
    # footer
    # ------------------------------------------------------------------
    def _build_footer(self) -> QFrame:
        frame = QFrame()
        frame.setObjectName("footerBar")
        layout = QHBoxLayout(frame)
        layout.setContentsMargins(14, 4, 14, 4)
        layout.setSpacing(24)

        self.total_label = QLabel()
        self.needs_status_label = QLabel()
        self.changed_label = QLabel()
        self.value_changed_label = QLabel()

        layout.addWidget(self.total_label)
        layout.addWidget(self.needs_status_label)
        layout.addWidget(self.changed_label)
        layout.addWidget(self.value_changed_label)
        layout.addStretch(1)

        self.footer_frame = frame
        self._refresh_footer()
        return frame

    def _refresh_footer(self):
        total = self.increment.total_count
        stage_count = self.increment.needs_status_stage_count
        row_count = self.increment.needs_status_row_count
        changed = self.increment.changed_count

        self.total_label.setText(f"Total items: {total}")

        if stage_count:
            stage_word = "stage" if stage_count == 1 else "stages"
            item_word = "item" if row_count == 1 else "items"
            self.needs_status_label.setText(
                f"Needing status: {stage_count} {stage_word} across {row_count} {item_word}"
            )
        else:
            self.needs_status_label.setText("Needing status: none")
        self._set_highlighted(self.needs_status_label, bool(stage_count))

        self.changed_label.setText(f"New items: {changed}")
        self._set_highlighted(self.changed_label, bool(changed))

        value_changed = self.increment.value_changed_count
        self.value_changed_label.setText(f"Values changed: {value_changed}")
        self._set_highlighted(self.value_changed_label, bool(value_changed))

    def _set_highlighted(self, label: QLabel, highlighted: bool):
        label.setObjectName("footerHighlight" if highlighted else "")
        # QSS object-name selectors only re-apply after an explicit
        # unpolish/polish -- needed here because these labels are updated
        # in place (see _on_cell_clicked) rather than rebuilt from scratch.
        label.style().unpolish(label)
        label.style().polish(label)
