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
core.excel_export. Respects whichever version is currently selected
(see below), since it always exports self.increment, not "the latest".

Version selector (next to the title): a dropdown of every stored
version ("v3 (2026-07-06)", newest first), defaulting to the latest.
Switching versions re-parses THAT version's file in the background
(same run_with_progress spinner pattern as the initial load -- this is
real, multi-second openpyxl parsing, not a cheap cache lookup) and
replaces self.increment wholesale, which is why All Data/Sum Data/
Report are each built inside their own container+layout (mirroring the
Sum Data tab's existing rebuild-in-place pattern) rather than built
once in __init__ like before. Status marks are NOT tracked historically
per version -- they always come from the current status.json (see
ui.mock_data.MockDataStore.get_increment_for_display) -- so viewing
anything other than the latest version shows a persistent warning
banner clarifying that status reflects today's progress, not a
historical snapshot.
"""

from __future__ import annotations

from PySide6.QtCore import QModelIndex, Qt, QTimer
from PySide6.QtGui import QColor, QFont, QIcon, QPainter, QPixmap, QStandardItem, QStandardItemModel
from PySide6.QtWidgets import (
    QComboBox,
    QFileDialog,
    QFrame,
    QHBoxLayout,
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
from core.excel_reader import live_sum_data_totals
from ui.mock_data import STAGE_COUNT, Increment, MockDataStore
from ui.widgets.frozen_table import FrozenTableView
from ui.workers import run_with_progress

FROZEN_COLUMNS = 3
RECENTLY_ADDED_COLOR = QColor("#eaf7ee")
FLASH_COLOR = QColor("#fff3b0")
TOTALS_ROW_BACKGROUND = QColor("#e4e8f0")

_STAGE_STATUS_COLORS = {
    "Done": QColor("#2e7d32"),
    "Open": QColor("#c62828"),
}
_STAGE_STATUS_LABELS = {
    "Done": "X",
    "Open": "1",
}

# Sum Data's own required-stage cell style -- literal "Done"/"Open" text
# on a full-cell fill, matching standard Excel conditional-formatting
# "Good"/"Bad" cell styles (fill + text colors below are those exact
# built-in Excel values), NOT the small icon-badge look All Data uses
# (_stage_status_icon) -- a deliberate visual distinction between the
# two tabs, per how Sum Data is meant to read like a status report.
_SUM_DATA_STATUS_FILL = {
    "Done": QColor("#C6EFCE"),
    "Open": QColor("#FFC7CE"),
}
_SUM_DATA_STATUS_TEXT_COLOR = {
    "Done": QColor("#006100"),
    "Open": QColor("#9C0006"),
}


def _format_total(value) -> str:
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def _format_pct(done: int, total: int) -> str:
    if total == 0:
        return "—"
    return f"{round(100 * done / total)}%"


def _totals_item(text: str) -> QStandardItem:
    """One cell of the bottom totals row -- bold + tinted background so
    it reads as a summary at a glance, not just another item; not
    editable, and (since it's never wired to any click handler) not
    interactive either.
    """
    item = QStandardItem(text)
    item.setTextAlignment(Qt.AlignCenter)
    item.setEditable(False)
    item.setBackground(TOTALS_ROW_BACKGROUND)
    font = item.font()
    font.setBold(True)
    item.setFont(font)
    return item


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
        self.latest_version = increment.version  # refined by _populate_version_combo below

        outer = QVBoxLayout(self)
        outer.setContentsMargins(20, 20, 20, 20)
        outer.setSpacing(12)

        outer.addWidget(self._build_header())
        outer.addWidget(self._build_version_notice())

        self.all_data_container = QWidget()
        self.all_data_layout = QVBoxLayout(self.all_data_container)
        self.all_data_layout.setContentsMargins(0, 0, 0, 0)
        self._refresh_all_data_tab()

        self.sum_data_container = QWidget()
        self.sum_data_layout = QVBoxLayout(self.sum_data_container)
        self.sum_data_layout.setContentsMargins(0, 0, 0, 0)
        self._refresh_sum_data_tab()

        self.report_container = QWidget()
        self.report_layout = QVBoxLayout(self.report_container)
        self.report_layout.setContentsMargins(0, 0, 0, 0)
        self._refresh_report_tab()

        self.tabs = QTabWidget()
        self.tabs.addTab(self.all_data_container, "All Data")
        self.tabs.addTab(self.sum_data_container, "Sum Data")
        self.tabs.addTab(self.report_container, "Report")
        self.tabs.currentChanged.connect(self._on_tab_changed)

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
    # header + version selector + export
    # ------------------------------------------------------------------
    def _build_header(self) -> QWidget:
        row = QWidget()
        layout = QHBoxLayout(row)
        layout.setContentsMargins(0, 0, 0, 0)

        back_button = QPushButton("< Back to Increments")
        back_button.clicked.connect(lambda: self.on_back())

        self.title_label = QLabel(f"{self.increment.name} — Version {self.increment.version}")
        self.title_label.setObjectName("pageTitle")

        self.subtitle_label = QLabel(f"{self.project_name}  ·  Last updated {self.increment.last_updated}")
        self.subtitle_label.setObjectName("hint")

        title_box = QVBoxLayout()
        title_box.setSpacing(2)
        title_box.addWidget(self.title_label)
        title_box.addWidget(self.subtitle_label)

        version_label = QLabel("Version:")
        self.version_combo = QComboBox()
        self.version_combo.setMinimumWidth(150)
        self._populate_version_combo()
        self.version_combo.currentIndexChanged.connect(self._on_version_changed)

        export_button = QPushButton("Export to Excel")
        export_button.clicked.connect(self._on_export_clicked)

        layout.addWidget(back_button)
        layout.addSpacing(16)
        layout.addLayout(title_box)
        layout.addSpacing(16)
        layout.addWidget(version_label)
        layout.addWidget(self.version_combo)
        layout.addStretch(1)
        layout.addWidget(export_button)
        return row

    def _populate_version_combo(self):
        """Every stored version of this increment, newest first (see
        core.project_store.list_versions) -- fetched once here, not
        re-fetched on every switch, since the set of stored versions
        doesn't change while this page is open.
        """
        version_records = self.store.list_versions(self.project_name, self.increment.name)
        self.latest_version = version_records[0].version if version_records else self.increment.version

        self.version_combo.blockSignals(True)
        self.version_combo.clear()
        for v in version_records:
            self.version_combo.addItem(f"v{v.version} ({v.uploaded_date})", userData=v.version)
        self._select_version_in_combo(self.increment.version)
        self.version_combo.blockSignals(False)

    def _select_version_in_combo(self, version: int):
        idx = self.version_combo.findData(version)
        if idx >= 0:
            self.version_combo.setCurrentIndex(idx)

    def _on_version_changed(self, index: int):
        if index < 0:
            return
        selected_version = self.version_combo.itemData(index)
        if selected_version == self.increment.version:
            return  # no real change -- combo was just re-synced programmatically

        def on_finished(new_increment):
            if new_increment is None:
                QMessageBox.critical(
                    self, "Could Not Load Version", f"Version {selected_version} could not be found."
                )
                self._revert_version_combo()
                return
            self._apply_new_increment(new_increment)

        def on_error(exc):
            QMessageBox.critical(self, "Could Not Load Version", f"This version couldn't be read:\n\n{exc}")
            self._revert_version_combo()

        run_with_progress(
            self, f"Loading version {selected_version}...", self.store.get_increment_for_display,
            on_finished, on_error, self.project_name, self.increment.name, selected_version,
        )

    def _revert_version_combo(self):
        """On a failed version load, the combo's selection has already
        visually moved to the version that FAILED to load -- snap it
        back to whatever's actually displayed so the dropdown never
        shows a version that isn't the one on screen.
        """
        self.version_combo.blockSignals(True)
        self._select_version_in_combo(self.increment.version)
        self.version_combo.blockSignals(False)

    def _apply_new_increment(self, new_increment: Increment):
        self.increment = new_increment
        self._refresh_all_data_tab()
        self._refresh_sum_data_tab()
        self._refresh_report_tab()
        self._refresh_footer()
        self.title_label.setText(f"{self.increment.name} — Version {self.increment.version}")
        self.subtitle_label.setText(f"{self.project_name}  ·  Last updated {self.increment.last_updated}")
        self._update_version_notice()

    def _build_version_notice(self) -> QFrame:
        """Hidden whenever the LATEST version is displayed; shown with a
        clear explanation otherwise -- status marks are never tracked
        historically per version (see ui.mock_data.MockDataStore.
        get_increment_for_display), so an older version's Done/Open marks
        are today's live progress overlaid on that version's data, not a
        preserved historical snapshot. Reuses the same warning styling
        core/../review_dialog.py's Column Anomalies section uses, for a
        consistent "pay attention" visual language across the app.
        """
        frame = QFrame()
        frame.setObjectName("changeSectionWarning")
        layout = QHBoxLayout(frame)
        layout.setContentsMargins(12, 8, 12, 8)
        self.version_notice_label = QLabel()
        self.version_notice_label.setObjectName("changeSectionTitleWarning")
        self.version_notice_label.setWordWrap(True)
        layout.addWidget(self.version_notice_label)
        self.version_notice_frame = frame
        self._update_version_notice()
        return frame

    def _update_version_notice(self):
        is_latest = self.increment.version == self.latest_version
        self.version_notice_frame.setVisible(not is_latest)
        if not is_latest:
            self.version_notice_label.setText(
                f"Viewing v{self.increment.version} ({self.increment.last_updated}) — status marks reflect "
                "current progress, not historical status."
            )

    def _on_export_clicked(self):
        default_name = excel_export.default_filename(self.project_name, self.increment.name, self.increment.version)
        path, _ = QFileDialog.getSaveFileName(self, "Export to Excel", default_name, "Excel Workbook (*.xlsx)")
        if not path:
            return  # user cancelled -- nothing changes

        def on_finished(_result):
            QMessageBox.information(self, "Export Complete", f"Saved to:\n\n{path}")

        def on_error(exc):
            QMessageBox.critical(self, "Could Not Export", f"This file couldn't be saved:\n\n{exc}")

        # Exports whatever self.increment currently holds -- the
        # currently SELECTED version, not necessarily the latest (see
        # _apply_new_increment) -- and default_filename() already bakes
        # self.increment.version into the filename, so an export made
        # while viewing an older version is never silently mistaken for
        # the latest one.
        run_with_progress(
            self, "Exporting to Excel...", excel_export.export_increment,
            on_finished, on_error, self.increment, path,
        )

    # ------------------------------------------------------------------
    # All Data tab
    # ------------------------------------------------------------------
    def _refresh_all_data_tab(self):
        while self.all_data_layout.count():
            child = self.all_data_layout.takeAt(0)
            if child.widget():
                child.widget().deleteLater()
        self.all_data_layout.addWidget(self._build_all_data_tab())

    def _build_all_data_tab(self) -> FrozenTableView:
        headers = ["Index", "Description", "Approval Agency"]
        headers += [f"Stage {i}" for i in range(1, STAGE_COUNT + 1)]
        headers += ["VCR", "SUM"]

        item_count = len(self.increment.all_data)
        model = QStandardItemModel(item_count + 1, len(headers))  # +1 for the totals row
        model.setHorizontalHeaderLabels(headers)

        for row_idx, row_data in enumerate(self.increment.all_data):
            is_recently_added = bool(row_data.get("recently_added"))

            index_item = QStandardItem(str(row_data["index"]))
            self._apply_row_badge(index_item, row_data)

            # Full 3-line description (component / code citation / test
            # name), not just the first line -- wrapped across multiple
            # visual lines via resizeRowsToContents() below, matching how
            # the exported .xlsx already shows it (core.excel_export's
            # _apply_description_wrap). Tooltip kept as a fallback for
            # any line long enough to still get clipped at this column
            # width.
            description_item = QStandardItem(row_data.get("description") or "")
            description_item.setToolTip(row_data.get("description") or "")

            agency_item = QStandardItem(row_data.get("approval_agency") or "")

            items = [index_item, description_item, agency_item]

            required = set(row_data.get("required_stages", []))
            stage_status = row_data.get("stage_status", {})
            for stage in range(1, STAGE_COUNT + 1):
                if stage in required:
                    stage_item = QStandardItem("")
                    status = stage_status.get(stage)
                    stage_item.setIcon(_stage_status_icon(status))
                    stage_item.setToolTip(_stage_tooltip(stage, status))
                else:
                    # Matches the source file's own convention (and
                    # core.excel_reader.build_all_data(), which already
                    # stores 0 here) -- a genuine "0", not blank.
                    stage_item = QStandardItem("0")
                stage_item.setTextAlignment(Qt.AlignCenter)
                items.append(stage_item)

            # Same "0, not blank" convention as non-required stage cells
            # above -- matches the source file's own convention.
            vcr_item = QStandardItem("0" if row_data.get("vcr") is None else str(row_data["vcr"]))
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

        totals = self.increment.all_data_totals
        totals_items = [_totals_item("Totals"), _totals_item(""), _totals_item("")]
        for stage in range(1, STAGE_COUNT + 1):
            totals_items.append(_totals_item(_format_total(totals.get(f"Stage {stage}", 0))))
        totals_items.append(_totals_item(_format_total(totals.get("VCR", 0))))
        totals_items.append(_totals_item(_format_total(totals.get("SUM", 0))))
        for col_idx, item in enumerate(totals_items):
            model.setItem(item_count, col_idx, item)

        table = FrozenTableView(frozen_columns=FROZEN_COLUMNS)
        table.setModel(model)
        table.setAlternatingRowColors(False)
        table.setWordWrap(True)
        table.verticalHeader().hide()  # Index column already identifies each row
        table.setColumnWidth(0, 90)
        table.setColumnWidth(1, 320)
        table.setColumnWidth(2, 140)
        for col in range(FROZEN_COLUMNS, FROZEN_COLUMNS + STAGE_COUNT):
            table.setColumnWidth(col, 56)
        table.setColumnWidth(FROZEN_COLUMNS + STAGE_COUNT, 56)
        table.setColumnWidth(FROZEN_COLUMNS + STAGE_COUNT + 1, 56)
        table.horizontalHeader().setMinimumSectionSize(40)
        # Each row's height sized to fit its own wrapped Description text
        # (some are 1 line, some wrap to several) -- must run AFTER
        # column widths are set, since wrapping depends on them.
        table.resizeRowsToContents()
        table.clicked.connect(self._on_cell_clicked)
        self.table = table
        return table

    # ------------------------------------------------------------------
    # status-setting interaction (All Data tab only -- Sum Data is read-only)
    # ------------------------------------------------------------------
    def _on_cell_clicked(self, model_index: QModelIndex):
        row, col = model_index.row(), model_index.column()
        if row >= len(self.increment.all_data):
            return  # the bottom totals row -- a summary, not a real item

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
        headers += ["VCR", "Open", "Done", "Total", "% Complete"]

        rows = self.increment.sum_data
        model = QStandardItemModel(len(rows) + 1, len(headers))  # +1 for the totals row
        model.setHorizontalHeaderLabels(headers)

        for row_idx, row_data in enumerate(rows):
            index_item = QStandardItem(str(row_data["index"]))
            description_item = QStandardItem(row_data.get("description") or "")
            description_item.setToolTip(row_data.get("description") or "")
            agency_item = QStandardItem(row_data.get("approval_agency") or "")

            items = [index_item, description_item, agency_item]

            required = row_data.get("required_stages", [])
            stage_status = row_data.get("stage_status", {})
            done_count = 0
            open_count = 0
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
                    else:
                        open_count += 1
                    stage_item.setText(status)
                    stage_item.setBackground(_SUM_DATA_STATUS_FILL[status])
                    stage_item.setForeground(_SUM_DATA_STATUS_TEXT_COLOR[status])
                    stage_item.setToolTip(f"Stage {stage} — {status}")
                items.append(stage_item)

            vcr_item = QStandardItem("" if row_data.get("vcr") is None else str(row_data["vcr"]))
            vcr_item.setTextAlignment(Qt.AlignCenter)
            items.append(vcr_item)

            open_item = QStandardItem(str(open_count))
            open_item.setTextAlignment(Qt.AlignCenter)
            items.append(open_item)

            done_item = QStandardItem(str(done_count))
            done_item.setTextAlignment(Qt.AlignCenter)
            items.append(done_item)

            total_item = QStandardItem(str(len(required)))
            total_item.setTextAlignment(Qt.AlignCenter)
            items.append(total_item)

            pct_item = QStandardItem(_format_pct(done_count, len(required)))
            pct_item.setTextAlignment(Qt.AlignCenter)
            items.append(pct_item)

            for item in items:
                item.setEditable(False)
            for col_idx, item in enumerate(items):
                model.setItem(row_idx, col_idx, item)

        totals = live_sum_data_totals(rows)
        totals_items = [_totals_item("Totals"), _totals_item(""), _totals_item("")]
        for stage in range(1, STAGE_COUNT + 1):
            totals_items.append(_totals_item(str(totals["stage_open_counts"][stage])))
        totals_items.append(_totals_item(str(totals["vcr_open_count"])))
        totals_items.append(_totals_item(str(totals["open_total"])))
        totals_items.append(_totals_item(str(totals["done_total"])))
        totals_items.append(_totals_item(str(totals["grand_total"])))
        totals_items.append(_totals_item(_format_pct(totals["done_total"], totals["grand_total"])))
        for col_idx, item in enumerate(totals_items):
            model.setItem(len(rows), col_idx, item)

        table = FrozenTableView(frozen_columns=FROZEN_COLUMNS)
        table.setModel(model)
        table.setAlternatingRowColors(False)
        table.setWordWrap(True)
        table.verticalHeader().hide()
        table.setColumnWidth(0, 90)
        table.setColumnWidth(1, 320)
        table.setColumnWidth(2, 140)
        for col in range(FROZEN_COLUMNS, FROZEN_COLUMNS + STAGE_COUNT):
            table.setColumnWidth(col, 56)
        vcr_col = FROZEN_COLUMNS + STAGE_COUNT
        table.setColumnWidth(vcr_col, 56)
        table.setColumnWidth(vcr_col + 1, 56)  # Open
        table.setColumnWidth(vcr_col + 2, 56)  # Done
        table.setColumnWidth(vcr_col + 3, 56)  # Total
        table.setColumnWidth(vcr_col + 4, 80)  # % Complete
        table.horizontalHeader().setMinimumSectionSize(40)
        table.resizeRowsToContents()
        # Read-only view of the live status -- no click handler; the All
        # Data tab is the sole editing surface (see module docstring).
        return table

    # ------------------------------------------------------------------
    # Report tab
    # ------------------------------------------------------------------
    def _refresh_report_tab(self):
        while self.report_layout.count():
            child = self.report_layout.takeAt(0)
            if child.widget():
                child.widget().deleteLater()
        self.report_layout.addWidget(self._build_report_tab())

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
                row_data.get("description") or "",
                _format_total(row_data.get("total", 0)),
            ]
            for col_idx, value in enumerate(values):
                item = QTableWidgetItem(str(value))
                if col_idx == 3:
                    item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                if is_grand_total:
                    item.setFont(bold_font)
                table.setItem(row_idx, col_idx, item)

        table.setWordWrap(True)
        table.setColumnWidth(0, 220)
        table.setColumnWidth(1, 90)
        # Fixed width, not Stretch -- resizeRowsToContents() below needs a
        # settled column width to compute wrapped-text row heights against;
        # a Stretch column's width isn't final until the widget is shown.
        table.setColumnWidth(2, 400)
        table.setColumnWidth(3, 100)
        table.resizeRowsToContents()
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
