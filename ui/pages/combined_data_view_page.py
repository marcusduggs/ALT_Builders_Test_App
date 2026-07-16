"""
Read-only preview of a multi-increment combined report -- All Data /
Sum Data / Report / Changes, the SAME four tabs (and the same visual
structure: frozen columns, totals rows, tab layout) as
ui.pages.data_view_page.DataViewPage, but for several increments at
once, and with NO editing surface anywhere: no stage cell is
click-to-toggle, nothing here ever writes to status.json. "Export
Combined Report" (ui/pages/home_page.py) skips this preview entirely
and goes straight to a save dialog; "View Combined Data" opens this
page first so the user can look before exporting.

Every tab is built from ui.mock_data.CombinedView, which is itself
built by ui.mock_data.build_combined_view() from the exact same public
core.combined_export.build_combined_*/combine_all_data_totals functions
core.combined_export.export_combined_report() calls to write the
.xlsx -- not a re-derivation. This is the whole guarantee the feature
depends on: there is no second, independently-maintained copy of
"which rows, what order, what totals" for this preview to drift from
what gets exported. "Export to Excel" on this page calls that exact
same export_combined_report() function, against the exact same
self.view.increments list this page was built from.

All Data/Sum Data/Changes' State Revision Log/Update History/Comments
all gain a leading "Increment" column versus the single-increment tabs
(concatenated across increments, in on-screen order -- see
core/combined_export.py's module docstring for why concatenation was
chosen over Report-style sectioning). Report keeps its own existing
per-increment section/grouping treatment. Comments is read-only here
like everything else on this page -- add/delete a comment from that
specific increment's own Data View instead. Recently-added row
highlighting and the "needs status" badge are kept in All Data
purely as visual aids carried through from each increment's own
already-computed row data -- informational only, not exported (the
single-increment export doesn't render them either), so keeping or
dropping them carries no risk of preview/export drift.
"""

from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QFont, QIcon, QStandardItem, QStandardItemModel
from PySide6.QtWidgets import (
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from core import combined_export
from ui.mock_data import STAGE_COUNT, CombinedView, MockDataStore
from ui.pages.data_view_page import (
    _describe_change,
    _first_line,
    _format_changes_date,
    _format_pct,
    _format_total,
    _history_entry_summary,
    _needs_status_icon,
    _stage_status_icon,
    _stage_tooltip,
    _totals_item,
)
from ui.widgets.frozen_table import FrozenTableView
from ui.workers import run_with_progress

FROZEN_COLUMNS = 4  # Increment, Index, Description, Approval Agency
RECENTLY_ADDED_COLOR = QColor("#eaf7ee")

_SUM_DATA_STATUS_FILL = {
    "Done": QColor("#C6EFCE"),
    "Open": QColor("#FFC7CE"),
}
_SUM_DATA_STATUS_TEXT_COLOR = {
    "Done": QColor("#006100"),
    "Open": QColor("#9C0006"),
}


class CombinedDataViewPage(QWidget):
    def __init__(self, project_name: str, view: CombinedView, store: MockDataStore, on_back, parent=None):
        super().__init__(parent)
        self.project_name = project_name
        self.view = view
        self.store = store
        self.on_back = on_back

        outer = QVBoxLayout(self)
        outer.setContentsMargins(20, 20, 20, 20)
        outer.setSpacing(12)

        outer.addWidget(self._build_header())
        outer.addWidget(self._build_readonly_notice())

        # Explicit attribute references (not just retrievable via
        # findChild) -- both for the same reason DataViewPage keeps
        # self.table, and because two FrozenTableView instances (All
        # Data's and Sum Data's) are both alive simultaneously here, so
        # an untargeted findChild(FrozenTableView) can't reliably tell
        # them apart.
        self.tabs = QTabWidget()
        self.all_data_table = self._build_all_data_tab()
        self.sum_data_table = self._build_sum_data_tab()
        tabs = self.tabs
        tabs.addTab(self.all_data_table, "All Data")
        tabs.addTab(self.sum_data_table, "Sum Data")
        tabs.addTab(self._build_report_tab(), "Report")
        tabs.addTab(self._build_changes_tab(), "Changes")
        outer.addWidget(tabs, stretch=1)

        outer.addWidget(self._build_footer())

    # ------------------------------------------------------------------
    # header + export
    # ------------------------------------------------------------------
    def _build_header(self) -> QWidget:
        row = QWidget()
        layout = QHBoxLayout(row)
        layout.setContentsMargins(0, 0, 0, 0)

        back_button = QPushButton("< Back to Increments")
        back_button.setObjectName("secondaryButton")
        back_button.clicked.connect(lambda: self.on_back())

        count = len(self.view.increments)
        title = QLabel(f"Combined View — {count} Increment{'s' if count != 1 else ''}")
        title.setObjectName("pageTitle")

        names = ", ".join(inc.name for inc in self.view.increments)
        subtitle = QLabel(f"{self.project_name}  ·  {names}")
        subtitle.setObjectName("hint")
        subtitle.setWordWrap(True)

        title_box = QVBoxLayout()
        title_box.setSpacing(2)
        title_box.addWidget(title)
        title_box.addWidget(subtitle)

        export_button = QPushButton("Export to Excel")
        export_button.setObjectName("primaryButton")
        export_button.clicked.connect(self._on_export_clicked)

        layout.addWidget(back_button)
        layout.addSpacing(16)
        layout.addLayout(title_box)
        layout.addStretch(1)
        layout.addWidget(export_button)
        return row

    def _build_readonly_notice(self) -> QFrame:
        frame = QFrame()
        frame.setObjectName("changeSectionWarning")
        layout = QHBoxLayout(frame)
        layout.setContentsMargins(12, 8, 12, 8)
        label = QLabel(
            "Read-only preview — status marks cannot be changed here. Use a single increment's Data View to "
            "update status; export below to save this combined report."
        )
        label.setObjectName("changeSectionTitleWarning")
        label.setWordWrap(True)
        layout.addWidget(label)
        return frame

    def _on_export_clicked(self):
        default_name = combined_export.default_combined_filename(self.project_name, len(self.view.increments))
        path, _ = QFileDialog.getSaveFileName(self, "Export Combined Report", default_name, "Excel Workbook (*.xlsx)")
        if not path:
            return  # user cancelled -- nothing changes

        def on_finished(_result):
            QMessageBox.information(self, "Export Complete", f"Saved to:\n\n{path}")

        def on_error(exc):
            QMessageBox.critical(self, "Could Not Export", f"This file couldn't be saved:\n\n{exc}")

        # The exact same increments list (same versions, same on-screen
        # order) this page was built from -- see module docstring's
        # guarantee.
        run_with_progress(
            self, "Exporting combined report...", combined_export.export_combined_report,
            on_finished, on_error, self.view.increments, path,
        )

    # ------------------------------------------------------------------
    # All Data tab
    # ------------------------------------------------------------------
    def _build_all_data_tab(self) -> FrozenTableView:
        headers = ["Increment", "Index", "Description", "Approval Agency"]
        headers += [f"Stage {i}" for i in range(1, STAGE_COUNT + 1)]
        headers += ["VCR", "SUM"]

        rows = self.view.all_data_rows
        model = QStandardItemModel(len(rows) + 1, len(headers))  # +1 for the totals row
        model.setHorizontalHeaderLabels(headers)

        for row_idx, row_data in enumerate(rows):
            is_recently_added = bool(row_data.get("recently_added"))

            increment_item = QStandardItem(row_data["increment"])
            index_item = QStandardItem(str(row_data["index"]))
            if row_data.get("needs_status"):
                index_item.setIcon(_needs_status_icon())
                stages = ", ".join(str(s) for s in row_data.get("needs_status_stages", []))
                index_item.setToolTip(f"Needs status for stage(s): {stages}")

            description_item = QStandardItem(row_data.get("description") or "")
            description_item.setToolTip(row_data.get("description") or "")
            agency_item = QStandardItem(row_data.get("approval_agency") or "")

            items = [increment_item, index_item, description_item, agency_item]

            required = set(row_data.get("required_stages", []))
            stage_status = row_data.get("stage_status", {})
            for stage in range(1, STAGE_COUNT + 1):
                if stage in required:
                    stage_item = QStandardItem("")
                    status = stage_status.get(stage)
                    stage_item.setIcon(_stage_status_icon(status))
                    stage_item.setToolTip(_stage_tooltip(stage, status))
                else:
                    stage_item = QStandardItem("0")
                stage_item.setTextAlignment(Qt.AlignCenter)
                items.append(stage_item)

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

        totals = self.view.all_data_totals
        totals_items = [_totals_item("Totals"), _totals_item(""), _totals_item(""), _totals_item("")]
        for stage in range(1, STAGE_COUNT + 1):
            totals_items.append(_totals_item(_format_total(totals.get(f"Stage {stage}", 0))))
        totals_items.append(_totals_item(_format_total(totals.get("VCR", 0))))
        totals_items.append(_totals_item(_format_total(totals.get("SUM", 0))))
        for col_idx, item in enumerate(totals_items):
            model.setItem(len(rows), col_idx, item)

        table = FrozenTableView(frozen_columns=FROZEN_COLUMNS)
        table.setModel(model)
        table.setAlternatingRowColors(True)
        table.setWordWrap(True)
        table.verticalHeader().hide()
        table.setColumnWidth(0, 220)
        table.setColumnWidth(1, 90)
        table.setColumnWidth(2, 320)
        table.setColumnWidth(3, 140)
        for col in range(FROZEN_COLUMNS, FROZEN_COLUMNS + STAGE_COUNT):
            table.setColumnWidth(col, 56)
        table.setColumnWidth(FROZEN_COLUMNS + STAGE_COUNT, 56)
        table.setColumnWidth(FROZEN_COLUMNS + STAGE_COUNT + 1, 56)
        table.horizontalHeader().setMinimumSectionSize(40)
        table.resizeRowsToContents()
        # No table.clicked connection anywhere on this page -- see module
        # docstring: this whole page is read-only, by omission, not by a
        # disabled-but-present handler.
        return table

    # ------------------------------------------------------------------
    # Sum Data tab
    # ------------------------------------------------------------------
    def _build_sum_data_tab(self) -> FrozenTableView:
        headers = ["Increment", "Index", "Description", "Approval Agency"]
        headers += [f"Stage {i}" for i in range(1, STAGE_COUNT + 1)]
        headers += ["VCR", "Open", "Done", "Total", "% Complete"]

        rows = self.view.sum_data_rows
        model = QStandardItemModel(len(rows) + 1, len(headers))
        model.setHorizontalHeaderLabels(headers)

        for row_idx, row_data in enumerate(rows):
            increment_item = QStandardItem(row_data["increment"])
            index_item = QStandardItem(str(row_data["index"]))
            description_item = QStandardItem(row_data.get("description") or "")
            description_item.setToolTip(row_data.get("description") or "")
            agency_item = QStandardItem(row_data.get("approval_agency") or "")

            items = [increment_item, index_item, description_item, agency_item]

            required = row_data.get("required_stages", [])
            live_status = row_data["live_status"]  # precomputed by core.combined_export.build_combined_sum_data_rows
            for stage in range(1, STAGE_COUNT + 1):
                stage_item = QStandardItem("")
                stage_item.setTextAlignment(Qt.AlignCenter)
                if stage in required:
                    status = live_status[stage]
                    stage_item.setText(status)
                    stage_item.setBackground(_SUM_DATA_STATUS_FILL[status])
                    stage_item.setForeground(_SUM_DATA_STATUS_TEXT_COLOR[status])
                    stage_item.setToolTip(f"Stage {stage} — {status}")
                items.append(stage_item)

            vcr_item = QStandardItem("" if row_data.get("vcr") is None else str(row_data["vcr"]))
            vcr_item.setTextAlignment(Qt.AlignCenter)
            items.append(vcr_item)

            open_item = QStandardItem(str(row_data["open_count"]))
            open_item.setTextAlignment(Qt.AlignCenter)
            items.append(open_item)

            done_item = QStandardItem(str(row_data["done_count"]))
            done_item.setTextAlignment(Qt.AlignCenter)
            items.append(done_item)

            total_item = QStandardItem(str(row_data["total_count"]))
            total_item.setTextAlignment(Qt.AlignCenter)
            items.append(total_item)

            pct_item = QStandardItem(_format_pct(row_data["done_count"], row_data["total_count"]))
            pct_item.setTextAlignment(Qt.AlignCenter)
            items.append(pct_item)

            for item in items:
                item.setEditable(False)
            for col_idx, item in enumerate(items):
                model.setItem(row_idx, col_idx, item)

        totals = self.view.sum_data_totals
        totals_items = [_totals_item("Totals"), _totals_item(""), _totals_item(""), _totals_item("")]
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
        table.setAlternatingRowColors(True)
        table.setWordWrap(True)
        table.verticalHeader().hide()
        table.setColumnWidth(0, 220)
        table.setColumnWidth(1, 90)
        table.setColumnWidth(2, 320)
        table.setColumnWidth(3, 140)
        for col in range(FROZEN_COLUMNS, FROZEN_COLUMNS + STAGE_COUNT):
            table.setColumnWidth(col, 56)
        vcr_col = FROZEN_COLUMNS + STAGE_COUNT
        table.setColumnWidth(vcr_col, 56)
        table.setColumnWidth(vcr_col + 1, 56)
        table.setColumnWidth(vcr_col + 2, 56)
        table.setColumnWidth(vcr_col + 3, 56)
        table.setColumnWidth(vcr_col + 4, 80)
        table.horizontalHeader().setMinimumSectionSize(40)
        table.resizeRowsToContents()
        return table

    # ------------------------------------------------------------------
    # Report tab
    # ------------------------------------------------------------------
    def _build_report_tab(self) -> QTableWidget:
        headers = ["Approval Agency", "Index", "Description", "Total"]
        rows = self.view.report_rows

        table = QTableWidget(len(rows), len(headers))
        table.setHorizontalHeaderLabels(headers)
        table.verticalHeader().hide()
        table.setEditTriggers(QTableWidget.NoEditTriggers)
        table.setAlternatingRowColors(True)
        table.setSelectionBehavior(QTableWidget.SelectRows)

        bold_font = QFont()
        bold_font.setBold(True)

        for row_idx, row_data in enumerate(rows):
            is_summary_row = row_data["is_section_header"] or row_data["is_combined_grand_total"]
            total_value = row_data["total"]
            values = [
                row_data["approval_agency"] or "",
                row_data["index"] or "",
                row_data["description"] or "",
                _format_total(total_value) if total_value is not None else "",
            ]
            for col_idx, value in enumerate(values):
                item = QTableWidgetItem(str(value))
                if col_idx == 3:
                    item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                if is_summary_row or row_data["is_grand_total"]:
                    item.setFont(bold_font)
                if is_summary_row:
                    item.setBackground(QColor("#e4e8f0"))
                table.setItem(row_idx, col_idx, item)

        table.setWordWrap(True)
        table.setColumnWidth(0, 220)
        table.setColumnWidth(1, 90)
        table.setColumnWidth(2, 400)
        table.setColumnWidth(3, 100)
        table.resizeRowsToContents()
        return table

    # ------------------------------------------------------------------
    # Changes tab
    # ------------------------------------------------------------------
    def _build_changes_tab(self) -> QSplitter:
        splitter = QSplitter(Qt.Vertical)
        splitter.addWidget(self._build_state_revision_log_section())
        splitter.addWidget(self._build_update_history_section())
        splitter.addWidget(self._build_comments_section())
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 1)
        splitter.setStretchFactor(2, 1)
        return splitter

    def _build_state_revision_log_section(self) -> QWidget:
        container = QWidget()
        layout = QVBoxLayout(container)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)

        header = QLabel("State Revision Log")
        header.setObjectName("sectionTitle")
        layout.addWidget(header)

        caption = QLabel(
            "Every selected increment's own revision log, concatenated in on-screen order with a leading "
            "Increment column -- see core/combined_export.py for why this is flattened, not sectioned."
        )
        caption.setObjectName("hint")
        caption.setWordWrap(True)
        layout.addWidget(caption)

        layout.addWidget(self._build_state_revision_log_table(), stretch=1)
        return container

    def _build_state_revision_log_table(self) -> QTableWidget:
        headers = [
            "Increment", "Rev #", "Synopsis of Change", "AOR Signature", "SEOR Signature", "Effective Date",
            "HCAI Concurrence",
        ]
        rows = self.view.revision_log_rows

        table = QTableWidget(len(rows), len(headers))
        table.setHorizontalHeaderLabels(headers)
        table.verticalHeader().hide()
        table.setEditTriggers(QTableWidget.NoEditTriggers)
        table.setAlternatingRowColors(True)
        table.setSelectionBehavior(QTableWidget.SelectRows)
        table.setWordWrap(True)

        for row_idx, row_data in enumerate(rows):
            values = [
                row_data["increment"],
                row_data.get("revision_number"),
                row_data.get("synopsis") or "",
                _format_changes_date(row_data.get("aor_signature_date")),
                _format_changes_date(row_data.get("seor_signature_date")),
                _format_changes_date(row_data.get("effective_date")),
                row_data.get("hcai_concurrence") or "",
            ]
            for col_idx, value in enumerate(values):
                item = QTableWidgetItem(str(value))
                if col_idx == 1:
                    item.setTextAlignment(Qt.AlignCenter)
                table.setItem(row_idx, col_idx, item)

        table.setColumnWidth(0, 220)
        table.setColumnWidth(1, 55)
        table.setColumnWidth(2, 400)
        table.setColumnWidth(3, 110)
        table.setColumnWidth(4, 110)
        table.setColumnWidth(5, 110)
        table.setColumnWidth(6, 160)
        table.resizeRowsToContents()
        return table

    def _build_update_history_section(self) -> QWidget:
        """change_history.json entries, flattened across every selected
        increment -- each increment's own entries newest-first, kept
        contiguous in on-screen increment order (NOT globally interleaved
        by date -- see core/combined_export.py's module docstring), same
        ordering the Changes sheet export uses. Each entry's summary line
        is prefixed with its increment name, since multiple increments'
        histories now share one flat list instead of each having its own
        section (ui.pages.data_view_page's single-increment version has
        no need for this prefix -- there's only ever one increment there).
        """
        container = QWidget()
        layout = QVBoxLayout(container)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)

        header = QLabel("Update History")
        header.setObjectName("sectionTitle")
        layout.addWidget(header)

        caption = QLabel(
            "Every confirmed update across every selected increment -- each increment's own entries "
            "newest-first, grouped by increment in on-screen order."
        )
        caption.setObjectName("hint")
        caption.setWordWrap(True)
        layout.addWidget(caption)

        entries = self.view.update_history_rows

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setStyleSheet("QScrollArea { background: transparent; }")
        scroll.viewport().setStyleSheet("background: transparent;")
        content = QWidget()
        content.setStyleSheet("background: transparent;")
        content_layout = QVBoxLayout(content)
        content_layout.setSpacing(6)
        content_layout.setAlignment(Qt.AlignTop)

        if not entries:
            empty = QLabel("No updates confirmed yet for any selected increment.")
            empty.setObjectName("hint")
            content_layout.addWidget(empty)
        else:
            for entry in entries:
                content_layout.addWidget(self._build_history_entry_widget(entry))

        content_layout.addStretch(1)
        scroll.setWidget(content)
        layout.addWidget(scroll, stretch=1)
        return container

    def _build_history_entry_widget(self, entry: dict) -> QFrame:
        frame = QFrame()
        frame.setObjectName("changeSectionInfo")
        layout = QVBoxLayout(frame)
        layout.setContentsMargins(12, 8, 12, 8)
        layout.setSpacing(4)

        date = (entry.get("timestamp") or "").split("T", 1)[0] or "unknown date"
        old_v, new_v = entry.get("old_version"), entry.get("new_version")
        summary_text = f"{entry['increment']} — v{old_v} → v{new_v} ({date}): {_history_entry_summary(entry)}"

        toggle = QPushButton(f"▸  {summary_text}")
        toggle.setObjectName("historyEntryToggle")
        toggle.setCheckable(True)
        toggle.setFlat(True)
        toggle.setStyleSheet("text-align: left; font-weight: 600; border: none;")
        toggle.setCursor(Qt.PointingHandCursor)

        detail = QWidget()
        detail.setVisible(False)
        detail_layout = QVBoxLayout(detail)
        detail_layout.setContentsMargins(20, 6, 0, 0)
        detail_layout.setSpacing(8)
        self._populate_history_entry_detail(detail_layout, entry)

        def on_toggled(checked: bool):
            detail.setVisible(checked)
            arrow = "▾" if checked else "▸"
            toggle.setText(f"{arrow}  {summary_text}")

        toggle.toggled.connect(on_toggled)

        layout.addWidget(toggle)
        layout.addWidget(detail)
        return frame

    def _populate_history_entry_detail(self, layout: QVBoxLayout, entry: dict):
        added = entry.get("added_items") or []
        removed = entry.get("removed_items") or []
        anomalies = entry.get("column_anomalies") or []
        value_changed = entry.get("value_changed_items") or []

        if not (added or removed or anomalies or value_changed):
            label = QLabel("No changes detected -- this update replaced the file with an item-for-item identical upload.")
            label.setObjectName("hint")
            label.setWordWrap(True)
            layout.addWidget(label)
            return

        if added:
            layout.addWidget(self._build_history_item_group(f"Added Items ({len(added)})", added))
        if removed:
            layout.addWidget(self._build_history_item_group(f"Removed Items ({len(removed)})", removed))
        if value_changed:
            layout.addWidget(self._build_history_value_changed_group(value_changed))
        if anomalies:
            layout.addWidget(self._build_history_anomaly_group(anomalies))

    def _build_history_item_group(self, title: str, items: list[dict]) -> QWidget:
        group = QWidget()
        layout = QVBoxLayout(group)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(2)
        header = QLabel(title)
        header.setStyleSheet("font-weight: 600;")
        layout.addWidget(header)
        for item in items:
            row = QLabel(f"<b>{item.get('index')}</b> &nbsp;&mdash;&nbsp; {_first_line(item.get('description'))}")
            row.setWordWrap(True)
            row.setTextFormat(Qt.RichText)
            layout.addWidget(row)
        return group

    def _build_history_value_changed_group(self, items: list[dict]) -> QWidget:
        group = QWidget()
        layout = QVBoxLayout(group)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(2)
        header = QLabel(f"Values Changed ({len(items)})")
        header.setStyleSheet("font-weight: 600;")
        layout.addWidget(header)
        for item in items:
            change_lines = "<br>".join(
                f"<span style='color:#6b7280;'>{_describe_change(field_label, old_v, new_v)}</span>"
                for field_label, (old_v, new_v) in (item.get("changes") or {}).items()
            )
            row = QLabel(
                f"<b>{item.get('index')}</b> &nbsp;&mdash;&nbsp; {_first_line(item.get('description'))}"
                f"<br>{change_lines}"
            )
            row.setWordWrap(True)
            row.setTextFormat(Qt.RichText)
            layout.addWidget(row)
        return group

    def _build_history_anomaly_group(self, anomalies: list[str]) -> QWidget:
        group = QWidget()
        layout = QVBoxLayout(group)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(2)
        header = QLabel(f"Column Anomalies ({len(anomalies)})")
        header.setStyleSheet("font-weight: 600; color: #9b6a00;")
        layout.addWidget(header)
        for anomaly in anomalies:
            row = QLabel(anomaly)
            row.setWordWrap(True)
            layout.addWidget(row)
        return group

    def _build_comments_section(self) -> QWidget:
        """comments.json entries, flattened across every selected
        increment -- each increment's own comments newest-first, kept
        contiguous in on-screen increment order, same as State Revision
        Log/Update History above (see core/combined_export.py's module
        docstring). Read-only here, unlike ui.pages.data_view_page's
        single-increment version -- this whole page never writes
        anything (see module docstring); adding/deleting a comment
        requires opening that specific increment's own Data View.
        """
        container = QWidget()
        layout = QVBoxLayout(container)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)

        header = QLabel("Comments")
        header.setObjectName("sectionTitle")
        layout.addWidget(header)

        caption = QLabel(
            "Every comment across every selected increment -- each increment's own comments newest-first, "
            "grouped by increment in on-screen order. Read-only here -- add or delete from that increment's "
            "own Data View."
        )
        caption.setObjectName("hint")
        caption.setWordWrap(True)
        layout.addWidget(caption)

        layout.addWidget(self._build_comments_table(), stretch=1)
        return container

    def _build_comments_table(self) -> QTableWidget:
        headers = ["Increment", "#", "Comment", "Date"]
        rows = self.view.comments_rows

        table = QTableWidget(len(rows), len(headers))
        table.setHorizontalHeaderLabels(headers)
        table.verticalHeader().hide()
        table.setEditTriggers(QTableWidget.NoEditTriggers)
        table.setAlternatingRowColors(True)
        table.setSelectionBehavior(QTableWidget.SelectRows)
        table.setWordWrap(True)

        for row_idx, row_data in enumerate(rows):
            values = [
                row_data["increment"],
                row_data.get("comment_number"),
                row_data.get("text") or "",
                (row_data.get("timestamp") or "").split("T", 1)[0],
            ]
            for col_idx, value in enumerate(values):
                item = QTableWidgetItem(str(value))
                if col_idx == 1:
                    item.setTextAlignment(Qt.AlignCenter)
                table.setItem(row_idx, col_idx, item)

        table.setColumnWidth(0, 220)
        table.setColumnWidth(1, 40)
        table.setColumnWidth(2, 500)
        table.setColumnWidth(3, 110)
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

        total_items = len(self.view.all_data_rows)
        total_label = QLabel(f"Total items: {total_items}")
        layout.addWidget(total_label)

        needing_status = sum(1 for row in self.view.all_data_rows if row.get("needs_status"))
        needing_label = QLabel(f"Needing status: {needing_status}" if needing_status else "Needing status: none")
        if needing_status:
            needing_label.setObjectName("footerHighlight")
        layout.addWidget(needing_label)

        layout.addStretch(1)
        return frame
