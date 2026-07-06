"""
The review screen shown after "simulating" a comparison between the
currently-stored version of an increment and a newly uploaded file.

Deliberately NOT a casual popup: full-size modal dialog, a distinctly
styled "Confirm & Update" action that requires a second, explicit
confirmation dialog restating the consequences, and a "Cancel/Discard"
action that requires none (declining a change is always safe).
"""

from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)

from ui.mock_data import ComparisonResult


def _first_line(text: str | None) -> str:
    if not text:
        return ""
    return text.split("\n", 1)[0]


def _describe_stage_value(value) -> str:
    """Translates a raw Stage-cell value into the same short vocabulary
    Rey's team already reads on the source file: blank/0 is "not
    required", "X" is the Done marker, and any other number is a literal
    test count (see core/value_diff.py's module docstring -- a value
    showing up where there was none before is exactly what this feature
    exists to surface, not just a Done/Open toggle).
    """
    if value in (0, None, ""):
        return "not required"
    if isinstance(value, str) and value.strip().lower() == "x":
        return "required (X)"
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        n = int(value) if float(value).is_integer() else value
        return f"{n} test{'s' if n != 1 else ''} required"
    return str(value)


def _describe_change(field_label: str, old_value, new_value) -> str:
    if field_label.startswith("Stage "):
        return f"{field_label}: {_describe_stage_value(old_value)} → {_describe_stage_value(new_value)}"
    old_text = old_value if old_value not in (None, "") else "(blank)"
    new_text = new_value if new_value not in (None, "") else "(blank)"
    return f"{field_label}: {old_text!s} → {new_text!s}"


def _format_total(value: float) -> str:
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


class ReviewDialog(QDialog):
    def __init__(self, increment_name: str, result: ComparisonResult, parent=None):
        super().__init__(parent)
        self.result = result
        self.setWindowTitle("Review Changes")
        self.setModal(True)
        self.resize(880, 640)

        outer = QVBoxLayout(self)
        outer.setContentsMargins(20, 20, 20, 20)
        outer.setSpacing(16)

        outer.addWidget(self._build_header(increment_name, result))

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        # QScrollArea's viewport auto-fills with the palette's Window color
        # by default, which follows the OS light/dark setting rather than
        # this dialog's own (light) background -- explicit transparent
        # backgrounds here keep the review screen legible under a dark
        # system theme instead of showing a black scroll area.
        scroll.setStyleSheet("QScrollArea { background: transparent; }")
        scroll.viewport().setStyleSheet("background: transparent;")
        content = QWidget()
        content.setStyleSheet("background: transparent;")
        content_layout = QVBoxLayout(content)
        content_layout.setSpacing(14)
        content_layout.setAlignment(Qt.AlignTop)

        if not result.has_changes:
            content_layout.addWidget(self._build_no_changes_banner())
        else:
            if result.added_items:
                content_layout.addWidget(
                    self._build_section(
                        "Added Items",
                        "changeSectionAdded",
                        "changeSectionTitleAdded",
                        result.added_items,
                        "New test item found in the uploaded file.",
                    )
                )
            if result.removed_items:
                content_layout.addWidget(
                    self._build_section(
                        "Removed Items",
                        "changeSectionRemoved",
                        "changeSectionTitleRemoved",
                        result.removed_items,
                        "No longer present in the uploaded file.",
                    )
                )
            if result.column_anomalies:
                content_layout.addWidget(self._build_anomaly_section(result.column_anomalies))
            if result.value_changed_items:
                content_layout.addWidget(self._build_value_changed_section(result.value_changed_items))
            if result.needs_status_items:
                content_layout.addWidget(self._build_needs_status_section(result.needs_status_items))

        content_layout.addStretch(1)
        scroll.setWidget(content)
        outer.addWidget(scroll, stretch=1)

        outer.addWidget(self._build_button_row())

    # ------------------------------------------------------------------
    def _build_header(self, increment_name: str, result: ComparisonResult) -> QFrame:
        frame = QFrame()
        frame.setObjectName("reviewHeader")
        layout = QVBoxLayout(frame)
        layout.setContentsMargins(18, 14, 18, 14)
        layout.setSpacing(4)

        title = QLabel(f"Comparing {increment_name}: Version {result.current_version} (current) vs. new upload")
        title.setStyleSheet("font-size: 16px; font-weight: 600;")
        title.setWordWrap(True)
        subtitle = QLabel(
            "Review every change below before confirming. Nothing is applied until you choose Confirm & Update."
        )
        subtitle.setObjectName("reviewSubtitle")
        subtitle.setWordWrap(True)

        layout.addWidget(title)
        layout.addWidget(subtitle)

        if result.has_changes:
            summary = QLabel(self._summary_text(result))
            summary.setObjectName("reviewSubtitle")
            summary.setStyleSheet("font-weight: 600;")
            summary.setWordWrap(True)
            layout.addWidget(summary)

        if result.report_total_before != result.report_total_after:
            report_preview = QLabel(
                f"Report total will change from {_format_total(result.report_total_before)} to "
                f"{_format_total(result.report_total_after)}."
            )
            report_preview.setObjectName("reviewSubtitle")
            layout.addWidget(report_preview)

        return frame

    @staticmethod
    def _summary_text(result: ComparisonResult) -> str:
        """A compact "N new, M removed, K value-changed" line so counts of
        different KINDS of change are distinguishable at a glance, rather
        than lumped into a single number (added items and value-changed
        items are never the same items -- see core/value_diff.py).
        """
        parts = []
        if result.added_items:
            n = len(result.added_items)
            parts.append(f"{n} new item{'s' if n != 1 else ''}")
        if result.removed_items:
            n = len(result.removed_items)
            parts.append(f"{n} item{'s' if n != 1 else ''} removed")
        if result.value_changed_items:
            n = len(result.value_changed_items)
            parts.append(f"{n} item{'s' if n != 1 else ''} with an updated value")
        if result.column_anomalies:
            n = len(result.column_anomalies)
            parts.append(f"{n} column anomal{'y' if n == 1 else 'ies'}")
        return ", ".join(parts)

    def _build_no_changes_banner(self) -> QFrame:
        frame = QFrame()
        frame.setObjectName("noChangesBanner")
        layout = QVBoxLayout(frame)
        layout.setContentsMargins(18, 18, 18, 18)
        layout.setSpacing(4)
        title = QLabel("No changes detected")
        title.setStyleSheet("font-size: 15px; font-weight: 600; color: #2c4c9e;")
        detail = QLabel(
            "The uploaded file matches the currently stored version item-for-item: no items were added, "
            "removed, or changed in value, and no column layout changes were found. You can still confirm to "
            "store this file as the new version, or cancel."
        )
        detail.setWordWrap(True)
        layout.addWidget(title)
        layout.addWidget(detail)
        return frame

    def _build_section(self, title, frame_object_name, title_object_name, items, caption) -> QFrame:
        frame = QFrame()
        frame.setObjectName(frame_object_name)
        layout = QVBoxLayout(frame)
        layout.setContentsMargins(16, 12, 16, 12)
        layout.setSpacing(6)

        header = QLabel(f"{title} ({len(items)})")
        header.setObjectName(title_object_name)
        header.setStyleSheet(header.styleSheet() + "font-size: 14px;")
        layout.addWidget(header)

        caption_label = QLabel(caption)
        caption_label.setObjectName("hint")
        layout.addWidget(caption_label)

        for item in items:
            row = QLabel(f"<b>{item['index']}</b> &nbsp;&mdash;&nbsp; {_first_line(item.get('description'))}")
            row.setWordWrap(True)
            row.setTextFormat(Qt.RichText)
            layout.addWidget(row)

        return frame

    def _build_value_changed_section(self, items: list[dict]) -> QFrame:
        """Items present in both the current and uploaded file where one
        or more field values differ -- e.g. a stage mark that used to be
        blank now shows a real test count. See core/value_diff.py.
        """
        frame = QFrame()
        frame.setObjectName("changeSectionInfo")
        layout = QVBoxLayout(frame)
        layout.setContentsMargins(16, 12, 16, 12)
        layout.setSpacing(6)

        header = QLabel(f"Values Changed ({len(items)})")
        header.setObjectName("changeSectionTitleInfo")
        header.setStyleSheet(header.styleSheet() + "font-size: 14px;")
        layout.addWidget(header)

        caption = QLabel(
            "Present in both files, but one or more values differ -- nothing was added or removed, the data "
            "itself changed."
        )
        caption.setObjectName("hint")
        caption.setWordWrap(True)
        layout.addWidget(caption)

        for item in items:
            change_lines = "<br>".join(
                f"<span style='color:#6b7280;'>{_describe_change(field_label, old_v, new_v)}</span>"
                for field_label, (old_v, new_v) in item["changes"].items()
            )
            row = QLabel(
                f"<b>{item['index']}</b> &nbsp;&mdash;&nbsp; {_first_line(item.get('description'))}"
                f"<br>{change_lines}"
            )
            row.setWordWrap(True)
            row.setTextFormat(Qt.RichText)
            layout.addWidget(row)

        return frame

    def _build_needs_status_section(self, items: list[dict]) -> QFrame:
        """Unlike _build_section (added/removed, which are whole-item
        concerns), status lives at the (item, stage) level -- so each row
        here names the specific stages that need a Done/Open mark, not
        just the item as a whole.
        """
        total_stages = sum(len(item.get("stages", [])) for item in items)
        frame = QFrame()
        frame.setObjectName("changeSectionInfo")
        layout = QVBoxLayout(frame)
        layout.setContentsMargins(16, 12, 16, 12)
        layout.setSpacing(6)

        header = QLabel(f"Items Needing Status ({total_stages} stage{'s' if total_stages != 1 else ''} "
                         f"across {len(items)} item{'s' if len(items) != 1 else ''})")
        header.setObjectName("changeSectionTitleInfo")
        header.setStyleSheet(header.styleSheet() + "font-size: 14px;")
        layout.addWidget(header)

        caption_label = QLabel(
            "New item(s) found in the uploaded file -- no prior Done/Open mark for these specific stages, "
            "Rey's team will need to set one."
        )
        caption_label.setObjectName("hint")
        caption_label.setWordWrap(True)
        layout.addWidget(caption_label)

        for item in items:
            stages_text = ", ".join(str(s) for s in item.get("stages", []))
            row = QLabel(
                f"<b>{item['index']}</b> &nbsp;&mdash;&nbsp; {_first_line(item.get('description'))}"
                f"<br><span style='color:#6b7280;'>Stages needing status: {stages_text}</span>"
            )
            row.setWordWrap(True)
            row.setTextFormat(Qt.RichText)
            layout.addWidget(row)

        return frame

    def _build_anomaly_section(self, anomalies: list[str]) -> QFrame:
        frame = QFrame()
        frame.setObjectName("changeSectionWarning")
        layout = QVBoxLayout(frame)
        layout.setContentsMargins(16, 12, 16, 12)
        layout.setSpacing(6)

        header = QLabel(f"Column Anomalies ({len(anomalies)})")
        header.setObjectName("changeSectionTitleWarning")
        header.setStyleSheet(header.styleSheet() + "font-size: 14px;")
        layout.addWidget(header)

        caption = QLabel(
            "The uploaded file's column layout doesn't match what was expected. Data may have been read from "
            "the wrong column -- double-check this sheet before confirming."
        )
        caption.setObjectName("hint")
        caption.setWordWrap(True)
        layout.addWidget(caption)

        for anomaly in anomalies:
            row = QLabel(anomaly)
            row.setWordWrap(True)
            layout.addWidget(row)

        return frame

    def _build_button_row(self) -> QWidget:
        row = QWidget()
        layout = QHBoxLayout(row)
        layout.setContentsMargins(0, 0, 0, 0)

        discard_button = QPushButton("Cancel / Discard")
        discard_button.setObjectName("discardButton")
        discard_button.setAutoDefault(False)
        discard_button.clicked.connect(self.reject)

        confirm_button = QPushButton("Confirm && Update")
        confirm_button.setObjectName("confirmButton")
        confirm_button.setAutoDefault(False)
        confirm_button.setDefault(False)
        confirm_button.clicked.connect(self._on_confirm_clicked)

        layout.addWidget(discard_button)
        layout.addStretch(1)
        layout.addWidget(confirm_button)
        return row

    def _on_confirm_clicked(self):
        total_stages = sum(len(item.get("stages", [])) for item in self.result.needs_status_items)
        answer = QMessageBox.question(
            self,
            "Confirm Update",
            (
                f"This will replace the current data for this increment (version {self.result.current_version}) "
                "with the new upload.\n\n"
                "Status marks (Done/Open) on unchanged stages will be preserved. "
                f"{total_stages} stage(s) across {len(self.result.needs_status_items)} item(s) will need a "
                "status set after this update.\n\n"
                "Continue?"
            ),
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if answer == QMessageBox.Yes:
            self.accept()
