"""
Home page: Project bar (Part 1) + per-project Increment list (Part 2) +
the Upload/Review flow (Part 3) that lives on top of it.
"""

from __future__ import annotations

import logging
from functools import partial

from PySide6.QtCore import Qt
from PySide6.QtGui import QPixmap
from PySide6.QtWidgets import (
    QComboBox,
    QDialog,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from core.increment_matcher import MatchType
from ui.dialogs.project_dialog import ProjectDialog
from ui.dialogs.review_dialog import ReviewDialog
from ui.mock_data import MockDataStore
from ui.paths import get_asset_path
from ui.workers import run_with_progress

logger = logging.getLogger(__name__)

EXCEL_FILE_FILTER = "TIO Workbooks (*.xlsm *.xlsx)"
LOGO_HEIGHT_PX = 36


class HomePage(QWidget):
    def __init__(self, store: MockDataStore, on_view_data, parent=None):
        super().__init__(parent)
        self.store = store
        self.on_view_data = on_view_data  # callback(project_name, increment_name)

        outer = QVBoxLayout(self)
        outer.setContentsMargins(20, 20, 20, 20)
        outer.setSpacing(16)

        outer.addWidget(self._build_project_bar())
        outer.addWidget(self._build_increment_section(), stretch=1)

        self._refresh_project_combo()

    # ------------------------------------------------------------------
    # Part 1: project bar
    # ------------------------------------------------------------------
    def _build_project_bar(self) -> QFrame:
        frame = QFrame()
        frame.setObjectName("card")
        layout = QHBoxLayout(frame)
        layout.setContentsMargins(16, 14, 16, 14)
        layout.setSpacing(12)

        logo_label = self._build_logo_label()
        if logo_label is not None:
            layout.addWidget(logo_label)
            layout.addSpacing(8)

        label = QLabel("Project:")
        label.setObjectName("sectionTitle")

        self.project_combo = QComboBox()
        self.project_combo.setMinimumWidth(320)
        self.project_combo.currentIndexChanged.connect(self._on_project_changed)

        add_button = QPushButton("Add")
        update_button = QPushButton("Update")
        delete_button = QPushButton("Delete")
        delete_button.setObjectName("dangerButton")
        add_button.clicked.connect(self._on_add_project)
        update_button.clicked.connect(self._on_update_project)
        delete_button.clicked.connect(self._on_delete_project)

        layout.addWidget(label)
        layout.addWidget(self.project_combo, stretch=1)
        layout.addWidget(add_button)
        layout.addWidget(update_button)
        layout.addWidget(delete_button)
        return frame

    def _build_logo_label(self) -> QLabel | None:
        """Loads assets/logo.png (resolved relative to the app's own
        location, not the cwd -- see ui/paths.py) and returns it scaled to
        toolbar height. Returns None if the file is missing/unreadable so
        the header bar just renders without it instead of crashing or
        showing a broken-image icon.
        """
        logo_path = get_asset_path("logo.png")
        pixmap = QPixmap(str(logo_path))
        if pixmap.isNull():
            logger.warning("Could not load logo image at %s; skipping.", logo_path)
            return None

        label = QLabel()
        label.setPixmap(
            pixmap.scaledToHeight(LOGO_HEIGHT_PX, Qt.SmoothTransformation)
        )
        return label

    def _refresh_project_combo(self, select_name: str | None = None):
        self.project_combo.blockSignals(True)
        self.project_combo.clear()
        self.project_combo.addItems(self.store.project_names())
        if select_name and select_name in self.store.project_names():
            self.project_combo.setCurrentText(select_name)
        self.project_combo.blockSignals(False)
        self._refresh_increment_table()

    def _current_project_name(self) -> str | None:
        return self.project_combo.currentText() or None

    def _on_project_changed(self, _index):
        self._refresh_increment_table()

    def _on_add_project(self):
        dialog = ProjectDialog(project=None, parent=self)
        if dialog.exec() == QDialog.Accepted:
            name = dialog.values()
            if name in self.store.project_names():
                QMessageBox.warning(self, "Project Exists", f"A project named '{name}' already exists.")
                return
            self.store.add_project(name)
            self._refresh_project_combo(select_name=name)

    def _on_update_project(self):
        current_name = self._current_project_name()
        if not current_name:
            return
        project = self.store.get_project(current_name)
        dialog = ProjectDialog(project=project, parent=self)
        if dialog.exec() == QDialog.Accepted:
            name = dialog.values()
            self.store.update_project(current_name, name)
            self._refresh_project_combo(select_name=name)

    def _on_delete_project(self):
        current_name = self._current_project_name()
        if not current_name:
            return
        answer = QMessageBox.question(
            self,
            "Delete Project",
            f"Delete project '{current_name}'? This removes all of its increments from this app and cannot be undone.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if answer == QMessageBox.Yes:
            self.store.delete_project(current_name)
            self._refresh_project_combo()

    # ------------------------------------------------------------------
    # Part 2: increment list
    # ------------------------------------------------------------------
    def _build_increment_section(self) -> QFrame:
        frame = QFrame()
        frame.setObjectName("card")
        layout = QVBoxLayout(frame)
        layout.setContentsMargins(16, 14, 16, 14)
        layout.setSpacing(10)

        header_row = QHBoxLayout()
        title = QLabel("Increments")
        title.setObjectName("sectionTitle")
        upload_button = QPushButton("Upload File")
        upload_button.setObjectName("primaryButton")
        upload_button.clicked.connect(self._on_upload_file)
        header_row.addWidget(title)
        header_row.addStretch(1)
        header_row.addWidget(upload_button)
        layout.addLayout(header_row)

        self.increment_table = QTableWidget(0, 4)
        self.increment_table.setHorizontalHeaderLabels(["Increment", "Version", "Last Updated", "Actions"])
        self.increment_table.verticalHeader().hide()
        self.increment_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.increment_table.setSelectionMode(QTableWidget.NoSelection)
        header = self.increment_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        layout.addWidget(self.increment_table)

        self.empty_state_label = QLabel("No increments yet for this project. Use “Upload File” to add one.")
        self.empty_state_label.setObjectName("hint")
        self.empty_state_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.empty_state_label)

        return frame

    def _refresh_increment_table(self):
        project_name = self._current_project_name()
        project = self.store.get_project(project_name) if project_name else None
        increments = project.increments if project else []

        self.increment_table.setRowCount(len(increments))
        for row, increment in enumerate(increments):
            self.increment_table.setItem(row, 0, QTableWidgetItem(increment.name))
            self.increment_table.setItem(row, 1, QTableWidgetItem(f"Version {increment.version}"))
            self.increment_table.setItem(row, 2, QTableWidgetItem(increment.last_updated))

            actions = QWidget()
            actions_layout = QHBoxLayout(actions)
            actions_layout.setContentsMargins(4, 2, 4, 2)
            actions_layout.setSpacing(6)

            view_button = QPushButton("View Data")
            view_button.setObjectName("primaryButton")
            view_button.clicked.connect(partial(self.on_view_data, project_name, increment.name))

            delete_button = QPushButton("Delete")
            delete_button.setObjectName("dangerButton")
            delete_button.clicked.connect(partial(self._on_delete_increment, project_name, increment.name))

            actions_layout.addWidget(view_button)
            actions_layout.addWidget(delete_button)
            self.increment_table.setCellWidget(row, 3, actions)

        self.increment_table.setVisible(bool(increments))
        self.empty_state_label.setVisible(not increments)

    def _on_delete_increment(self, project_name: str, increment_name: str):
        answer = QMessageBox.question(
            self,
            "Delete Increment",
            f"Delete increment '{increment_name}'? This removes all of its versions from this app and cannot be undone.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if answer == QMessageBox.Yes:
            self.store.delete_increment(project_name, increment_name)
            self._refresh_increment_table()

    # ------------------------------------------------------------------
    # Part 3: upload + review flow
    # ------------------------------------------------------------------
    def _on_upload_file(self):
        """The single upload entry point: identifies the file's increment
        (core.excel_reader.get_record_name) and routes to whichever of
        the two flows below applies, based on
        core.increment_matcher.match_increment against this project's
        existing increments -- the user never manually chooses "new
        increment" vs. "new version" or picks which increment to update.
        """
        project_name = self._current_project_name()
        if not project_name:
            QMessageBox.information(self, "No Project Selected", "Add or select a project first.")
            return

        file_path, _ = QFileDialog.getOpenFileName(self, "Upload File", "", EXCEL_FILE_FILTER)
        if not file_path:
            return

        def on_matched(match_result):
            if match_result.match_type == MatchType.EXACT_MATCH:
                self._upload_as_new_version(project_name, match_result.matched_increment_name, file_path)
            elif match_result.match_type == MatchType.CLOSE_MATCH:
                self._confirm_close_match(project_name, file_path, match_result.matched_increment_name)
            else:
                self._upload_as_new_increment(project_name, file_path)

        def on_error(exc):
            QMessageBox.critical(
                self, "Could Not Read File", f"This file's increment identity couldn't be read:\n\n{exc}"
            )

        run_with_progress(
            self, "Reading file...", self.store.match_upload,
            on_matched, on_error, project_name, file_path,
        )

    def _confirm_close_match(self, project_name: str, file_path: str, matched_increment_name: str):
        box = QMessageBox(self)
        box.setIcon(QMessageBox.Question)
        box.setWindowTitle("Similar Increment Found")
        box.setText(
            f"This looks similar to '{matched_increment_name}' — is this an update to that "
            "increment, or a new one?"
        )
        update_button = box.addButton("Update Existing", QMessageBox.AcceptRole)
        box.addButton("Create New", QMessageBox.RejectRole)
        box.setDefaultButton(update_button)
        box.exec()

        if box.clickedButton() is update_button:
            self._upload_as_new_version(project_name, matched_increment_name, file_path)
        else:
            self._upload_as_new_increment(project_name, file_path)

    def _upload_as_new_version(self, project_name: str, increment_name: str, file_path: str):
        def on_finished(result):
            review = ReviewDialog(increment_name, result, parent=self)
            if review.exec() == QDialog.Accepted:
                # File copy + small JSON write -- fast, no need to background this part.
                self.store.confirm_update(project_name, increment_name, file_path, result)
                self._refresh_increment_table()
                self.on_view_data(project_name, increment_name)
            # Cancel/Discard: nothing changes, we're already back on this page.

        def on_error(exc):
            QMessageBox.critical(self, "Could Not Read File", f"This file couldn't be compared:\n\n{exc}")

        run_with_progress(
            self, "Comparing with current version...", self.store.simulate_comparison,
            on_finished, on_error, project_name, increment_name, file_path,
        )

    def _upload_as_new_increment(self, project_name: str, file_path: str):
        def on_finished(increment):
            self._refresh_increment_table()
            QMessageBox.information(
                self, "Increment Uploaded", f"'{increment.name}' was added as Version 1 for {project_name}."
            )

        def on_error(exc):
            if isinstance(exc, ValueError):
                # e.g. an increment with this file's Record Name already
                # exists in this project -- shouldn't normally happen
                # here (match_upload() already checked), but is still a
                # real error if it does.
                QMessageBox.warning(self, "Could Not Add Increment", str(exc))
            else:
                QMessageBox.critical(self, "Could Not Read File", f"This file couldn't be parsed:\n\n{exc}")

        run_with_progress(
            self, "Reading file...", self.store.add_new_increment,
            on_finished, on_error, project_name, file_path,
        )
