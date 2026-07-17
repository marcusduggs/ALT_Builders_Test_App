"""
Home page: Project bar (Part 1) + per-project Increment list (Part 2) +
the Upload/Review flow (Part 3) that lives on top of it.
"""

from __future__ import annotations

import logging
from functools import partial

from PySide6.QtCore import QUrl, Qt
from PySide6.QtGui import QDesktopServices, QImage, QPixmap
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QDialog,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMenu,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QToolButton,
    QVBoxLayout,
    QWidget,
)

from core import app_config, combined_export, update_check
from core.app_version import APP_VERSION
from core.increment_matcher import MatchType
from core.project_store import ProjectStore
from core.update_check import UpdateStatus
from ui.dialogs.about_dialog import APP_NAME, AboutDialog
from ui.dialogs.data_location_dialog import DataLocationDialog
from ui.dialogs.help_dialog import HelpDialog
from ui.dialogs.project_dialog import ProjectDialog
from ui.dialogs.review_dialog import ReviewDialog
from ui.mock_data import MockDataStore
from ui.paths import get_asset_path, get_bundled_path
from ui.workers import run_with_progress

logger = logging.getLogger(__name__)

EXCEL_FILE_FILTER = "TIO Workbooks (*.xlsm *.xlsx)"
# 72px: large enough to read as a masthead mark (tested against 64/80 too --
# 64 read a little small next to the title's bold 22px, 80 started crowding
# the header's vertical padding) without visually overwhelming the title
# next to it. Replaces the old 48px inline size used when the logo sat in
# the Project row instead of its own header -- see _build_header().
HEADER_LOGO_HEIGHT_PX = 72
INCREMENT_TABLE_MAX_HEIGHT_PX = 320  # ~7-8 rows before a scrollbar kicks in -- see _refresh_increment_table


def _make_white_background_transparent(pixmap: QPixmap, threshold: int = 248) -> QPixmap:
    """logo.png's canvas has a baked-in OPAQUE white background, not a
    transparent one -- verified directly: ~72% of its pixels are exactly
    (255,255,255,255); only a thin rounded-corner trim (~1.5% of pixels)
    is actually alpha=0. That was invisible at the old 48px inline size
    sitting on the Project row's own white card background, but becomes a
    plainly visible white box now that the logo sits on the header's
    light-blue band (see _build_header()) instead. Chroma-keys near-white
    pixels to transparent so the mark itself floats on whatever background
    it's placed on, rather than needing to re-export the source asset.
    """
    image = pixmap.toImage().convertToFormat(QImage.Format_ARGB32)
    for y in range(image.height()):
        for x in range(image.width()):
            color = image.pixelColor(x, y)
            if color.red() >= threshold and color.green() >= threshold and color.blue() >= threshold:
                color.setAlpha(0)
                image.setPixelColor(x, y, color)
    return QPixmap.fromImage(image)


class HomePage(QWidget):
    def __init__(self, store: MockDataStore, on_view_data, on_view_combined, parent=None):
        super().__init__(parent)
        self.store = store
        self.on_view_data = on_view_data  # callback(project_name, increment_name)
        self.on_view_combined = on_view_combined  # callback(project_name, [(increment_name, version), ...])

        # Outer layout has NO margins -- the header masthead needs to sit
        # flush against the window's own top/left/right edges (a light
        # band that stops short of the edges reads as a mistake, not a
        # deliberate masthead). The old 20px page padding moves onto its
        # own `content` widget below instead, so the Project row/
        # Increments section keep their previous inset unchanged.
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)
        outer.addWidget(self._build_header())

        content = QWidget()
        content_layout = QVBoxLayout(content)
        content_layout.setContentsMargins(20, 20, 20, 20)
        content_layout.setSpacing(16)
        content_layout.addWidget(self._build_project_bar())
        content_layout.addWidget(self._build_increment_section())
        # No stretch on the section above -- it now sizes to its own
        # content (see INCREMENT_TABLE_MAX_HEIGHT_PX/_refresh_increment_table),
        # so any space left over past that shows the plain page
        # background (QMainWindow's #f4f5f7) instead of an empty white
        # card stretched to fill the window.
        content_layout.addStretch(1)
        outer.addWidget(content, stretch=1)

        self._refresh_project_combo()

    # ------------------------------------------------------------------
    # header / masthead
    # ------------------------------------------------------------------
    def _build_header(self) -> QFrame:
        """Full-width masthead above the Project row -- see QFrame#appHeader
        in ui/style.qss for the light-blue-tint band + bottom-border
        treatment (reuses the same blue as #primaryButton rather than a
        new color) that gives it visual weight distinct from the plain
        page background below it.
        """
        frame = QFrame()
        frame.setObjectName("appHeader")
        layout = QHBoxLayout(frame)
        layout.setContentsMargins(24, 16, 24, 16)
        layout.setSpacing(16)

        logo_label = self._build_logo_label(HEADER_LOGO_HEIGHT_PX)
        if logo_label is not None:
            layout.addWidget(logo_label)

        title_label = QLabel(APP_NAME)
        title_label.setObjectName("mastheadTitle")
        layout.addWidget(title_label)

        layout.addStretch(1)
        return frame

    # ------------------------------------------------------------------
    # Part 1: project bar
    # ------------------------------------------------------------------
    def _build_project_bar(self) -> QFrame:
        frame = QFrame()
        frame.setObjectName("card")
        layout = QHBoxLayout(frame)
        layout.setContentsMargins(16, 14, 16, 14)
        layout.setSpacing(12)

        layout.addWidget(self._build_hamburger_button())
        layout.addSpacing(8)

        label = QLabel("Project:")
        label.setObjectName("sectionTitle")

        self.project_combo = QComboBox()
        self.project_combo.setMinimumWidth(320)
        self.project_combo.currentIndexChanged.connect(self._on_project_changed)

        add_button = QPushButton("Add")
        add_button.setObjectName("secondaryButton")
        update_button = QPushButton("Update")
        update_button.setObjectName("secondaryButton")
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

    def _build_logo_label(self, height_px: int) -> QLabel | None:
        """Loads assets/logo.png (resolved relative to the app's own
        location, not the cwd -- see ui/paths.py) and returns it scaled to
        `height_px` tall. Returns None if the file is missing/unreadable so
        the header just renders without it instead of crashing or showing
        a broken-image icon.

        scaledToHeight() preserves aspect ratio correctly (verified, not a
        scaling bug); logo.png is a tall, narrow (portrait) mark, so a
        short fixed height alone necessarily renders a thin result -- e.g.
        ~28px wide at 36px tall, too small to read its own detail clearly.
        See HEADER_LOGO_HEIGHT_PX's own comment for how the current height
        was picked.
        """
        logo_path = get_asset_path("logo.png")
        pixmap = QPixmap(str(logo_path))
        if pixmap.isNull():
            logger.warning("Could not load logo image at %s; skipping.", logo_path)
            return None

        label = QLabel()
        label.setPixmap(
            _make_white_background_transparent(pixmap).scaledToHeight(
                height_px, Qt.SmoothTransformation
            )
        )
        return label

    # ------------------------------------------------------------------
    # hamburger menu: About / Help / Check for Updates / Move App Data / Exit
    # ------------------------------------------------------------------
    def _build_hamburger_button(self) -> QToolButton:
        button = QToolButton()
        button.setText("☰")
        button.setToolTip("Menu")
        button.setPopupMode(QToolButton.InstantPopup)
        button.setAutoRaise(True)

        menu = QMenu(button)
        menu.addAction("About", self._on_about)
        menu.addAction("Help", self._on_help)
        menu.addSeparator()
        menu.addAction("Check for Updates", self._on_check_for_updates)
        menu.addAction("Move App Data", self._on_move_app_data)
        menu.addSeparator()
        menu.addAction("Exit", self._on_exit)
        button.setMenu(menu)
        return button

    def _on_about(self):
        AboutDialog(parent=self).exec()

    def _on_help(self):
        """Opens assets/user_guide.html in an in-app QTextBrowser dialog
        -- same bundled-resource resolution logo.png/style.qss already
        use (get_bundled_path(), which resolves correctly whether running
        as a script or a frozen PyInstaller build -- see ui/paths.py).
        """
        html_path = get_bundled_path("assets", "user_guide.html")
        if not html_path.exists():
            QMessageBox.warning(
                self, "Help File Not Found", f"The help file couldn't be found at:\n\n{html_path}"
            )
            return
        HelpDialog(html_path, parent=self).exec()

    def _on_check_for_updates(self):
        """Option A: a manual check that opens a browser link, never an
        automatic in-place update. Runs on a background thread (same
        run_with_progress pattern as every other slow call in this app,
        even though a GitHub API call is normally fast) so a slow/stalled
        connection never freezes the UI.
        """
        run_with_progress(
            self, "Checking for updates...", update_check.check_for_update,
            self._on_update_check_finished, self._on_update_check_error,
        )

    def _on_update_check_finished(self, result):
        if result.status is UpdateStatus.UPDATE_AVAILABLE:
            self._show_update_available_dialog(result)
        elif result.status is UpdateStatus.UP_TO_DATE:
            QMessageBox.information(self, "Up to Date", f"You're on the latest version (v{APP_VERSION}).")
        else:  # CHECK_FAILED -- deliberately low-key, not an error dialog (see core/update_check.py)
            QMessageBox.information(
                self, "Update Check", "Couldn't check for updates right now. Try again later."
            )

    def _on_update_check_error(self, _exc: Exception):
        # check_for_update() is designed to never raise (see its
        # docstring) -- this exists only as the same non-alarming
        # fallback in case something truly unexpected still slips
        # through run_with_progress's error path.
        QMessageBox.information(self, "Update Check", "Couldn't check for updates right now. Try again later.")

    def _show_update_available_dialog(self, result):
        version_label = f"v{result.latest_version}, pre-release" if result.is_prerelease else f"v{result.latest_version}"
        box = QMessageBox(self)
        box.setIcon(QMessageBox.Information)
        box.setWindowTitle("Update Available")
        box.setText(
            f"A new version ({version_label}) is available. You're currently on v{APP_VERSION}."
        )
        download_button = box.addButton("Download", QMessageBox.AcceptRole)
        box.addButton("Later", QMessageBox.RejectRole)
        box.exec()
        if box.clickedButton() is download_button:
            # The release PAGE (has context: notes, assets, etc.), not a
            # raw asset link -- Rey should land somewhere with context,
            # not a bare file download.
            QDesktopServices.openUrl(QUrl(result.release_url))

    def _on_exit(self):
        # Closes the top-level window exactly as clicking its own close
        # button would -- goes through the normal Qt close event rather
        # than bypassing it with QApplication.quit(), so any cleanup a
        # closeEvent might do in the future still runs.
        self.window().close()

    def _on_move_app_data(self):
        """Formerly "Change Data Location..." under a standalone Settings
        menu -- relocated into the hamburger menu (that menu was the
        only thing living there, so it was removed entirely) and
        relabeled. Same confirm/move behavior as before, just moved:
        HomePage already holds the one MockDataStore instance directly,
        so this needs no reference back to MainWindow at all.
        """
        current_location = self.store.store.base_dir
        dialog = DataLocationDialog(current_location=current_location, is_first_launch=False)
        if dialog.exec() != QDialog.Accepted:
            return

        chosen = dialog.chosen_path()
        if chosen == current_location:
            return  # nothing actually changed

        answer = QMessageBox.question(
            self,
            "Move Data?",
            (
                f"This will move your existing data to:\n\n{chosen}\n\n"
                "This may take a moment for large amounts of data. Continue?"
            ),
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if answer != QMessageBox.Yes:
            return

        try:
            app_config.relocate_data_dir(current_location, chosen)
        except FileExistsError as exc:
            QMessageBox.critical(self, "Could Not Move Data", str(exc))
            return
        app_config.write_configured_data_dir(chosen)

        # Point THIS SAME MockDataStore at a fresh ProjectStore over the
        # new location -- no restart needed.
        self.store.store = ProjectStore(base_dir=chosen)
        self._refresh_project_combo()

        QMessageBox.information(self, "Data Moved", f"Your data is now stored at:\n\n{chosen}")

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

        self.select_all_checkbox = QCheckBox("Select All")
        self.select_all_checkbox.toggled.connect(self._on_select_all_toggled)

        self.view_combined_button = QPushButton("View Combined Data")
        self.view_combined_button.setObjectName("secondaryButton")
        self.view_combined_button.setEnabled(False)
        self.view_combined_button.clicked.connect(self._on_view_combined_data)

        self.export_combined_button = QPushButton("Export Combined Report")
        self.export_combined_button.setObjectName("secondaryButton")
        self.export_combined_button.setEnabled(False)
        self.export_combined_button.clicked.connect(self._on_export_combined_report)

        upload_button = QPushButton("Upload File")
        upload_button.setObjectName("primaryButton")
        upload_button.clicked.connect(self._on_upload_file)

        header_row.addWidget(title)
        header_row.addStretch(1)
        header_row.addWidget(self.select_all_checkbox)
        header_row.addWidget(self.view_combined_button)
        header_row.addWidget(self.export_combined_button)
        header_row.addWidget(upload_button)
        layout.addLayout(header_row)

        self.increment_table = QTableWidget(0, 5)
        self.increment_table.setHorizontalHeaderLabels(["", "Increment", "Version", "Last Updated", "Actions"])
        self.increment_table.verticalHeader().hide()
        self.increment_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.increment_table.setSelectionMode(QTableWidget.NoSelection)
        self.increment_table.setAlternatingRowColors(True)
        header = self.increment_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(4, QHeaderView.ResizeToContents)
        layout.addWidget(self.increment_table)

        self.empty_state_label = QLabel("No increments yet for this project. Use “Upload File” to add one.")
        self.empty_state_label.setObjectName("hint")
        self.empty_state_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.empty_state_label)

        # One entry per row, in on-screen order: {"increment_name",
        # "checkbox", "version_combo", "last_updated_item",
        # "dates_by_version"}. Export Combined Report reads this
        # directly at click time rather than re-deriving anything from
        # the table's cell widgets.
        self._increment_row_widgets: list[dict] = []

        return frame

    def _refresh_increment_table(self):
        project_name = self._current_project_name()
        project = self.store.get_project(project_name) if project_name else None
        increments = project.increments if project else []

        # Explicitly tear down any PREVIOUS cell widgets before
        # repopulating -- this table can legitimately be refreshed more
        # than once for the same rows (e.g. MainWindow.show_home()'s own
        # direct call when navigating back from Data View, or re-selecting
        # an already-selected project), and setCellWidget()'s documented
        # "replaces and deletes the old widget" behavior does NOT reliably
        # hold in practice: confirmed by direct reproduction that calling
        # setCellWidget() again at the same (row, col) leaves the previous
        # widget alive as an orphaned child of the viewport, invisible at
        # first (old/new positions nearly coincide) but revealed as a
        # visually duplicated row the moment a later window resize moves
        # the Stretch column and only the NEW widgets follow along.
        # removeCellWidget() alone isn't enough either -- per its own Qt
        # docs, it only stops the table from TRACKING the widget as that
        # cell's editor and hands ownership back to the caller; nothing
        # deletes it automatically. deleteLater() on the retrieved widget
        # is what actually destroys the orphan.
        for row in range(self.increment_table.rowCount()):
            for col in (0, 2, 4):  # the only columns ever given a cell widget -- see below
                old_widget = self.increment_table.cellWidget(row, col)
                if old_widget is not None:
                    self.increment_table.removeCellWidget(row, col)
                    old_widget.deleteLater()

        self._increment_row_widgets = []
        self.increment_table.setRowCount(len(increments))
        for row, increment in enumerate(increments):
            checkbox = QCheckBox()
            checkbox.toggled.connect(partial(self._on_row_checkbox_toggled, row))
            checkbox_container = QWidget()
            checkbox_layout = QHBoxLayout(checkbox_container)
            checkbox_layout.setContentsMargins(0, 0, 0, 0)
            checkbox_layout.setAlignment(Qt.AlignCenter)
            checkbox_layout.addWidget(checkbox)
            self.increment_table.setCellWidget(row, 0, checkbox_container)

            self.increment_table.setItem(row, 1, QTableWidgetItem(increment.name))

            version_combo = QComboBox()
            version_records = self.store.list_versions(project_name, increment.name)
            dates_by_version = {v.version: v.uploaded_date for v in version_records}
            for v in version_records:
                version_combo.addItem(f"v{v.version} ({v.uploaded_date})", userData=v.version)
            latest_index = version_combo.findData(increment.version)
            if latest_index >= 0:
                version_combo.setCurrentIndex(latest_index)
            version_combo.setEnabled(False)  # only interactable once this row's checkbox is checked
            self.increment_table.setCellWidget(row, 2, version_combo)

            last_updated_item = QTableWidgetItem(increment.last_updated)
            self.increment_table.setItem(row, 3, last_updated_item)
            version_combo.currentIndexChanged.connect(partial(self._on_row_version_changed, row))

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
            self.increment_table.setCellWidget(row, 4, actions)

            self._increment_row_widgets.append(
                {
                    "increment_name": increment.name,
                    "checkbox": checkbox,
                    "version_combo": version_combo,
                    "last_updated_item": last_updated_item,
                    "dates_by_version": dates_by_version,
                }
            )

        self.increment_table.setVisible(bool(increments))
        self._resize_increment_table_to_content()
        self.empty_state_label.setVisible(not increments)
        self.select_all_checkbox.blockSignals(True)
        self.select_all_checkbox.setChecked(False)
        self.select_all_checkbox.blockSignals(False)
        self._refresh_combined_buttons()

    def _resize_increment_table_to_content(self):
        """Caps the table's height to its actual content (header + real
        rows) instead of always stretching to fill the page -- with few
        increments, that used to leave a large empty white area below
        the last row (most visible when the window is maximized). Above
        INCREMENT_TABLE_MAX_HEIGHT_PX, the table keeps its own normal
        internal scrollbar (a QTableWidget already has one available;
        nothing extra to wire up) instead of growing further. Recomputed
        every refresh since row count changes each time.
        """
        header_height = self.increment_table.horizontalHeader().height()
        rows_height = sum(
            self.increment_table.rowHeight(row) for row in range(self.increment_table.rowCount())
        )
        border_allowance = 2 * self.increment_table.frameWidth() + 2
        natural_height = header_height + rows_height + border_allowance
        self.increment_table.setMaximumHeight(min(natural_height, INCREMENT_TABLE_MAX_HEIGHT_PX))

    # ------------------------------------------------------------------
    # Part 2b: multi-increment selection + combined view/export
    # ------------------------------------------------------------------
    def _on_select_all_toggled(self, checked: bool):
        for row_widgets in self._increment_row_widgets:
            row_widgets["checkbox"].setChecked(checked)

    def _on_row_checkbox_toggled(self, row: int, checked: bool):
        self._increment_row_widgets[row]["version_combo"].setEnabled(checked)
        self._refresh_combined_buttons()

    def _on_row_version_changed(self, row: int, _index: int):
        widgets = self._increment_row_widgets[row]
        selected_version = widgets["version_combo"].currentData()
        date = widgets["dates_by_version"].get(selected_version, "")
        widgets["last_updated_item"].setText(date)

    def _refresh_combined_buttons(self):
        any_checked = any(w["checkbox"].isChecked() for w in self._increment_row_widgets)
        self.view_combined_button.setEnabled(any_checked)
        self.export_combined_button.setEnabled(any_checked)

    def _selected_increments(self) -> list[tuple[str, int]]:
        """[(increment_name, version), ...] for every checked row, in
        on-screen order (top to bottom) -- NOT the order checkboxes were
        clicked -- simpler, and matches what the user is already looking
        at. Shared by View Combined Data and Export Combined Report, so
        both always agree on exactly which increments/versions are
        included.
        """
        return [
            (w["increment_name"], w["version_combo"].currentData())
            for w in self._increment_row_widgets
            if w["checkbox"].isChecked()
        ]

    def _on_view_combined_data(self):
        project_name = self._current_project_name()
        selected = self._selected_increments()
        if not selected:
            return  # button should be disabled anyway; guard regardless
        self.on_view_combined(project_name, selected)

    def _on_export_combined_report(self):
        project_name = self._current_project_name()
        selected = self._selected_increments()
        if not selected:
            return  # button should be disabled anyway; guard regardless

        default_name = combined_export.default_combined_filename(project_name, len(selected))
        path, _ = QFileDialog.getSaveFileName(self, "Export Combined Report", default_name, "Excel Workbook (*.xlsx)")
        if not path:
            return

        def do_export():
            increments = [
                self.store.get_increment_for_display(project_name, name, version=version)
                for name, version in selected
            ]
            missing = [name for (name, _), inc in zip(selected, increments) if inc is None]
            if missing:
                raise ValueError(f"Could not load: {', '.join(missing)}")
            combined_export.export_combined_report(increments, path)

        def on_finished(_result):
            QMessageBox.information(self, "Export Complete", f"Saved to:\n\n{path}")

        def on_error(exc):
            QMessageBox.critical(self, "Could Not Export", f"This file couldn't be saved:\n\n{exc}")

        run_with_progress(self, "Exporting combined report...", do_export, on_finished, on_error)

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
