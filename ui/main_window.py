"""Top-level window: a QStackedWidget switching between the Home page
(project bar + increment list) and the Data View page. The review screen
(Part 3) is a modal QDialog launched from HomePage, not a stack page --
it's a deliberate, transient step, not somewhere you navigate "back" from.
"""

from __future__ import annotations

from PySide6.QtWidgets import QDialog, QMainWindow, QMessageBox, QStackedWidget

from core import app_config
from core.project_store import ProjectStore
from ui.dialogs.data_location_dialog import DataLocationDialog
from ui.mock_data import MockDataStore
from ui.pages.data_view_page import DataViewPage
from ui.pages.home_page import HomePage
from ui.workers import run_with_progress


class MainWindow(QMainWindow):
    def __init__(self, store: MockDataStore | None = None):
        super().__init__()
        self.setWindowTitle("TIO Compliance Tracker")
        self.resize(1240, 800)

        self.store = store or MockDataStore()

        self.stack = QStackedWidget()
        self.setCentralWidget(self.stack)

        self.home_page = HomePage(self.store, on_view_data=self.show_data_view)
        self.stack.addWidget(self.home_page)
        self.stack.setCurrentWidget(self.home_page)

        self.data_view_page: DataViewPage | None = None

        self._build_settings_menu()

    def _build_settings_menu(self):
        settings_menu = self.menuBar().addMenu("Settings")
        change_location_action = settings_menu.addAction("Change Data Location...")
        change_location_action.triggered.connect(self._on_change_data_location)

    def _on_change_data_location(self):
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

        # Point the SAME MockDataStore at a fresh ProjectStore over the
        # new location -- no restart needed. HomePage/DataViewPage hold
        # a reference to this one MockDataStore, not a path, so this is
        # the one place that needs updating.
        self.store.store = ProjectStore(base_dir=chosen)
        self.home_page._refresh_project_combo()

        QMessageBox.information(self, "Data Moved", f"Your data is now stored at:\n\n{chosen}")

    def show_data_view(self, project_name: str, increment_name: str):
        def on_finished(increment):
            if increment is None:
                return
            if self.data_view_page is not None:
                self.stack.removeWidget(self.data_view_page)
                self.data_view_page.deleteLater()

            self.data_view_page = DataViewPage(project_name, increment, self.store, on_back=self.show_home)
            self.stack.addWidget(self.data_view_page)
            self.stack.setCurrentWidget(self.data_view_page)

        def on_error(exc):
            QMessageBox.critical(self, "Could Not Load Data", f"This file couldn't be read:\n\n{exc}")

        run_with_progress(
            self, "Loading data...", self.store.get_increment_for_display,
            on_finished, on_error, project_name, increment_name,
        )

    def show_home(self):
        self.home_page._refresh_increment_table()
        self.stack.setCurrentWidget(self.home_page)
