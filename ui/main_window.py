"""Top-level window: a QStackedWidget switching between the Home page
(project bar + increment list) and the Data View page. The review screen
(Part 3) is a modal QDialog launched from HomePage, not a stack page --
it's a deliberate, transient step, not somewhere you navigate "back" from.
"""

from __future__ import annotations

from PySide6.QtWidgets import QMainWindow, QMessageBox, QStackedWidget

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
