"""Top-level window: a QStackedWidget switching between the Home page
(project bar + increment list), the Data View page, and the Combined
Data View page (read-only multi-increment preview). The review screen
(Part 3) is a modal QDialog launched from HomePage, not a stack page --
it's a deliberate, transient step, not somewhere you navigate "back" from.
"""

from __future__ import annotations

from PySide6.QtWidgets import QMainWindow, QMessageBox, QStackedWidget

from ui.mock_data import MockDataStore, build_combined_view
from ui.pages.combined_data_view_page import CombinedDataViewPage
from ui.pages.data_view_page import DataViewPage
from ui.pages.home_page import HomePage
from ui.workers import run_with_progress


class MainWindow(QMainWindow):
    def __init__(self, store: MockDataStore | None = None):
        super().__init__()
        self.setWindowTitle("Altamirano Builders TIO Compliance and Reporting")
        self.resize(1240, 800)

        self.store = store or MockDataStore()

        self.stack = QStackedWidget()
        self.setCentralWidget(self.stack)

        self.home_page = HomePage(
            self.store, on_view_data=self.show_data_view, on_view_combined=self.show_combined_data_view
        )
        self.stack.addWidget(self.home_page)
        self.stack.setCurrentWidget(self.home_page)

        self.data_view_page: DataViewPage | None = None
        self.combined_data_view_page: CombinedDataViewPage | None = None

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

    def show_combined_data_view(self, project_name: str, selected: list[tuple[str, int]]):
        """selected: [(increment_name, version), ...], in on-screen
        order -- exactly what ui.pages.home_page.HomePage._selected_increments()
        returns, and exactly what _on_export_combined_report builds its
        own increments list from too (same selection, same versions).
        """

        def load_increments():
            increments = [
                self.store.get_increment_for_display(project_name, name, version=version)
                for name, version in selected
            ]
            missing = [name for (name, _), inc in zip(selected, increments) if inc is None]
            if missing:
                raise ValueError(f"Could not load: {', '.join(missing)}")
            return build_combined_view(increments)

        def on_finished(view):
            if self.combined_data_view_page is not None:
                self.stack.removeWidget(self.combined_data_view_page)
                self.combined_data_view_page.deleteLater()

            self.combined_data_view_page = CombinedDataViewPage(
                project_name, view, self.store, on_back=self.show_home
            )
            self.stack.addWidget(self.combined_data_view_page)
            self.stack.setCurrentWidget(self.combined_data_view_page)

        def on_error(exc):
            QMessageBox.critical(self, "Could Not Load Data", f"This file couldn't be read:\n\n{exc}")

        run_with_progress(self, "Loading combined data...", load_increments, on_finished, on_error)

    def show_home(self):
        self.home_page._refresh_increment_table()
        self.stack.setCurrentWidget(self.home_page)
