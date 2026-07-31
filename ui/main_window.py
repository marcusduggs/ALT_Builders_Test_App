"""Top-level window: a QStackedWidget switching between the Home page
(project bar + increment list), the Data View page, and the Combined
Data View page (read-only multi-increment preview). The review screen
(Part 3) is a modal QDialog launched from HomePage, not a stack page --
it's a deliberate, transient step, not somewhere you navigate "back" from.
"""

from __future__ import annotations

import logging

from PySide6.QtWidgets import QMainWindow, QMessageBox, QStackedWidget

from core import app_logging
from ui.mock_data import MockDataStore, build_combined_view
from ui.pages.combined_data_view_page import CombinedDataViewPage
from ui.pages.data_view_page import DataViewPage
from ui.pages.home_page import HomePage
from ui.workers import run_with_progress

logger = logging.getLogger(__name__)


class MainWindow(QMainWindow):
    def __init__(self, store: MockDataStore | None = None):
        super().__init__()
        self.setWindowTitle("Altamirano Builders TIO Compliance & Reporting")
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
        # get_increment_for_display() doesn't always raise on failure --
        # it returns None if the increment record or this version's
        # file can't be found (see its docstring) -- so on_finished has
        # to check for that itself; a real exception is a SEPARATE path
        # (on_error below), already handled and already logged in full
        # by ui.workers._Worker.run(). This None check used to be
        # missing entirely, which is exactly what let a failed load
        # show the progress dialog and then silently do nothing --
        # ui.pages.data_view_page's version-switch dropdown
        # (_on_version_changed) already had this same check; it just
        # hadn't been applied to this, the INITIAL load, too.
        def on_finished(increment):
            if increment is None:
                logger.warning(
                    "get_increment_for_display returned None for project=%r increment=%r -- "
                    "record or version file missing/unreadable",
                    project_name, increment_name,
                )
                QMessageBox.critical(
                    self,
                    "Could Not Load Data",
                    f'"{increment_name}" could not be found. It may have been deleted, moved, or '
                    "renamed since this list was last refreshed.\n\nReturning to the increments list.",
                )
                self.show_home()
                return
            if self.data_view_page is not None:
                self.stack.removeWidget(self.data_view_page)
                self.data_view_page.deleteLater()

            self.data_view_page = DataViewPage(project_name, increment, self.store, on_back=self.show_home)
            self.stack.addWidget(self.data_view_page)
            self.stack.setCurrentWidget(self.data_view_page)

        def on_error(exc):
            # The exception itself is already logged in full (with
            # traceback) by ui.workers._Worker.run() -- this dialog
            # just needs to tell Rey something useful happened and
            # where to find more detail if support asks for it.
            QMessageBox.critical(
                self,
                "Could Not Load Data",
                f'"{increment_name}" couldn\'t be loaded -- something about its file wasn\'t '
                f"readable.\n\nDetails: {exc}\n\n"
                f"This has been recorded in the app's log file:\n{app_logging.get_log_path()}",
            )

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
            # Unlike show_data_view above, this already converts a
            # get_increment_for_display() None (a missing record/version
            # file) into a raised exception BEFORE build_combined_view()
            # ever sees it -- so on_finished below never receives a
            # combined view built from a hole in the data; on_error
            # handles this the same way it handles any other failure.
            increments = [
                self.store.get_increment_for_display(project_name, name, version=version)
                for name, version in selected
            ]
            missing = [name for (name, _), inc in zip(selected, increments) if inc is None]
            if missing:
                raise ValueError(
                    f"Could not load: {', '.join(missing)} -- may have been deleted, moved, or "
                    "renamed since this list was last refreshed."
                )
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
            # Already logged in full (with traceback) by
            # ui.workers._Worker.run() -- see show_data_view's matching
            # comment above.
            QMessageBox.critical(
                self,
                "Could Not Load Data",
                f"The combined view couldn't be loaded.\n\nDetails: {exc}\n\n"
                f"This has been recorded in the app's log file:\n{app_logging.get_log_path()}",
            )

        run_with_progress(self, "Loading combined data...", load_increments, on_finished, on_error)

    def show_home(self):
        self.home_page._refresh_increment_table()
        self.stack.setCurrentWidget(self.home_page)
