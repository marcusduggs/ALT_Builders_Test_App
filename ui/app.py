"""Entry point. Run from the repo root: `python -m ui.app`"""

from __future__ import annotations

import sys

from PySide6.QtWidgets import QApplication, QDialog

from core import app_config, project_store
from ui.dialogs.data_location_dialog import DataLocationDialog
from ui.main_window import MainWindow
from ui.paths import get_bundled_path


def _ensure_data_location_configured() -> bool:
    """First-launch check: if no data-location config exists yet, show
    the picker BEFORE the main window opens, and write whatever's chosen
    (moving any data already sitting at the historical hardcoded default
    -- e.g. a machine with real data from before this config existed --
    to the newly chosen location, if it differs). Returns False if the
    user cancels, since the app can't proceed without knowing where to
    store its data; the caller should exit cleanly in that case.

    A no-op (returns True immediately) on every launch after the first,
    since read_configured_data_dir() then finds the config already
    written.
    """
    if app_config.read_configured_data_dir() is not None:
        return True

    current_location = project_store.get_default_data_dir()  # runs the legacy migration as a side effect
    dialog = DataLocationDialog(current_location=current_location, is_first_launch=True)
    if dialog.exec() != QDialog.Accepted:
        return False

    chosen = dialog.chosen_path()
    app_config.relocate_data_dir(current_location, chosen)
    app_config.write_configured_data_dir(chosen)
    return True


def main():
    app = QApplication(sys.argv)

    # get_bundled_path(), not Path(__file__).parent: a frozen PyInstaller
    # build doesn't extract plain .py modules as loose files even under
    # sys._MEIPASS, so __file__-relative lookups silently fail there --
    # see ui/paths.py's get_bundled_path docstring.
    style_path = get_bundled_path("ui", "style.qss")
    if style_path.exists():
        app.setStyleSheet(style_path.read_text())

    if not _ensure_data_location_configured():
        sys.exit(0)

    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
