"""Entry point. Run from the repo root: `python -m ui.app`"""

from __future__ import annotations

import sys

from PySide6.QtWidgets import QApplication

from ui.main_window import MainWindow
from ui.paths import get_bundled_path


def main():
    app = QApplication(sys.argv)

    # get_bundled_path(), not Path(__file__).parent: a frozen PyInstaller
    # build doesn't extract plain .py modules as loose files even under
    # sys._MEIPASS, so __file__-relative lookups silently fail there --
    # see ui/paths.py's get_bundled_path docstring.
    style_path = get_bundled_path("ui", "style.qss")
    if style_path.exists():
        app.setStyleSheet(style_path.read_text())

    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
