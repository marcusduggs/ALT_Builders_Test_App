"""In-app viewer for the bundled user guide (assets/user_guide.html) --
a QTextBrowser fed via setHtml(), not an external PDF viewer.
"""

from __future__ import annotations

from pathlib import Path

from PySide6.QtWidgets import QDialog, QDialogButtonBox, QTextBrowser, QVBoxLayout


class HelpDialog(QDialog):
    def __init__(self, html_path: Path, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Help")
        self.resize(700, 600)

        layout = QVBoxLayout(self)

        browser = QTextBrowser()
        # Explicit read + setHtml(), not setSource()/QUrl -- setSource()'s
        # relative-path resolution has been unreliable against bundled/
        # frozen paths (get_bundled_path()'s sys._MEIPASS extraction dir),
        # so the content is read directly and handed to setHtml() instead.
        browser.setHtml(html_path.read_text(encoding="utf-8"))
        layout.addWidget(browser)

        buttons = QDialogButtonBox(QDialogButtonBox.Close)
        buttons.rejected.connect(self.accept)  # Close button emits rejected(); this is a pure info dialog
        layout.addWidget(buttons)
