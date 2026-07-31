"""Shown when Check for Updates (ui/pages/home_page.py) finds a newer
release than APP_VERSION. Offers to open the release page in the
system default browser, with a manual-copy fallback -- a read-only,
selectable URL field -- since QDesktopServices.openUrl() has no way to
report back if the default browser is missing, broken, or just not
the one Rey actually wants to use.
"""

from __future__ import annotations

from PySide6.QtCore import QUrl
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import QDialog, QHBoxLayout, QLabel, QLineEdit, QPushButton, QVBoxLayout


class UpdateAvailableDialog(QDialog):
    def __init__(
        self, current_version: str, latest_version: str, release_url: str, is_prerelease: bool, parent=None
    ):
        super().__init__(parent)
        self.setWindowTitle("Update Available")
        self.setMinimumWidth(460)

        layout = QVBoxLayout(self)

        version_label = f"v{latest_version}, pre-release" if is_prerelease else f"v{latest_version}"
        message = QLabel(f"A new version ({version_label}) is available. You're currently on v{current_version}.")
        message.setWordWrap(True)
        layout.addWidget(message)

        layout.addSpacing(10)

        link_label = QLabel("Or copy this link:")
        link_label.setObjectName("hint")
        layout.addWidget(link_label)

        # Read-only rather than disabled -- disabled QLineEdits can't be
        # selected/copied on every platform, which would defeat the
        # point of this fallback. setCursorPosition(0) so a long URL
        # shows its start (the meaningful, distinguishing part) rather
        # than wherever it happens to scroll to by default.
        self.url_edit = QLineEdit(release_url)
        self.url_edit.setReadOnly(True)
        self.url_edit.setCursorPosition(0)
        layout.addWidget(self.url_edit)

        layout.addSpacing(10)

        button_row = QHBoxLayout()
        button_row.addStretch(1)

        # Deliberately does NOT close the dialog -- if the system's
        # default browser is the exact problem this dialog exists to
        # work around, closing immediately on click would strand Rey
        # with nothing to fall back on. "Later" (or the window's own
        # close button) is the only way to dismiss it.
        download_button = QPushButton("Download")
        download_button.setObjectName("primaryButton")
        download_button.clicked.connect(lambda: QDesktopServices.openUrl(QUrl(release_url)))
        button_row.addWidget(download_button)

        later_button = QPushButton("Later")
        later_button.setObjectName("secondaryButton")
        later_button.clicked.connect(self.reject)
        button_row.addWidget(later_button)

        layout.addLayout(button_row)
