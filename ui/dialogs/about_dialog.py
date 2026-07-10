"""Simple "About" dialog -- app name, version, and a cumulative
changelog summary (core/app_version.py). No functionality beyond
display; Close is the only action.
"""

from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QDialog, QDialogButtonBox, QLabel, QVBoxLayout

from core.app_version import APP_VERSION, CHANGELOG_SUMMARY

APP_NAME = "Altamirano Builders TIO Compliance and Reporting"


class AboutDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("About")
        self.setMinimumWidth(420)

        layout = QVBoxLayout(self)
        layout.setSpacing(10)

        name_label = QLabel(APP_NAME)
        name_label.setObjectName("pageTitle")

        version_label = QLabel(f"Version {APP_VERSION}")
        version_label.setObjectName("hint")

        changelog_label = QLabel(CHANGELOG_SUMMARY.replace("\n", "<br>"))
        changelog_label.setTextFormat(Qt.RichText)
        changelog_label.setWordWrap(True)

        layout.addWidget(name_label)
        layout.addWidget(version_label)
        layout.addSpacing(6)
        layout.addWidget(changelog_label)

        buttons = QDialogButtonBox(QDialogButtonBox.Close)
        buttons.rejected.connect(self.accept)  # Close button emits rejected(); this is a pure info dialog
        layout.addSpacing(10)
        layout.addWidget(buttons)

        # A word-wrapped QLabel's height-for-width doesn't reliably reach
        # a plain QDialog's initial sizeHint() on some platforms -- without
        # this, the changelog's last couple of lines get silently clipped
        # (verified: it happened here). adjustSize() forces a correct
        # recompute against the actual wrapped content.
        self.adjustSize()
