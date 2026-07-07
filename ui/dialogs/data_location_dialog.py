"""One dialog, two callers: the first-launch picker (ui/app.py, before
the main window ever opens) and Settings > Change Data Location
(ui/main_window.py). Same "pick a folder" UI either way -- only the
explanatory text and what the caller does with the result differ.
"""

from __future__ import annotations

from pathlib import Path

from PySide6.QtWidgets import (
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QVBoxLayout,
)

from core.app_config import DATA_DIR_NAME


class DataLocationDialog(QDialog):
    def __init__(self, current_location: Path, is_first_launch: bool, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Choose Data Location" if is_first_launch else "Change Data Location")
        self.setMinimumWidth(520)

        layout = QVBoxLayout(self)

        if is_first_launch:
            explanation = (
                "This is where the app will store your projects, uploaded files, and status -- "
                "everything it keeps is a plain folder on disk, right here.\n\n"
                "The suggested location below works fine for most people. If you'd rather keep "
                "this data somewhere else (an external drive, a synced folder, etc.), use Browse."
            )
        else:
            explanation = (
                "Choose a new location for your projects, uploaded files, and status. "
                "Your existing data will be MOVED there (not copied) -- nothing is duplicated, "
                "and the old location will be empty afterward."
            )
        label = QLabel(explanation)
        label.setWordWrap(True)
        layout.addWidget(label)

        path_row = QHBoxLayout()
        self.path_edit = QLineEdit(str(current_location))
        browse_button = QPushButton("Browse...")
        browse_button.clicked.connect(self._on_browse)
        path_row.addWidget(self.path_edit, stretch=1)
        path_row.addWidget(browse_button)
        layout.addLayout(path_row)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self._on_accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _on_browse(self):
        # getExistingDirectory picks a PARENT folder; the actual data
        # folder is created/moved as a named subfolder of it, so casual
        # browsing to e.g. an external drive doesn't dump loose
        # projects/_deleted folders into it with nothing identifying
        # them as this app's.
        parent_dir = QFileDialog.getExistingDirectory(self, "Choose a Parent Folder", self.path_edit.text())
        if parent_dir:
            self.path_edit.setText(str(Path(parent_dir) / DATA_DIR_NAME))

    def _on_accept(self):
        if not self.path_edit.text().strip():
            self.path_edit.setFocus()
            return
        self.accept()

    def chosen_path(self) -> Path:
        return Path(self.path_edit.text().strip())
