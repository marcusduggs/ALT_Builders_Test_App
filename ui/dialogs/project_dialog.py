"""Add/Update project form. Same dialog serves both -- pass `project` to
pre-fill fields for an update, or None for a fresh Add.
"""

from __future__ import annotations

from PySide6.QtWidgets import (
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFormLayout,
    QHBoxLayout,
    QLineEdit,
    QPushButton,
    QVBoxLayout,
)


class ProjectDialog(QDialog):
    def __init__(self, project=None, parent=None):
        super().__init__(parent)
        self.is_update = project is not None
        self.setWindowTitle("Update Project" if self.is_update else "Add Project")
        self.setMinimumWidth(420)

        self.name_edit = QLineEdit()
        self.folder_edit = QLineEdit()
        browse_button = QPushButton("Browse...")
        browse_button.clicked.connect(self._browse_folder)

        folder_row = QHBoxLayout()
        folder_row.addWidget(self.folder_edit)
        folder_row.addWidget(browse_button)

        if project is not None:
            self.name_edit.setText(project.name)
            self.folder_edit.setText(project.home_folder)

        form = QFormLayout()
        form.addRow("Project Name:", self.name_edit)
        form.addRow("Home Folder:", folder_row)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self._on_accept)
        buttons.rejected.connect(self.reject)

        layout = QVBoxLayout(self)
        layout.addLayout(form)
        layout.addWidget(buttons)

    def _browse_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Home Folder", self.folder_edit.text())
        if folder:
            self.folder_edit.setText(folder)

    def _on_accept(self):
        if not self.name_edit.text().strip():
            self.name_edit.setFocus()
            return
        self.accept()

    def values(self) -> tuple[str, str]:
        return (
            self.name_edit.text().strip(),
            self.folder_edit.text().strip(),
        )
