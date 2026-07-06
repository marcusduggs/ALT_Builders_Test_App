"""Add/Update project form. Same dialog serves both -- pass `project` to
pre-fill fields for an update, or None for a fresh Add.
"""

from __future__ import annotations

from PySide6.QtWidgets import (
    QDialog,
    QDialogButtonBox,
    QFormLayout,
    QLineEdit,
    QVBoxLayout,
)


class ProjectDialog(QDialog):
    def __init__(self, project=None, parent=None):
        super().__init__(parent)
        self.is_update = project is not None
        self.setWindowTitle("Update Project" if self.is_update else "Add Project")
        self.setMinimumWidth(420)

        self.name_edit = QLineEdit()

        # Home Folder is purely informational -- always this SPECIFIC
        # project's own folder (core.project_store.ProjectRecord.home_folder,
        # computed fresh from the project's slug, never stored/user-set --
        # see core/project_store.py's module docstring). Disabled, not just
        # read-only, so it visually reads as "this isn't yours to edit."
        # A new project's slug doesn't exist until it's actually created
        # (create_project() generates it), so there's nothing real to show
        # here yet -- the placeholder says so, and the real path appears
        # the next time this project is opened for Update.
        self.folder_edit = QLineEdit()
        self.folder_edit.setEnabled(False)
        if project is not None:
            self.name_edit.setText(project.name)
            self.folder_edit.setText(project.home_folder)
        else:
            self.folder_edit.setPlaceholderText("(assigned automatically once this project is created)")

        form = QFormLayout()
        form.addRow("Project Name:", self.name_edit)
        form.addRow("Home Folder:", self.folder_edit)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self._on_accept)
        buttons.rejected.connect(self.reject)

        layout = QVBoxLayout(self)
        layout.addLayout(form)
        layout.addWidget(buttons)

    def _on_accept(self):
        if not self.name_edit.text().strip():
            self.name_edit.setFocus()
            return
        self.accept()

    def values(self) -> str:
        return self.name_edit.text().strip()
