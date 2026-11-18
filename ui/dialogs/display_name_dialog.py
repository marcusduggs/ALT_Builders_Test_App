"""Confirms (or edits) the short label shown for a brand-new increment
before it's actually saved -- Rey's TIO_Changes.xlsx "Sum Data" tab
change request #4: "Add the Option to change the name of the project
before the Upload". Shown once, between ui.mock_data.MockDataStore.
prepare_new_increment() (the real parse, which also computes the
derived short name) and finalize_new_increment() (the actual save) --
see ui.pages.home_page._upload_as_new_increment.

Same QFormLayout/QLineEdit/QDialogButtonBox pattern as
ui.dialogs.project_dialog.ProjectDialog, the closest existing template
for a single-field editable-text dialog in this app.
"""

from __future__ import annotations

from PySide6.QtWidgets import (
    QDialog,
    QDialogButtonBox,
    QFormLayout,
    QLabel,
    QLineEdit,
    QVBoxLayout,
)

from core.project_store import DISPLAY_NAME_MAX_LENGTH


class DisplayNameDialog(QDialog):
    def __init__(self, record_name: str, derived_name: str, existing_display_names: list[str], parent=None):
        """record_name: the file's real, full Record Name (shown as
        context, never edited here -- identity is untouched by this
        dialog). derived_name: ui.mock_data._derive_fallback_display_name()'s
        result (Rey's own "SMR 3D"/"HMR 10A" code when the Record Name
        matches that convention, else derive_short_display_name()'s
        shortened prose name), used to pre-fill the field.
        existing_display_names:
        every OTHER increment's current display label already in this
        project, used only to show a non-blocking warning on a
        duplicate -- see module docstring on why duplicates are allowed,
        not blocked.
        """
        super().__init__(parent)
        self.setWindowTitle("Name This Increment")
        self.setMinimumWidth(460)
        self._existing_display_names = {n.strip().lower() for n in existing_display_names}

        intro = QLabel(
            "This short name appears in the Sum Data \"Project Name / Increment #\" column, "
            "the Increments list, and exports. Leave it as-is, or type your own."
        )
        intro.setWordWrap(True)

        self.name_edit = QLineEdit(derived_name)
        self.name_edit.setMaxLength(DISPLAY_NAME_MAX_LENGTH)

        self.warning_label = QLabel()
        self.warning_label.setWordWrap(True)
        self.warning_label.setObjectName("footerHighlight")
        self.warning_label.hide()
        self.name_edit.textChanged.connect(self._refresh_warning)
        self._refresh_warning(derived_name)

        record_name_label = QLabel(record_name)
        record_name_label.setWordWrap(True)
        record_name_label.setObjectName("hint")

        form = QFormLayout()
        form.addRow("Project Name / Increment #:", self.name_edit)
        form.addRow("Full record name:", record_name_label)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)

        layout = QVBoxLayout(self)
        layout.addWidget(intro)
        layout.addLayout(form)
        layout.addWidget(self.warning_label)
        layout.addWidget(buttons)

    def _refresh_warning(self, text: str) -> None:
        # Allowed, not blocked (see module docstring) -- display_name is
        # purely cosmetic, never used for identity/matching/folder
        # slugs, so two increments sharing one is a visual ambiguity at
        # worst, not a real conflict. Warned here so the user can still
        # choose to make it unique if they want to.
        if text.strip().lower() in self._existing_display_names:
            self.warning_label.setText(
                f"Another increment in this project is already named \"{text.strip()}\". "
                "You can still use it -- this is just a label, not an identifier."
            )
            self.warning_label.show()
        else:
            self.warning_label.hide()

    def value(self) -> str:
        """The confirmed name, or "" if the user cleared the field --
        the caller (ui.pages.home_page) passes this straight to
        MockDataStore.finalize_new_increment(), which (via
        core.project_store._sanitize_display_name()) already treats a
        blank/whitespace-only value as "fall back to the derived name".
        """
        return self.name_edit.text().strip()
