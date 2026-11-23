"""Confirms (or edits) the short label shown for an increment -- used in
TWO places, both via this one dialog class so the sanitization rule,
40-char cap, and duplicate-name warning can never drift between them:

1. BEFORE a brand-new increment is actually saved -- Rey's
   TIO_Changes.xlsx "Sum Data" tab change request #4: "Add the Option
   to change the name of the project before the Upload". Shown once,
   between ui.mock_data.MockDataStore.prepare_new_increment() (the real
   parse, which also computes the derived short name) and
   finalize_new_increment() (the actual save) -- see
   ui.pages.home_page._upload_as_new_increment. `initial_value` is left
   at its default (None), so the field starts pre-filled with the
   derived name itself (there's no earlier custom name to show yet).

2. RENAMING an increment that's already on disk -- see
   ui.pages.home_page._on_rename_increment and
   ui.pages.data_view_page.DataViewPage._on_rename_increment, both of
   which pass `initial_value=` the increment's CURRENT resolved display
   name (custom if set, else the same derived code), plus a
   rename-specific `window_title`/`intro_text`. Clearing the field
   entirely and confirming is the "reset to automatic" action -- see
   value()'s own docstring.

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
    def __init__(
        self,
        record_name: str,
        derived_name: str,
        existing_display_names: list[str],
        initial_value: str | None = None,
        window_title: str = "Name This Increment",
        intro_text: str | None = None,
        parent=None,
    ):
        """record_name: the file's real, full Record Name (shown as
        context, never edited here -- identity is untouched by this
        dialog). derived_name: ui.mock_data.MockDataStore.
        derive_automatic_display_name()'s result (Rey's own "SMR 3D"/
        "HMR 10A" code when the Record Name matches that convention,
        else derive_short_display_name()'s shortened prose name) --
        always shown as the field's PLACEHOLDER (visible only when the
        field is empty), and also used as the field's starting TEXT
        when `initial_value` is None (the upload-flow case: there's no
        earlier custom name to pre-fill with yet).

        initial_value: the increment's CURRENT resolved display name,
        for the rename flow -- pre-fills the field with it instead of
        `derived_name`, so re-opening this dialog on an already-named
        increment shows what it's actually named right now, not a reset
        back to the automatic name.

        existing_display_names: every OTHER increment's current display
        label already in this project (the caller excludes the
        increment being renamed itself, if any) -- used only to show a
        non-blocking warning on a duplicate -- see module docstring on
        why duplicates are allowed, not blocked.
        """
        super().__init__(parent)
        self.setWindowTitle(window_title)
        self.setMinimumWidth(460)
        self._existing_display_names = {n.strip().lower() for n in existing_display_names}

        intro = QLabel(
            intro_text
            or (
                "This short name appears in the Sum Data \"Project Name / Increment #\" column, "
                "the Increments list, and exports. Leave it as-is, or type your own."
            )
        )
        intro.setWordWrap(True)

        self.name_edit = QLineEdit(initial_value if initial_value is not None else derived_name)
        self.name_edit.setPlaceholderText(derived_name)
        self.name_edit.setMaxLength(DISPLAY_NAME_MAX_LENGTH)

        self.warning_label = QLabel()
        self.warning_label.setWordWrap(True)
        self.warning_label.setObjectName("footerHighlight")
        self.warning_label.hide()
        self.name_edit.textChanged.connect(self._refresh_warning)
        self._refresh_warning(self.name_edit.text())

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
        the caller passes this straight to MockDataStore.
        finalize_new_increment() (upload flow) or .rename_increment()
        (rename flow), both of which (via core.project_store.
        _sanitize_display_name()) already treat a blank/whitespace-only
        value as "fall back to the automatic name" -- for a rename,
        that's the explicit "reset to automatic" action; there is no
        separate control for it.
        """
        return self.name_edit.text().strip()
