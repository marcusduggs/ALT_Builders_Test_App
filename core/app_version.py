"""Single source of truth for the app's displayed version and changelog
summary -- shown in the About dialog (ui/dialogs/about_dialog.py).

Manually updated each release, same as the GitHub release notes already
are -- not auto-generated from git history. CHANGELOG_SUMMARY is a
cumulative, plain-language list of what the app can do (matching the
User Guide's tone), not a raw diff of the latest release.
"""

APP_VERSION = "0.6.0"

CHANGELOG_SUMMARY = """\
- Track progress per stage with Done/Open status
- Review changes before accepting an updated file, including new items \
and updated values
- See stage, status, and progress totals at a glance
- Automatically detect whether an upload is a new increment or an \
update to an existing one
- Keep every previous version of an increment, and view any of them at \
any time
- Export All Data, Sum Data, and Report to Excel
- Choose where your data is stored"""
