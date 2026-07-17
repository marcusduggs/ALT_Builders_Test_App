"""Single source of truth for the app's displayed version and changelog
summary -- shown in the About dialog (ui/dialogs/about_dialog.py).

APP_VERSION is currently hand-bumped per release (as of the 0.8.0
release, by explicit request -- see git history) rather than left at
the "0.0.0-dev" placeholder this file previously documented.
.github/workflows/release.yml still OVERWRITES this line in its own CI
workspace (never committed back) with the exact tag being released,
immediately before packaging, so a shipped .exe's version can never
drift from the tag regardless of whatever value is committed here --
but since that value is no longer "0.0.0-dev", local/dev and ordinary
main-push CI builds now show the last-released version number too,
not an obvious "unreleased" marker. Remember to bump this by hand on
each release if that distinction matters again.

CHANGELOG_SUMMARY IS still manually maintained -- prose can't be
derived from a tag name -- and is reused verbatim as the GitHub Release
body by that same workflow, so release notes and the About dialog's
changelog can never drift apart from EACH OTHER either.
"""

APP_VERSION = "0.11.0"

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
- Choose where your data is stored
- Check for app updates, and open the built-in help guide, from the \
menu
- Renamed to Altamirano Builders TIO Compliance and Reporting
- Combine multiple increments into one report, with a read-only \
preview before exporting
- See a state's official revision log, this app's own update history, \
and your own notes together on one Changes tab
- View the help guide in a built-in window instead of an external PDF \
viewer
- Added a proper app icon, shown in the title bar, taskbar/dock, and \
Alt-Tab switcher
- Fixed the menu dropdown rendering in the wrong (dark) theme
- Sized the logo and the increments table to read clearly instead of \
looking squeezed or leaving empty space
- Unified button styling app-wide (primary/secondary/destructive) for \
a consistent look
- Added zebra striping and hover highlighting to data tables for \
easier row scanning
- Fixed a bug where the increments table could show duplicate Version/\
Actions controls after navigating back to it
- Combined All Data/Sum Data now show a subtotal after each \
increment's items, in addition to the overall Grand Total
- Added a header with the app logo and title above the Project row
- Increments now sort numerically by increment number (e.g. INC 9 \
before INC 10) instead of an arbitrary order
- Updated the copyright notice's wording (About dialog and LICENSE.txt)
- Fixed the Project dropdown and tooltips rendering in the wrong \
(dark) theme"""
