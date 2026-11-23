"""Single source of truth for the app's displayed version and changelog
summary -- shown in the About dialog (ui/dialogs/about_dialog.py).

APP_VERSION is back to the "0.0.0-dev" placeholder (by explicit
request, reversing the v0.8.0-era decision to hand-bump it per release
-- see git history for that earlier rationale, no longer in effect).
.github/workflows/release.yml OVERWRITES this line in its own CI
workspace (never committed back) with the exact tag being released,
immediately before packaging, so a shipped .exe's version can never
drift from the tag regardless of whatever value is committed here --
and since that committed value IS "0.0.0-dev" again, local/dev and
ordinary main-push CI builds now show an obvious "unreleased" marker
instead of the last-released version number.

CHANGELOG_SUMMARY IS still manually maintained -- prose can't be
derived from a tag name -- and is reused verbatim as the GitHub Release
body by that same workflow, so release notes and the About dialog's
changelog can never drift apart from EACH OTHER either.
"""

APP_VERSION = "0.0.0-dev"

# The usual cumulative feature list -- v0.12.0's single "no functional
# changes" line covered only that one test-only release (see git
# history); every real feature release since has appended to the full
# history below, v0.15.0 included.
#
# NOTE: this is a cumulative HISTORY, so a later entry can supersede an
# earlier one (e.g. the "Combined All Data/Sum Data now show a subtotal"
# entry below predates Sum Data's subtotal/totals rows being removed
# again at the client's request -- both entries are accurate for the
# release they describe).
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
(dark) theme
- Streamlined the increment view to show Sum Data and Changes
- Fixed View Data sometimes appearing to do nothing, with no error \
message, if a file couldn't be loaded
- The app now keeps a log file to help diagnose problems if something \
goes wrong
- Improved reliability of increment naming when a new file is uploaded
- Sum Data's stage columns now show the real stage name and are \
color-coded (red/green) based on whether every item for that stage \
is complete
- The update-available dialog now also shows a copyable link, for \
browsers that don't open automatically
- Fixed a bug where projects with a different number of stages than \
usual (e.g. a narrower-scope increment) could cause a crash or make \
the app appear to do nothing -- the number of stages is now detected \
automatically from each file instead of assumed to always be the same
- Background errors that could previously fail silently now show a \
clear message instead
- Fixed a bug where an item with no VCR marking yet was silently \
excluded from Sum Data's Open/Done/Total counts instead of counting \
as Open, which could make a project look more complete than it \
actually was
- Fixed a crash when combining increments with different numbers of \
stages in View Combined Data's Sum Data tab
- Fixed grid items' Done/Open status to come from the "HCAI/FDD Use" \
column instead of the stage's own "Required" marker, which was \
marking an item Done just because a stage applied to it, not because \
it was actually signed off -- existing projects keep their current \
Done/Open marks as-is; only increments uploaded fresh after this \
update get the corrected status automatically
- Sum Data now has one row per project and stage, with a Project # \
column, instead of one row per item with a column for every stage
- Added Notes and Completion Date columns to the Sum Data export
- Fixed real HMR/SMR project files with very long names, and files \
using an older template, that previously failed to load
- Added filtering to the Sum Data tables, both in the app and as real \
Excel column filters in the exported file
- Fixed several layout issues: the increments table cutting off \
early, narrow columns clipping their own headers, and long project \
names overflowing their display area
- Fixed: files with very long project names failed to upload on \
Windows when the data folder was deeply nested
- Clearer error message when the data folder can't be written to
- A leftover folder from a failed upload no longer affects the \
project list
- Fixed an item that appears on more than one row (the same reference \
number listed several times, for different locations or scopes) having \
only one of those rows' stages tracked -- every row's stages are now \
tracked separately, and ordinary grid items no longer show a VCR mark, \
which belongs only to the milestone (F-Cons Verif) items. Confirmed \
against a real HMR 10A file.
- Milestone (F-Cons Verif) items that have no reference number filled \
in, but DO have at least one discipline sign-off mark, are now listed \
with "?" in place of their Index, instead of being left out of the \
reports entirely. A row with no reference number AND no sign-off mark \
(a per-project boilerplate line, not an actual tracked item) stays \
hidden, same as before. The app also stops reading that sheet at its \
ABBREVIATIONS legend, and skips NOTE lines, so neither is mistaken for \
a real item.
- Reference numbers that aren't plain whole numbers are now kept as \
their own separate items instead of being combined into one -- both \
letter-suffixed numbers (7a, 7b, 7c) and decimal ones (F-1.1, F-4.2)
- Sum Data now identifies each increment by its short project code in \
a "Project Name / Increment #" column -- for example "SMR 3D" or \
"HMR 10A" -- on screen and in every export
- Sum Data's Open, Done, Total and % Complete columns, along with its \
Subtotal and Totals rows, have been removed -- it's now a plain list \
of items and their stage status, matching your own reference sheet
- When you upload a brand-new increment, you can now review and change \
the short name it will be listed under, before it's saved
- You can now rename an increment at any time after uploading it, using \
the new Rename option in the Increments list or in its own data view
- Please note: for HMR 10A, and for any other file where the same \
reference number appears on more than one row, upload the file once as \
a new version after installing this update. That one re-upload \
refreshes its stage tracking using the corrected behavior above. Every \
other file is unaffected and needs no action."""
