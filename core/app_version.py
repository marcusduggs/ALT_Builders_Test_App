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

APP_VERSION = "0.12.0"

# Deliberately NOT the usual cumulative feature list this release -- v0.12.0
# exists solely to test the "Check for Updates" flow end-to-end (see
# core/update_check.py) and has no functional changes of its own since
# v0.11.0. Inventing feature bullets here would misrepresent the release in
# both the About dialog and the GitHub Release body (that same text, per
# .github/workflows/release.yml) -- say so plainly instead. The next real
# feature release should go back to a full, cumulative changelog.
CHANGELOG_SUMMARY = "Version bump for update-check testing -- no functional changes since v0.11.0."
