"""Single source of truth for the app's displayed version and changelog
summary -- shown in the About dialog (ui/dialogs/about_dialog.py).

APP_VERSION below is a PLACEHOLDER for local/dev runs and ordinary
main-push CI builds -- it is NOT what a released build shows.
.github/workflows/release.yml overwrites this line in its own CI
workspace (never committed back) with the exact tag being released,
immediately before packaging, so a shipped .exe's version can never
drift from the tag that produced it -- there is no longer a
hand-maintained release-version string to forget to bump. This
placeholder staying "0.0.0-dev" is expected and correct; it just means
"this isn't a tagged release build."

CHANGELOG_SUMMARY IS still manually maintained -- prose can't be
derived from a tag name -- and is reused verbatim as the GitHub Release
body by that same workflow, so release notes and the About dialog's
changelog can never drift apart from EACH OTHER either.
"""

APP_VERSION = "0.0.0-dev"

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
menu"""
