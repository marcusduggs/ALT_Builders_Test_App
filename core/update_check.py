"""
Manual "Check for Updates" (Option A: opens a browser link to the release
page -- never an automatic in-place update). Compares
core.app_version.APP_VERSION against the highest-versioned GitHub release
-- INCLUDING pre-releases -- via GitHub's public releases API (no auth
needed for a public repo).

Uses GET /releases (the full list), not GET /releases/latest, since
GitHub's own "latest" endpoint deliberately excludes pre-releases by
definition. The full list is returned newest-by-creation-date first, but
this deliberately does NOT just take releases[0] -- it parses every
entry's tag and picks the highest PARSED version, regardless of list
position or prerelease flag, since creation-date order could in
principle differ from version order (e.g. a release published out of
sequence).

check_for_update() is a plain blocking network call with no Qt/UI
dependency -- meant to be run on a background thread via
ui.workers.run_with_progress, same as every other slow call in this app.
It is designed to NEVER raise: any network error, timeout, or unexpected/
malformed response becomes CHECK_FAILED, since a failed check is meant to
be a complete non-event for the user (see ui/pages/home_page.py's
_on_check_for_updates), not an error condition -- this is a low-stakes,
purely informational check, not something worth alarming anyone over.
"""

from __future__ import annotations

import json
import re
import urllib.request
from dataclasses import dataclass
from enum import Enum, auto

from core.app_version import APP_VERSION

RELEASES_LIST_API_URL = "https://api.github.com/repos/marcusduggs/ALT_Builders_Test_App/releases"
REQUEST_TIMEOUT_SECONDS = 10


class UpdateStatus(Enum):
    UP_TO_DATE = auto()
    UPDATE_AVAILABLE = auto()
    CHECK_FAILED = auto()


@dataclass
class UpdateCheckResult:
    status: UpdateStatus
    latest_version: str | None = None  # e.g. "0.7.0" -- only set for UPDATE_AVAILABLE
    release_url: str | None = None  # the release PAGE (.../releases/tag/v0.7.0), not a raw asset link
    is_prerelease: bool = False  # straight from that release's own "prerelease" API field, never inferred from the tag


def _parse_version(tag_name: str) -> tuple[int, ...]:
    """Extracts a numeric (major, minor, patch, ...) tuple from a release
    tag, tolerating whatever punctuation happens to be around the digits
    -- this repo's own tags aren't perfectly consistent (e.g. "v0.6.0"
    vs the stray "v.0.6.0" with an extra dot), so this just pulls out
    every run of digits rather than matching a strict "vX.Y.Z" pattern.
    """
    parts = re.findall(r"\d+", tag_name)
    return tuple(int(p) for p in parts) if parts else (0,)


def _compare_versions(a: tuple[int, ...], b: tuple[int, ...]) -> int:
    """-1 if a<b, 0 if a==b, 1 if a>b.

    Compares as TUPLES OF INTS, not strings -- a plain string comparison
    would put "0.10.0" before "0.9.0" (lexicographic: '1' < '9'), which
    is wrong; (0, 10, 0) > (0, 9, 0) as integer tuples is the correct
    semantic-version ordering. Pads the shorter tuple with zeros first
    so e.g. (0, 7) compares equal to (0, 7, 0).
    """
    length = max(len(a), len(b))
    a_padded = a + (0,) * (length - len(a))
    b_padded = b + (0,) * (length - len(b))
    if a_padded < b_padded:
        return -1
    if a_padded > b_padded:
        return 1
    return 0


def check_for_update() -> UpdateCheckResult:
    """Hits GitHub's public releases list API and compares the HIGHEST
    parsed version across every release (including pre-releases) against
    APP_VERSION. Always returns a result -- never raises.
    """
    try:
        request = urllib.request.Request(
            RELEASES_LIST_API_URL, headers={"Accept": "application/vnd.github+json"}
        )
        with urllib.request.urlopen(request, timeout=REQUEST_TIMEOUT_SECONDS) as response:
            releases = json.loads(response.read().decode("utf-8"))

        if not isinstance(releases, list) or not releases:
            return UpdateCheckResult(status=UpdateStatus.CHECK_FAILED)

        candidates = []
        for release in releases:
            tag_name = release.get("tag_name")
            release_url = release.get("html_url")
            if not tag_name or not release_url:
                continue  # skip a malformed entry rather than failing the whole check
            candidates.append((_parse_version(tag_name), tag_name, release_url, bool(release.get("prerelease"))))
        if not candidates:
            return UpdateCheckResult(status=UpdateStatus.CHECK_FAILED)

        # Highest PARSED version wins, regardless of the list's
        # creation-date order or prerelease flag -- see module docstring.
        latest_parsed, latest_tag, latest_url, latest_is_prerelease = max(candidates, key=lambda c: c[0])

        current = _parse_version(APP_VERSION)
        if _compare_versions(latest_parsed, current) > 0:
            display_version = re.sub(r"^v\.?", "", latest_tag)  # "v0.7.0"/"v.0.7.0" -> "0.7.0"
            return UpdateCheckResult(
                status=UpdateStatus.UPDATE_AVAILABLE,
                latest_version=display_version,
                release_url=latest_url,
                is_prerelease=latest_is_prerelease,
            )
        return UpdateCheckResult(status=UpdateStatus.UP_TO_DATE)
    except Exception:  # noqa: BLE001 -- intentionally broad: this must NEVER raise, see module docstring
        return UpdateCheckResult(status=UpdateStatus.CHECK_FAILED)
