"""
Minimal local, file-based storage for projects/increments/status/file
versions. No database -- matches the app's goal of being a fully offline,
zero-setup local tool. Every uploaded file is kept forever (versioned,
never overwritten), so there's always an audit trail of exactly what the
state sent and when.

Default data directory: wherever the user configured via the first-launch
picker or Settings > Change Data Location (see core/app_config.py) --
falling back to ~/AltamiranoBuildersAppData (auto-migrated in place from
~/SubmissionAppData -- this project's old working name -- if found; see
get_default_data_dir() below) if nothing has been configured yet.

Layout:

    <base_dir>/
      projects/
        <project-slug>/
          project.json                  {name} -- home_folder is NOT stored: it's
                                         purely informational, always computed
                                         fresh as this project's own folder
                                         (<base_dir>/projects/<slug>/), never
                                         user-editable and never at risk of
                                         drifting stale if base_dir itself is
                                         ever renamed (see _migrate_legacy_data_dir)
          increments/
            <increment-slug>/
              increment.json            {name, version, last_updated, display_name}
                                         -- display_name is optional (None/absent for
                                         any increment created before this field
                                         existed, or left blank at upload time); see
                                         IncrementRecord.display_name
              status.json               {index: {stage_number: "Done"|"Open"}}
                                         e.g. {"B-C1": {"5": "Done", "12": "Open"}}
                                         -- JSON object keys are always strings, so
                                         stage numbers are stored as string digits;
                                         load_status()/save_status() convert to/from
                                         int transparently, so every OTHER module in
                                         this codebase deals in int stage numbers.
              change_history.json       [{timestamp, old_version, new_version,
                                         added_items, removed_items, column_anomalies,
                                         value_changed_items}, ...] -- append-only, one
                                         entry per CONFIRMED update (never at review
                                         time -- a review the user backs out of leaves
                                         no trace here). Oldest first, same as
                                         list_version_files(); see
                                         append_change_history_entry()/
                                         load_change_history() and
                                         ui.mock_data.MockDataStore.confirm_update(),
                                         which builds each entry from the SAME
                                         ComparisonResult the review screen already
                                         computed and showed the user, not a
                                         recomputation.
              comments.json              [{id, timestamp, text}, ...] -- free-text
                                         notes a user of this app typed directly onto
                                         the Changes tab, not derived from any file.
                                         Append-only like change_history.json: removed
                                         via delete_comment() (an explicit action), not
                                         in-place edits. See add_comment()/
                                         load_comments()/delete_comment().
              files/
                v1_2026-05-10.xlsm
                v2_2026-06-26.xlsm       <- current version = highest v#
      _deleted/
        <timestamp>_<project-slug>/     <- delete_project() moves the whole
                                            project folder here, untouched,
                                            instead of removing it -- manual
                                            recovery only, no restore UI

An increment's slug (its folder name) is generated once, from its name, at
creation time, and never changes afterward. Coupling a stable-forever
filesystem path to a display name a user can freely edit is exactly the
identity-vs-position confusion this codebase has deliberately avoided
elsewhere (see core/excel_reader.py's "PARSING STRATEGY" section).

A PROJECT's slug is the one deliberate exception: since Home Folder is a
live, user-visible path computed from the slug (<base_dir>/projects/<slug>/,
see the module docstring above), update_project() renaming the folder to
match a new name is exactly what keeps that display honest -- a project
whose folder still read like its pre-rename name would look broken, not
reassuring. Uses the exact same _unique_slug() collision strategy
create_project() uses (excluding the project's OWN current slug from the
"taken" set, so renaming to a name that still slugifies the same is a
no-op, not a spurious collision against itself); skips the filesystem
rename entirely when the new slug equals the old one. Every list/get/update
call still matches on the display name stored *inside* project.json /
increment.json, never on the slug -- callers never see or need to know
slugs exist, this is purely an internal consequence of the rename.
"""

from __future__ import annotations

import hashlib
import json
import logging
import re
import shutil
import uuid
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any

from core import app_config

logger = logging.getLogger(__name__)

DELETED_DIRNAME = "_deleted"


def _migrate_legacy_data_dir(target_dir: Path, home_dir: Path | None = None) -> None:
    """One-time migration: if ~/SubmissionAppData (this project's old
    working name, from before it even had an AltamiranoBuildersAppData
    default, let alone a configurable location) still exists and
    target_dir doesn't yet, rename the folder in place -- same parent
    directory, so this is an instant filesystem rename, not a copy, and
    every project/increment/status file underneath it moves untouched
    along with it.

    If target_dir already exists, this does nothing at all: never
    merges, never deletes -- there's no path by which existing data
    could be lost to an edge case here. Safe to call on every startup:
    idempotent, since the second call onward always finds target_dir
    already in place and no-ops. Only ever called with
    app_config.default_data_dir_suggestion() as target_dir -- this is
    strictly about that historical hardcoded location, not wherever the
    user might have since configured (see get_default_data_dir()).

    home_dir defaults to the real Path.home() in production; tests pass
    an explicit temp dir instead so this never resolves to (let alone
    touches) a real user's actual home directory.
    """
    legacy_dir = (home_dir or Path.home()) / app_config.LEGACY_DATA_DIR_NAME
    if legacy_dir.exists() and not target_dir.exists():
        legacy_dir.rename(target_dir)
        print(f"Migrated existing data from {app_config.LEGACY_DATA_DIR_NAME} to {target_dir}")


def get_default_data_dir(config_path: Path | None = None, home_dir: Path | None = None) -> Path:
    """The data folder ProjectStore() uses when no base_dir is given --
    whatever the user configured via the first-launch picker or Settings
    > Change Data Location (core.app_config.read_configured_data_dir()),
    or the historical hardcoded ~/AltamiranoBuildersAppData location
    (running the legacy SubmissionAppData migration first) for any
    caller that constructs a ProjectStore before that config exists --
    real app startup always writes it first, via ui/app.py's
    first-launch check, so this fallback only matters for other/direct
    callers (a script, for instance).

    config_path/home_dir are production no-ops (both default to the
    real config file / real Path.home()) -- present purely so tests can
    override BOTH and never resolve, let alone touch, anything real.
    """
    configured = app_config.read_configured_data_dir(config_path)
    if configured is not None:
        return configured
    suggestion = app_config.default_data_dir_suggestion(home_dir)
    _migrate_legacy_data_dir(suggestion, home_dir)
    return suggestion


@dataclass
class ProjectRecord:
    name: str
    home_folder: str
    slug: str


@dataclass
class IncrementRecord:
    name: str
    version: int
    last_updated: str
    slug: str
    # A user-editable display label, set once at upload time (see
    # _sanitize_display_name() and ProjectStore.create_increment()) --
    # None for every increment created before this field existed, or
    # for one where the user left the "Project Name / Increment #"
    # field blank at upload (both fall back to
    # core.excel_reader.derive_short_display_name(name) wherever this
    # is actually shown -- see ui.mock_data.Increment.display_name).
    # NEVER fed into _slugify_increment() or anything else that touches
    # the on-disk folder slug -- purely a display label, carried
    # forward unchanged on every version bump (see save_new_version()).
    display_name: str | None = None


@dataclass
class VersionRecord:
    version: int
    uploaded_date: str


def _slugify(name: str) -> str:
    slug = re.sub(r"[^a-z0-9]+", "-", name.strip().lower()).strip("-")
    return slug or "untitled"


# Characters illegal in a Windows filename, plus control characters --
# deliberately conservative (a user-typed display label has no reason
# to need any of these) rather than an exhaustive "what Excel/openpyxl
# will actually accept" check. This exists purely so a display_name
# stays safe if it's ever used to SUGGEST a filename somewhere (e.g. an
# export's default save-as name) -- it is NEVER passed to
# _slugify_increment() or anything else that computes the on-disk
# folder slug, so it has no bearing on the Windows MAX_PATH budget
# logic there either way; this is a separate, purely cosmetic safety
# net, not a second identity concern.
_DISPLAY_NAME_ILLEGAL_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
DISPLAY_NAME_MAX_LENGTH = 40


def _sanitize_display_name(value: str | None) -> str | None:
    """Trims whitespace, collapses internal whitespace runs, strips
    filename-illegal/control characters, and caps length at
    DISPLAY_NAME_MAX_LENGTH. Returns None (not "") for a blank/
    whitespace-only input, matching "the user left it blank -> fall
    back to the derived name" (see IncrementRecord.display_name).
    """
    if value is None:
        return None
    cleaned = _DISPLAY_NAME_ILLEGAL_CHARS.sub("", value)
    cleaned = " ".join(cleaned.split())  # trim + collapse internal whitespace
    cleaned = cleaned[:DISPLAY_NAME_MAX_LENGTH].rstrip()
    return cleaned or None


# Increment directory names are derived from core.excel_reader.get_record_name()'s
# FULL free-text Record Name -- for most real files a short title, but
# for some (see tests/fixtures/hmr_smr_projects -- real HMR/SMR TIO
# files) the entire "Scope of Project" paragraph, 300-550+ characters.
# A raw _slugify() of that text alone can exceed the OS's 255-byte
# single-path-component limit and crash create_increment() with
# OSError: File name too long (reproduced on 5 of 13 real fixture
# files). This is an ON-DISK STORAGE concern only: core.increment_matcher's
# identity matching still compares the FULL, untruncated Record Name
# text, exactly as before. NOT applied to create_project()'s own
# _slugify() call -- project names are short, user-typed text, not
# derived from a long file-sourced paragraph, so they were never at
# risk of this.
#
# *** The cap below is a ceiling, not THE answer -- see
# _increment_slug_max_length() for why a single fixed number can't be
# right for every installation ***
# A 100-character cap (sized for macOS's 255-byte single-PATH-COMPONENT
# limit) was confirmed, on a real client Windows machine, to still
# produce a TOTAL PATH over Windows' separate ~259-character MAX_PATH
# limit -- which caps the whole path, not just one component -- once a
# deep enough base data directory (e.g. several folders into a
# OneDrive sync target) is added on top of it: a 114-character slug
# plus a 127-character prefix (base dir + "projects/<slug>/increments")
# plus "/increment.json.tmp" came to 261 characters, 2 over the
# ~259-character point past which Windows refuses to create the file
# (confirmed directly: the 242-character DIRECTORY create succeeded;
# only the file inside it failed). _increment_slug_max_length()
# computes the REAL, installation-specific budget at runtime instead
# (bounded by the separate, FIXED _FIXED_INCREMENT_SLUG_LENGTH_CAP
# below, not by this constant -- see that constant's own docstring for
# why a fixed, machine-independent ceiling matters on top of the
# per-installation one). This constant now only remains as
# _slugify_increment()'s own default `max_length` for a caller that
# doesn't go through the dynamic path at all.
_MAX_INCREMENT_SLUG_LENGTH = 100

# The longest relative path the app ever writes INSIDE an increment
# folder, relative to the increment folder itself -- confirmed by
# listing every file create_increment()/save_new_version()/
# save_status()/comments functions write: increment.json, status.json,
# comments.json, change_history.json (each via _write_json(), which
# briefly creates a same-named ".tmp" sibling during its atomic write
# -- "change_history.json.tmp" = 23 characters is the longest of
# those), and each stored version file under its own "files/"
# subdirectory as "files/v<N>_<YYYY-MM-DD><ext>" (the version NUMBER
# and upload DATE only -- never the original uploaded filename; up to
# a 3-digit version number = 26 characters). 30 leaves a little slack
# beyond the longest of those confirmed cases without needing to be
# exact.
_LONGEST_RELATIVE_FILE_PATH_LENGTH = 30

# Keeps the full worst-case path (base data dir + .../increments/ +
# slug + the longest file inside it, see above) safely under Windows'
# ~259-character failure point confirmed on a real client machine --
# 240 leaves real margin on top of that, since Excel, Explorer, and
# OneDrive are all independently known to have their own long-path
# quirks that can bite before the raw OS limit does. Deliberately NOT
# worked around with the \\?\ long-path prefix: that prefix itself
# isn't reliably honored by Excel/Explorer/OneDrive either, so a
# genuinely SHORTER path is the fix that actually holds up everywhere,
# not just past Windows' own API.
_TARGET_MAX_PATH_LENGTH = 240

# However short _increment_slug_max_length()'s real-budget computation
# comes out, never go below this -- enough room for a real Project #
# (e.g. "S221919-38-00", 13 characters) plus a short, still-recognizable
# fragment of the real name. A base data directory deep enough to push
# the computed budget below this is pathological enough that the
# write will likely still fail -- but it fails with the clear
# "folder path may be too long" message (see ui/workers.py.
# show_storage_error) either way, never silently.
_MIN_INCREMENT_SLUG_LENGTH = 24

# Reserves room for _unique_slug()'s own "-2".."-99" collision suffix,
# so two long names that truncate to the same prefix can't overflow
# the budget _increment_slug_max_length() already computed to be safe.
_COLLISION_SUFFIX_RESERVE = 4

# A FIXED ceiling on top of the dynamic, per-installation budget
# (_increment_slug_max_length_for_prefix()) -- independent of how deep
# any one machine's own base data directory happens to be nested. The
# client's team shares one OneDrive-synced data folder across several
# computers; without this, a brand-new increment's on-disk slug could
# come out a DIFFERENT length (and so different text) purely because
# of which machine happened to create it -- a shallower local
# OneDrive/username path on one computer, a deeper one on another,
# both pointed at the exact same shared folder. Folder names shouldn't
# depend on incidental per-machine path depth. The dynamic budget can
# only ever shorten the result further than this fixed cap, never
# lengthen it past it.
_FIXED_INCREMENT_SLUG_LENGTH_CAP = 64


def _increment_slug_max_length_for_prefix(prefix_length: int) -> int:
    """The actual budget arithmetic behind
    ProjectStore._increment_slug_max_length() -- pulled out as a
    standalone function of a plain int (not a method needing a real,
    OS-native Path) specifically so it can be verified directly
    against a SIMULATED Windows path length (via PureWindowsPath, on
    any host OS, including the macOS/Linux machines this app is
    actually developed and tested on) rather than only indirectly
    through whatever path separator convention the DEV machine's own
    pathlib happens to use.

    Bounded above by _FIXED_INCREMENT_SLUG_LENGTH_CAP (a fixed ceiling,
    independent of any one machine's own path depth -- see that
    constant's own docstring for why) and below by
    _MIN_INCREMENT_SLUG_LENGTH (see that constant's docstring).
    """
    budget = (
        _TARGET_MAX_PATH_LENGTH
        - prefix_length
        - 2  # separator before the slug, and before the file inside it
        - _LONGEST_RELATIVE_FILE_PATH_LENGTH
        - _COLLISION_SUFFIX_RESERVE
    )
    return max(min(budget, _FIXED_INCREMENT_SLUG_LENGTH_CAP), _MIN_INCREMENT_SLUG_LENGTH)


def _slugify_increment(name: str, project_number: Any = None, max_length: int = _MAX_INCREMENT_SLUG_LENGTH) -> str:
    """_slugify(), capped to `max_length` for increment directory names
    specifically -- see _MAX_INCREMENT_SLUG_LENGTH's docstring for why
    that's a per-installation-computed value in practice (passed in by
    create_increment() via _increment_slug_max_length()), not always
    this function's own default.

    A name short enough to fit is slugified exactly as before (so every
    existing, already-short increment's directory name is completely
    unaffected by this change). A name whose slug would exceed the cap
    is truncated, then disambiguated so two genuinely different long
    names can never collide on the same truncated prefix and silently
    share (or overwrite) one folder: the Project # when given (already
    confirmed short -- e.g. "S221919-38-00" -- and unique per real
    project), or else a short content hash of the FULL, untruncated
    name (covers the rare case project_number itself is missing/blank).

    The disambiguator (plus its separator) is carved OUT of max_length,
    not appended on top of it -- max_length is the final result's own
    true hard cap. The previous version of this function truncated the
    name to its cap FIRST and only then appended "-{disambiguator}",
    so the actual result could run up to len(disambiguator) + 1
    characters PAST the stated cap -- confirmed as the exact root cause
    of a real client's Windows path-length failure (a 100-character cap
    plus a 13-character Project # disambiguator produced a 114-
    character slug; see _MAX_INCREMENT_SLUG_LENGTH's own docstring).
    """
    base = _slugify(name)
    if len(base) <= max_length:
        return base
    disambiguator = _slugify(str(project_number)) if project_number else hashlib.sha1(name.encode("utf-8")).hexdigest()[:8]
    available_for_prefix = max_length - len(disambiguator) - 1  # -1: the "-" joining them
    if available_for_prefix < 1:
        # max_length is too small even for the disambiguator alone --
        # only possible with a pathologically deep base data directory
        # (see _MIN_INCREMENT_SLUG_LENGTH, which this should never
        # actually fall below in practice). Falls back to the
        # disambiguator by itself rather than raising -- still unique,
        # even if no longer a recognizable fragment of the real name.
        return disambiguator[:max(max_length, 1)]
    truncated = base[:available_for_prefix].rstrip("-")
    return f"{truncated}-{disambiguator}"


_INCREMENT_NUMBER_RE = re.compile(r"^\s*INC\s+(\d+)", re.IGNORECASE)


def _increment_sort_key(name: str) -> tuple:
    """See list_increments()'s comment -- numbered increments ("INC 2 -
    ...") sort first, in numeric order (so INC 9 sorts before INC 10,
    unlike a plain string sort); anything else sorts after, alphabetically.
    """
    match = _INCREMENT_NUMBER_RE.match(name)
    if match:
        return (0, int(match.group(1)), name)
    return (1, 0, name)


def _unique_slug(base_slug: str, taken: set[str]) -> str:
    if base_slug not in taken:
        return base_slug
    n = 2
    while f"{base_slug}-{n}" in taken:
        n += 1
    return f"{base_slug}-{n}"


def _unique_dir_name(parent: Path, base_name: str) -> str:
    if not (parent / base_name).exists():
        return base_name
    n = 2
    while (parent / f"{base_name}-{n}").exists():
        n += 1
    return f"{base_name}-{n}"


def _write_json(path: Path, data: dict | list) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_name(path.name + ".tmp")
    tmp.write_text(json.dumps(data, indent=2))
    tmp.replace(path)  # atomic on the same filesystem -- no half-written file on crash


def _read_json(path: Path) -> Any:
    return json.loads(path.read_text())


class ProjectStore:
    """File-based storage for projects, increments, file versions, and
    per-increment status. No caching -- every call reads/writes disk
    directly. This is a single-user local desktop tool with small data
    volumes; correctness and simplicity win over speed here.
    """

    def __init__(self, base_dir: str | Path | None = None):
        # base_dir=None (not a frozen default-parameter value) so this
        # re-resolves get_default_data_dir() -- and therefore whatever's
        # CURRENTLY configured -- on every construction, rather than
        # freezing whatever was configured at import time. Matters
        # concretely after Settings > Change Data Location: the app
        # constructs a fresh ProjectStore() to pick up the new location
        # without a restart (see ui/pages/home_page.py).
        #
        # Every test in this repo passes its own explicit temp base_dir,
        # so it never resolves the default and never touches the real
        # configured/fallback location.
        self.base_dir = Path(base_dir) if base_dir is not None else get_default_data_dir()
        self.projects_dir = self.base_dir / "projects"
        self.projects_dir.mkdir(parents=True, exist_ok=True)

    # ------------------------------------------------------------------
    # projects
    # ------------------------------------------------------------------
    def _project_dirs(self) -> list[Path]:
        return [p for p in self.projects_dir.iterdir() if p.is_dir()]

    def _project_json_path(self, slug: str) -> Path:
        return self.projects_dir / slug / "project.json"

    def _project_dir_path(self, slug: str) -> Path:
        return self.projects_dir / slug

    def list_projects(self) -> list[ProjectRecord]:
        records = []
        for project_dir in self._project_dirs():
            json_path = project_dir / "project.json"
            if not json_path.exists():
                continue
            # NOTE: data may still have "home_folder"/"max_increments" keys
            # from a project.json written before those were removed --
            # simply never read here, so old files load exactly like new
            # ones instead of erroring on an unexpected key. home_folder is
            # always computed fresh below, never trusted from disk, so it
            # can never go stale relative to base_dir/slug.
            data = _read_json(json_path)
            records.append(
                ProjectRecord(
                    name=data["name"],
                    home_folder=str(self._project_dir_path(project_dir.name)),
                    slug=project_dir.name,
                )
            )
        return records

    def get_project(self, name: str) -> ProjectRecord | None:
        return next((p for p in self.list_projects() if p.name == name), None)

    def create_project(self, name: str) -> ProjectRecord:
        taken = {p.name for p in self._project_dirs()}
        slug = _unique_slug(_slugify(name), taken)
        _write_json(self._project_json_path(slug), {"name": name})
        return ProjectRecord(name=name, home_folder=str(self._project_dir_path(slug)), slug=slug)

    def update_project(self, name: str, new_name: str) -> ProjectRecord | None:
        """Renaming also renames the project's folder to match, so Home
        Folder (a live path computed from the slug) never shows a stale
        pre-rename name -- see the module docstring's "PROJECT's slug is
        the one deliberate exception" note. Same _unique_slug() collision
        strategy create_project() uses, with the project's OWN current
        slug excluded from the "taken" set (so a rename that still
        slugifies the same is a no-op, not a spurious self-collision).
        Path.rename() moves the whole directory tree in one atomic,
        same-filesystem operation -- every increment subfolder, version
        file, and status.json underneath moves with it automatically.
        """
        project = self.get_project(name)
        if project is None:
            return None

        taken = {p.name for p in self._project_dirs() if p.name != project.slug}
        new_slug = _unique_slug(_slugify(new_name), taken)

        if new_slug != project.slug:
            self._project_dir_path(project.slug).rename(self._project_dir_path(new_slug))

        _write_json(self._project_json_path(new_slug), {"name": new_name})
        return ProjectRecord(name=new_name, home_folder=str(self._project_dir_path(new_slug)), slug=new_slug)

    def delete_project(self, name: str) -> None:
        """Soft-delete: moves the project's whole folder -- every
        increment, version file, and status.json, untouched -- out of
        projects/ into <base_dir>/_deleted/, timestamp-prefixed so
        repeated deletions of similarly-named projects never collide.
        Nothing is permanently destroyed; recovery is a manual filesystem
        move back into projects/, not a UI feature (this app stores
        compliance/inspection records, so accidental data loss here is
        the failure mode to avoid).
        """
        project = self.get_project(name)
        if project is None:
            return
        deleted_dir = self.base_dir / DELETED_DIRNAME
        deleted_dir.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S%f")
        dest_name = _unique_dir_name(deleted_dir, f"{timestamp}_{project.slug}")
        shutil.move(str(self.projects_dir / project.slug), str(deleted_dir / dest_name))

    # ------------------------------------------------------------------
    # increments
    # ------------------------------------------------------------------
    def _increments_dir(self, project_slug: str) -> Path:
        return self.projects_dir / project_slug / "increments"

    def _increment_dirs(self, project_slug: str) -> list[Path]:
        increments_dir = self._increments_dir(project_slug)
        if not increments_dir.exists():
            return []
        return [d for d in increments_dir.iterdir() if d.is_dir()]

    def _increment_json_path(self, project_slug: str, increment_slug: str) -> Path:
        return self._increments_dir(project_slug) / increment_slug / "increment.json"

    def _increment_slug_max_length(self, project_slug: str) -> int:
        """How long a NEW increment's slug may be, computed from the
        ACTUAL base data directory's real length at runtime -- see
        _increment_slug_max_length_for_prefix() for the actual budget
        arithmetic (a standalone, host-OS-independent function, so it
        can be verified directly against a SIMULATED Windows path via
        PureWindowsPath regardless of which OS the test runs on --
        see tests/test_increment_path_length.py). This method's only
        job is computing the one real, installation-specific number
        that function needs: the length of THIS project's actual
        ".../increments" prefix, in whatever form pathlib resolves it
        to on the OS actually running.
        """
        prefix_length = len(str(self._increments_dir(project_slug)))
        return _increment_slug_max_length_for_prefix(prefix_length)

    def list_increments(self, project_name: str) -> list[IncrementRecord]:
        project = self.get_project(project_name)
        if project is None:
            return []
        records = []
        for increment_dir in self._increment_dirs(project.slug):
            json_path = increment_dir / "increment.json"
            if not json_path.exists():
                logger.warning("Skipping increment directory with no increment.json: %s", increment_dir)
                continue
            try:
                data = _read_json(json_path)
                records.append(
                    IncrementRecord(
                        name=data["name"],
                        version=data["version"],
                        last_updated=data["last_updated"],
                        slug=increment_dir.name,
                        display_name=data.get("display_name"),
                    )
                )
            except (json.JSONDecodeError, KeyError, OSError) as exc:
                logger.warning("Skipping increment directory with invalid increment.json: %s (%s)", increment_dir, exc)
                continue
        # _increment_dirs() reflects raw filesystem iterdir() order, which
        # is arbitrary (not upload order, not alphabetical) and can vary
        # by OS/filesystem -- so without this, the increment list's order
        # isn't just "not numeric", it's not even stable. Every increment
        # name in practice starts with "INC <number>" (the record name
        # convention read by core.excel_reader.get_record_name()), so sort
        # on that leading number; anything that doesn't match the pattern
        # sorts after all numbered increments, alphabetically, rather than
        # raising -- this is a display-order convenience, not the
        # identity-matching logic in core.increment_matcher, so it should
        # degrade gracefully instead of erroring on an unexpected name.
        return sorted(records, key=lambda r: _increment_sort_key(r.name))

    def get_increment(self, project_name: str, increment_name: str) -> IncrementRecord | None:
        return next((i for i in self.list_increments(project_name) if i.name == increment_name), None)

    def create_increment(
        self,
        project_name: str,
        increment_name: str,
        source_file_path: str,
        project_number: Any = None,
        display_name: str | None = None,
    ) -> IncrementRecord:
        """Creates a brand-new increment and saves source_file_path as its
        version 1, with an empty status.json (nothing has ever been marked
        for any index in a file that didn't exist in this app a moment
        ago). Raises ValueError if the project doesn't exist or an
        increment with this name already exists in it -- callers are
        expected to route that to "Upload New Version" instead.

        project_number (core.excel_reader.parse_project_info()'s
        "project_number") is optional and ONLY used to disambiguate the
        on-disk slug if increment_name's own slug is too long to use as-is
        -- see _slugify_increment(). Passing None here never changes
        behavior for any increment_name short enough not to need it.

        display_name is a separate, purely cosmetic label (see
        _sanitize_display_name() and IncrementRecord.display_name's own
        docstring) -- sanitized here, the one place every caller's
        display_name ends up written, so there's a single source of
        truth for what's actually allowed in it. NEVER used to compute
        `slug` above -- increment_name/project_number (the real,
        immutable identity) already fully determine it before this
        parameter is even looked at.
        """
        project = self.get_project(project_name)
        if project is None:
            raise ValueError(f"No project named {project_name!r}")
        if self.get_increment(project_name, increment_name) is not None:
            raise ValueError(f"Increment {increment_name!r} already exists in project {project_name!r}")

        taken = {d.name for d in self._increment_dirs(project.slug)}
        max_slug_length = self._increment_slug_max_length(project.slug)
        slug = _unique_slug(_slugify_increment(increment_name, project_number, max_slug_length), taken)
        display_name = _sanitize_display_name(display_name)
        today = date.today().isoformat()

        # If any of these three writes fails partway through (most
        # realistically a Windows path-length OSError, but any OSError
        # is handled the same way), the half-created folder is removed
        # rather than left behind as an orphan with no increment.json
        # for list_increments() to find -- see that function's own
        # skip-and-log handling for the complementary side of this (an
        # orphan that somehow still exists on disk, from before this
        # cleanup existed, must never break the project's increment
        # list). The exception itself is re-raised unchanged, for the
        # UI layer to show as a clear "could not save" message instead
        # of the misleading "couldn't be parsed" one a lower-level
        # OSError used to surface as (see ui/workers.py.show_storage_error).
        increment_dir = self._increments_dir(project.slug) / slug
        try:
            _write_json(
                self._increment_json_path(project.slug, slug),
                {"name": increment_name, "version": 1, "last_updated": today, "display_name": display_name},
            )
            _write_json(self._status_json_path(project.slug, slug), {})
            self._copy_version_file(project.slug, slug, source_file_path, version=1, as_of=today)
        except OSError:
            shutil.rmtree(increment_dir, ignore_errors=True)
            raise

        return IncrementRecord(name=increment_name, version=1, last_updated=today, slug=slug, display_name=display_name)

    def save_new_version(self, project_name: str, increment_name: str, source_file_path: str) -> IncrementRecord:
        """Copies source_file_path in as the next version and bumps the
        increment's version number / last_updated date. Never overwrites
        an existing version file -- every upload is kept.
        """
        project = self.get_project(project_name)
        if project is None:
            raise ValueError(f"No project named {project_name!r}")
        increment = self.get_increment(project_name, increment_name)
        if increment is None:
            raise ValueError(f"No increment named {increment_name!r} in project {project_name!r}")

        new_version = increment.version + 1
        today = date.today().isoformat()
        _write_json(
            self._increment_json_path(project.slug, increment.slug),
            # display_name is carried forward from the CURRENT record,
            # not re-derived or dropped -- a user-set label must survive
            # every future version bump, the same way status.json's
            # Done/Open marks already do (see
            # core.status_tracker.carry_forward_status).
            {"name": increment_name, "version": new_version, "last_updated": today, "display_name": increment.display_name},
        )
        self._copy_version_file(project.slug, increment.slug, source_file_path, version=new_version, as_of=today)
        return IncrementRecord(
            name=increment_name, version=new_version, last_updated=today, slug=increment.slug,
            display_name=increment.display_name,
        )

    def delete_increment(self, project_name: str, increment_name: str) -> None:
        """Soft-delete: moves the increment's whole folder -- every
        version file and status.json, untouched -- out of
        projects/{project}/increments/ into <base_dir>/_deleted/,
        exactly the same pattern as delete_project() (timestamp-prefixed
        so repeated deletions of similarly-named increments never
        collide; nothing merged, nothing overwritten). The project
        itself and its other increments are untouched -- only this one
        increment's folder moves.

        Prefixed with the project's slug as well as the increment's
        (not just the increment's) since increment slugs are only
        unique within a project -- two different projects could each
        have an "inc-1", and _deleted/ is a single flat folder shared
        across every project, not one per project.
        """
        project = self.get_project(project_name)
        if project is None:
            return
        increment = self.get_increment(project_name, increment_name)
        if increment is None:
            return
        deleted_dir = self.base_dir / DELETED_DIRNAME
        deleted_dir.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S%f")
        dest_name = _unique_dir_name(deleted_dir, f"{timestamp}_{project.slug}_{increment.slug}")
        shutil.move(str(self._increments_dir(project.slug) / increment.slug), str(deleted_dir / dest_name))

    def _files_dir(self, project_slug: str, increment_slug: str) -> Path:
        return self._increments_dir(project_slug) / increment_slug / "files"

    def _copy_version_file(
        self, project_slug: str, increment_slug: str, source_file_path: str, version: int, as_of: str
    ) -> Path:
        source = Path(source_file_path)
        files_dir = self._files_dir(project_slug, increment_slug)
        files_dir.mkdir(parents=True, exist_ok=True)
        dest = files_dir / f"v{version}_{as_of}{source.suffix.lower()}"
        shutil.copy2(source, dest)  # copy, not move -- the user's original upload is theirs to keep
        return dest

    def list_version_files(self, project_name: str, increment_name: str) -> list[Path]:
        """Every stored version file for this increment, oldest (v1) first."""
        project = self.get_project(project_name)
        increment = self.get_increment(project_name, increment_name) if project else None
        if project is None or increment is None:
            return []
        files_dir = self._files_dir(project.slug, increment.slug)
        if not files_dir.exists():
            return []

        def version_of(path: Path) -> int:
            match = re.match(r"v(\d+)_", path.name)
            return int(match.group(1)) if match else 0

        return sorted((f for f in files_dir.iterdir() if f.is_file()), key=version_of)

    def current_version_path(self, project_name: str, increment_name: str) -> Path | None:
        files = self.list_version_files(project_name, increment_name)
        return files[-1] if files else None

    def version_path(self, project_name: str, increment_name: str, version: int) -> Path | None:
        for f in self.list_version_files(project_name, increment_name):
            if f.name.startswith(f"v{version}_"):
                return f
        return None

    def list_versions(self, project_name: str, increment_name: str) -> list[VersionRecord]:
        """Every stored version of this increment, NEWEST first (the
        reverse of list_version_files(), which is oldest-first for
        version-history bookkeeping) -- for a UI version-selector
        dropdown, which should default to the top entry. Parsed
        straight from each stored file's "vN_YYYY-MM-DD.ext" name (see
        _copy_version_file), not from increment.json, since that only
        ever holds the CURRENT version's date.
        """
        records = []
        for f in self.list_version_files(project_name, increment_name):
            match = re.match(r"v(\d+)_(\d{4}-\d{2}-\d{2})", f.name)
            if match:
                records.append(VersionRecord(version=int(match.group(1)), uploaded_date=match.group(2)))
        return list(reversed(records))

    # ------------------------------------------------------------------
    # status
    # ------------------------------------------------------------------
    def _status_json_path(self, project_slug: str, increment_slug: str) -> Path:
        return self._increments_dir(project_slug) / increment_slug / "status.json"

    def load_status(self, project_name: str, increment_name: str) -> dict[str, dict[int, Any]]:
        """Returns {index: {stage_number: status}} with stage_number as
        int -- see the on-disk format note in the module docstring for why
        the JSON file itself uses string keys.
        """
        project = self.get_project(project_name)
        increment = self.get_increment(project_name, increment_name) if project else None
        if project is None or increment is None:
            return {}
        path = self._status_json_path(project.slug, increment.slug)
        if not path.exists():
            return {}
        raw = _read_json(path)
        return {index: {int(stage): value for stage, value in stages.items()} for index, stages in raw.items()}

    def save_status(self, project_name: str, increment_name: str, status: dict[str, dict[int, Any]]) -> None:
        """Takes {index: {stage_number: status}} with stage_number as int
        -- converted to string keys before writing, since JSON object keys
        must be strings.
        """
        project = self.get_project(project_name)
        increment = self.get_increment(project_name, increment_name) if project else None
        if project is None or increment is None:
            raise ValueError(f"No increment named {increment_name!r} in project {project_name!r}")
        raw = {index: {str(stage): value for stage, value in stages.items()} for index, stages in status.items()}
        _write_json(self._status_json_path(project.slug, increment.slug), raw)

    # ------------------------------------------------------------------
    # change history (app-detected diffs between versions, persisted at
    # confirm time -- see the module docstring's change_history.json entry)
    # ------------------------------------------------------------------
    def _change_history_json_path(self, project_slug: str, increment_slug: str) -> Path:
        return self._increments_dir(project_slug) / increment_slug / "change_history.json"

    def append_change_history_entry(self, project_name: str, increment_name: str, entry: dict[str, Any]) -> None:
        """Appends one entry to change_history.json -- read, append,
        write-back-whole-file, same as every other JSON file in this
        store; never overwrites or drops prior entries. `entry` is
        caller-built (see ui.mock_data.MockDataStore.confirm_update) from
        whatever the review screen already computed for this specific
        update, so this method's only job is persisting it, not deriving
        it.
        """
        project = self.get_project(project_name)
        increment = self.get_increment(project_name, increment_name) if project else None
        if project is None or increment is None:
            raise ValueError(f"No increment named {increment_name!r} in project {project_name!r}")
        path = self._change_history_json_path(project.slug, increment.slug)
        history = _read_json(path) if path.exists() else []
        history.append(entry)
        _write_json(path, history)

    def load_change_history(self, project_name: str, increment_name: str) -> list[dict[str, Any]]:
        """All persisted change-history entries for this increment,
        OLDEST first (the order they were appended, same as
        list_version_files()) -- callers wanting newest-first (e.g. a
        Changes-tab "Update History" section) reverse this themselves.
        """
        project = self.get_project(project_name)
        increment = self.get_increment(project_name, increment_name) if project else None
        if project is None or increment is None:
            return []
        path = self._change_history_json_path(project.slug, increment.slug)
        if not path.exists():
            return []
        return _read_json(path)

    # ------------------------------------------------------------------
    # comments -- free-text notes a user of THIS app types directly onto
    # an increment's Changes tab; distinct from both status.json (a
    # Done/Open mark) and change_history.json (app-DETECTED diffs) -- a
    # human's own words, not derived from the file at all. Unlike
    # change_history.json's append-only convention, comments support
    # full CRUD: update_comment() edits a comment's text in place
    # (preserving its original "id" and creation "timestamp" -- an edit
    # is not delete+recreate, which would lose when the comment was
    # first written -- and sets "edited_timestamp" so the fact and time
    # of the edit is visible, not silently hidden). delete_comment()
    # removes an entry outright, an explicit, visible action.
    # ------------------------------------------------------------------
    def _comments_json_path(self, project_slug: str, increment_slug: str) -> Path:
        return self._increments_dir(project_slug) / increment_slug / "comments.json"

    def add_comment(self, project_name: str, increment_name: str, text: str) -> dict[str, Any]:
        """Appends one {id, timestamp, text, edited_timestamp} entry to
        comments.json and returns it -- `id` (a uuid4 hex string) is
        what update_comment()/delete_comment() target, since comment
        text itself isn't a safe identity key (two comments can
        legitimately have identical text). `edited_timestamp` starts
        None -- see update_comment().
        """
        project = self.get_project(project_name)
        increment = self.get_increment(project_name, increment_name) if project else None
        if project is None or increment is None:
            raise ValueError(f"No increment named {increment_name!r} in project {project_name!r}")
        entry = {
            "id": uuid.uuid4().hex,
            "timestamp": datetime.now().isoformat(),
            "text": text,
            "edited_timestamp": None,
        }
        path = self._comments_json_path(project.slug, increment.slug)
        comments = _read_json(path) if path.exists() else []
        comments.append(entry)
        _write_json(path, comments)
        return entry

    def load_comments(self, project_name: str, increment_name: str) -> list[dict[str, Any]]:
        """All persisted comments for this increment, OLDEST first
        (append order, same convention as load_change_history) --
        callers wanting newest-first reverse this themselves.
        """
        project = self.get_project(project_name)
        increment = self.get_increment(project_name, increment_name) if project else None
        if project is None or increment is None:
            return []
        path = self._comments_json_path(project.slug, increment.slug)
        if not path.exists():
            return []
        return _read_json(path)

    def update_comment(
        self, project_name: str, increment_name: str, comment_id: str, new_text: str
    ) -> dict[str, Any] | None:
        """Edits exactly the comment with this id IN PLACE: only "text"
        and "edited_timestamp" change -- "id" and the original creation
        "timestamp" are untouched, so an edit can never be mistaken for
        a newly-written comment or lose when it was actually first
        posted. Returns the updated entry, or None if no comment with
        this id exists (a no-op, not an error -- same
        stale-list-tolerant convention as delete_comment()).
        """
        project = self.get_project(project_name)
        increment = self.get_increment(project_name, increment_name) if project else None
        if project is None or increment is None:
            return None
        path = self._comments_json_path(project.slug, increment.slug)
        if not path.exists():
            return None
        comments = _read_json(path)
        updated = None
        for comment in comments:
            if comment.get("id") == comment_id:
                comment["text"] = new_text
                comment["edited_timestamp"] = datetime.now().isoformat()
                updated = comment
                break
        if updated is not None:
            _write_json(path, comments)
        return updated

    def delete_comment(self, project_name: str, increment_name: str, comment_id: str) -> None:
        """Removes exactly the comment with this id, if present -- a
        no-op (not an error) if it's already gone, so a UI double-click
        or a stale in-memory list can never raise.
        """
        project = self.get_project(project_name)
        increment = self.get_increment(project_name, increment_name) if project else None
        if project is None or increment is None:
            return
        path = self._comments_json_path(project.slug, increment.slug)
        if not path.exists():
            return
        comments = [c for c in _read_json(path) if c.get("id") != comment_id]
        _write_json(path, comments)
