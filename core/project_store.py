"""
Minimal local, file-based storage for projects/increments/status/file
versions. No database -- matches the app's goal of being a fully offline,
zero-setup local tool. Every uploaded file is kept forever (versioned,
never overwritten), so there's always an audit trail of exactly what the
state sent and when.

Default data directory: ~/AltamiranoBuildersAppData (auto-migrated in
place from ~/SubmissionAppData -- this project's old working name -- if
found; see _migrate_legacy_data_dir())

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
              increment.json            {name, version, last_updated}
              status.json               {index: {stage_number: "Done"|"Open"}}
                                         e.g. {"B-C1": {"5": "Done", "12": "Open"}}
                                         -- JSON object keys are always strings, so
                                         stage numbers are stored as string digits;
                                         load_status()/save_status() convert to/from
                                         int transparently, so every OTHER module in
                                         this codebase deals in int stage numbers.
              files/
                v1_2026-05-10.xlsm
                v2_2026-06-26.xlsm       <- current version = highest v#
      _deleted/
        <timestamp>_<project-slug>/     <- delete_project() moves the whole
                                            project folder here, untouched,
                                            instead of removing it -- manual
                                            recovery only, no restore UI

A project's or increment's "slug" (its folder name) is generated once, from
its name, at creation time, and never changes afterward -- update_project()
only rewrites project.json's "name" field, not the folder. Coupling a
stable-forever filesystem path to a display name a user can freely edit is
exactly the identity-vs-position confusion this codebase has deliberately
avoided elsewhere (see core/excel_reader.py's "PARSING STRATEGY" section).
Every list/get/update call matches on the display name stored *inside*
project.json / increment.json, never on the slug -- callers never see or
need to know slugs exist.
"""

from __future__ import annotations

import json
import re
import shutil
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any

LEGACY_DATA_DIR_NAME = "SubmissionAppData"  # this project's old working name
DATA_DIR_NAME = "AltamiranoBuildersAppData"
DEFAULT_DATA_DIR = Path.home() / DATA_DIR_NAME
DELETED_DIRNAME = "_deleted"


def _migrate_legacy_data_dir() -> None:
    """One-time migration: if ~/SubmissionAppData (the old name) still
    exists and ~/AltamiranoBuildersAppData (the new one) doesn't yet,
    rename the folder in place -- same parent directory, so this is an
    instant filesystem rename, not a copy, and every project/increment/
    status file underneath it moves untouched along with it.

    If the new folder already exists, this does nothing at all: never
    merges, never deletes -- there's no path by which existing data
    could be lost to an edge case here. Safe to call on every startup:
    idempotent, since the second call onward always finds the new
    folder already in place and no-ops.
    """
    legacy_dir = Path.home() / LEGACY_DATA_DIR_NAME
    if legacy_dir.exists() and not DEFAULT_DATA_DIR.exists():
        legacy_dir.rename(DEFAULT_DATA_DIR)
        print(f"Migrated existing data from {LEGACY_DATA_DIR_NAME} to {DATA_DIR_NAME}")


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


@dataclass
class VersionRecord:
    version: int
    uploaded_date: str


def _slugify(name: str) -> str:
    slug = re.sub(r"[^a-z0-9]+", "-", name.strip().lower()).strip("-")
    return slug or "untitled"


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


def _write_json(path: Path, data: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_name(path.name + ".tmp")
    tmp.write_text(json.dumps(data, indent=2))
    tmp.replace(path)  # atomic on the same filesystem -- no half-written file on crash


def _read_json(path: Path) -> dict:
    return json.loads(path.read_text())


class ProjectStore:
    """File-based storage for projects, increments, file versions, and
    per-increment status. No caching -- every call reads/writes disk
    directly. This is a single-user local desktop tool with small data
    volumes; correctness and simplicity win over speed here.
    """

    def __init__(self, base_dir: str | Path = DEFAULT_DATA_DIR):
        self.base_dir = Path(base_dir)
        # Only when using the real default location -- never for a
        # caller-supplied base_dir (every test in this repo passes its
        # own temp dir specifically so it never touches the real
        # ~/AltamiranoBuildersAppData, and migrating some unrelated temp
        # dir would make no sense anyway).
        if self.base_dir == DEFAULT_DATA_DIR:
            _migrate_legacy_data_dir()
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
        project = self.get_project(name)
        if project is None:
            return None
        _write_json(self._project_json_path(project.slug), {"name": new_name})
        return ProjectRecord(name=new_name, home_folder=str(self._project_dir_path(project.slug)), slug=project.slug)

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

    def list_increments(self, project_name: str) -> list[IncrementRecord]:
        project = self.get_project(project_name)
        if project is None:
            return []
        records = []
        for increment_dir in self._increment_dirs(project.slug):
            json_path = increment_dir / "increment.json"
            if not json_path.exists():
                continue
            data = _read_json(json_path)
            records.append(
                IncrementRecord(
                    name=data["name"],
                    version=data["version"],
                    last_updated=data["last_updated"],
                    slug=increment_dir.name,
                )
            )
        return records

    def get_increment(self, project_name: str, increment_name: str) -> IncrementRecord | None:
        return next((i for i in self.list_increments(project_name) if i.name == increment_name), None)

    def create_increment(self, project_name: str, increment_name: str, source_file_path: str) -> IncrementRecord:
        """Creates a brand-new increment and saves source_file_path as its
        version 1, with an empty status.json (nothing has ever been marked
        for any index in a file that didn't exist in this app a moment
        ago). Raises ValueError if the project doesn't exist or an
        increment with this name already exists in it -- callers are
        expected to route that to "Upload New Version" instead.
        """
        project = self.get_project(project_name)
        if project is None:
            raise ValueError(f"No project named {project_name!r}")
        if self.get_increment(project_name, increment_name) is not None:
            raise ValueError(f"Increment {increment_name!r} already exists in project {project_name!r}")

        taken = {d.name for d in self._increment_dirs(project.slug)}
        slug = _unique_slug(_slugify(increment_name), taken)
        today = date.today().isoformat()

        _write_json(
            self._increment_json_path(project.slug, slug),
            {"name": increment_name, "version": 1, "last_updated": today},
        )
        _write_json(self._status_json_path(project.slug, slug), {})
        self._copy_version_file(project.slug, slug, source_file_path, version=1, as_of=today)

        return IncrementRecord(name=increment_name, version=1, last_updated=today, slug=slug)

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
            {"name": increment_name, "version": new_version, "last_updated": today},
        )
        self._copy_version_file(project.slug, increment.slug, source_file_path, version=new_version, as_of=today)
        return IncrementRecord(name=increment_name, version=new_version, last_updated=today, slug=increment.slug)

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
