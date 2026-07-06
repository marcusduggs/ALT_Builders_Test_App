"""
Minimal local, file-based storage for projects/increments/status/file
versions. No database -- matches the app's goal of being a fully offline,
zero-setup local tool. Every uploaded file is kept forever (versioned,
never overwritten), so there's always an audit trail of exactly what the
state sent and when.

Default data directory: ~/SubmissionAppData

Layout:

    <base_dir>/
      projects/
        <project-slug>/
          project.json                  {name, home_folder, max_increments}
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

DEFAULT_DATA_DIR = Path.home() / "SubmissionAppData"
DELETED_DIRNAME = "_deleted"


@dataclass
class ProjectRecord:
    name: str
    home_folder: str
    max_increments: int
    slug: str


@dataclass
class IncrementRecord:
    name: str
    version: int
    last_updated: str
    slug: str


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
        self.projects_dir = self.base_dir / "projects"
        self.projects_dir.mkdir(parents=True, exist_ok=True)

    # ------------------------------------------------------------------
    # projects
    # ------------------------------------------------------------------
    def _project_dirs(self) -> list[Path]:
        return [p for p in self.projects_dir.iterdir() if p.is_dir()]

    def _project_json_path(self, slug: str) -> Path:
        return self.projects_dir / slug / "project.json"

    def list_projects(self) -> list[ProjectRecord]:
        records = []
        for project_dir in self._project_dirs():
            json_path = project_dir / "project.json"
            if not json_path.exists():
                continue
            data = _read_json(json_path)
            records.append(
                ProjectRecord(
                    name=data["name"],
                    home_folder=data.get("home_folder", ""),
                    max_increments=data.get("max_increments", 0),
                    slug=project_dir.name,
                )
            )
        return records

    def get_project(self, name: str) -> ProjectRecord | None:
        return next((p for p in self.list_projects() if p.name == name), None)

    def create_project(self, name: str, home_folder: str, max_increments: int) -> ProjectRecord:
        taken = {p.name for p in self._project_dirs()}
        slug = _unique_slug(_slugify(name), taken)
        _write_json(
            self._project_json_path(slug),
            {"name": name, "home_folder": home_folder, "max_increments": max_increments},
        )
        return ProjectRecord(name=name, home_folder=home_folder, max_increments=max_increments, slug=slug)

    def update_project(self, name: str, new_name: str, home_folder: str, max_increments: int) -> ProjectRecord | None:
        project = self.get_project(name)
        if project is None:
            return None
        _write_json(
            self._project_json_path(project.slug),
            {"name": new_name, "home_folder": home_folder, "max_increments": max_increments},
        )
        return ProjectRecord(name=new_name, home_folder=home_folder, max_increments=max_increments, slug=project.slug)

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
