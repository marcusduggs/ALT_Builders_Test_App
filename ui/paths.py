"""Shared filesystem-path helpers for the UI package."""

from __future__ import annotations

import sys
from pathlib import Path


def get_app_dir() -> Path:
    """Directory the app should resolve bundled files (assets, etc.) from,
    whether running as a script (`python -m ui.app`) or as a frozen
    PyInstaller build. Mirrors the get_app_dir() pattern in the legacy
    app.py so both entry points resolve relative paths the same way,
    instead of relying on the current working directory.
    """
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


def get_bundled_path(*parts: str) -> Path:
    """Resolves a bundled, read-only resource -- assets/logo.png,
    ui/style.qss, anything shipped alongside the code rather than
    written at runtime.

    Deliberately separate from get_app_dir(): a --onefile PyInstaller
    build extracts data added via --add-data into a temp directory at
    sys._MEIPASS, NOT next to the built exe, so bundled resources must
    resolve from there when frozen. Plain Python modules' own __file__
    (e.g. ui/app.py's) is NOT reliable for this in a frozen build either
    -- PyInstaller normally packs .py/.pyc sources into a zipped archive
    inside the exe rather than extracting them as loose files, even under
    _MEIPASS, so `Path(__file__).parent` can't be trusted there the way
    it can when running as a script. get_app_dir()'s dirname(sys.executable)
    is still correct for user data (~/SubmissionAppData is home-based
    anyway, so it's untouched by this) and would be correct for a
    --onedir build's resources too, but --onefile is what this project
    builds, so sys._MEIPASS must be checked first.
    """
    if hasattr(sys, "_MEIPASS"):
        base_path = Path(sys._MEIPASS)
    else:
        base_path = get_app_dir()
    return base_path.joinpath(*parts)


def get_asset_path(*parts: str) -> Path:
    return get_bundled_path("assets", *parts)
