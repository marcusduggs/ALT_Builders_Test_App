"""
Reads/writes the tiny per-user config file that tells the app WHERE its
data folder (AltamiranoBuildersAppData, by default) actually lives, and
relocates that folder when the user changes their mind (the first-launch
picker, or Settings > Change Data Location).

This config has to live somewhere separate from the data folder itself --
the app needs to know where to look BEFORE it can look inside it -- so it
goes in the OS-standard per-user app config location, not alongside (or
inside) the data folder it points at:

    Windows: %APPDATA%\\AltamiranoBuildersApp\\config.json
    macOS:   ~/Library/Application Support/AltamiranoBuildersApp/config.json
    other:   ~/.config/AltamiranoBuildersApp/config.json (XDG convention --
             this app has no documented Linux target, but should degrade
             sensibly rather than crash if it's ever run there)

The file just holds {"data_dir": "<absolute path>"}. Its absence is the
exact first-launch signal ui/app.py's startup picker checks for -- see
read_configured_data_dir().

Every function here takes an optional explicit path override
(config_path, or old_dir/new_dir for relocate_data_dir) specifically so
tests can exercise this logic without ever touching the real, OS-specific
config file or the real data directory on a developer's machine.
"""

from __future__ import annotations

import json
import os
import shutil
import sys
from pathlib import Path

APP_CONFIG_DIRNAME = "AltamiranoBuildersApp"
CONFIG_FILENAME = "config.json"

LEGACY_DATA_DIR_NAME = "SubmissionAppData"  # this project's old working name
DATA_DIR_NAME = "AltamiranoBuildersAppData"


def get_config_path() -> Path:
    """The real, OS-specific config.json location. Production code calls
    this with no arguments; tests should instead pass an explicit
    config_path to read_configured_data_dir()/write_configured_data_dir()
    rather than calling this at all.
    """
    if sys.platform == "win32":
        appdata = os.environ.get("APPDATA")
        base = Path(appdata) if appdata else Path.home() / "AppData" / "Roaming"
    elif sys.platform == "darwin":
        base = Path.home() / "Library" / "Application Support"
    else:
        xdg = os.environ.get("XDG_CONFIG_HOME")
        base = Path(xdg) if xdg else Path.home() / ".config"
    return base / APP_CONFIG_DIRNAME / CONFIG_FILENAME


def default_data_dir_suggestion(home_dir: Path | None = None) -> Path:
    """The historical hardcoded location -- pre-filled as the suggested
    default in the first-launch picker, and the fallback
    core.project_store.get_default_data_dir() uses if somehow no config
    exists AND no picker has run yet (should only happen for a caller
    that constructs a ProjectStore without going through ui/app.py's
    first-launch check at all -- e.g. a script).

    home_dir defaults to the real Path.home() in production; tests pass
    an explicit temp dir instead so this never resolves to (let alone
    touches) a real user's actual home directory.
    """
    return (home_dir or Path.home()) / DATA_DIR_NAME


def read_configured_data_dir(config_path: Path | None = None) -> Path | None:
    """None means no config file yet -- the first-launch condition."""
    path = config_path or get_config_path()
    if not path.exists():
        return None
    try:
        data = json.loads(path.read_text())
    except (json.JSONDecodeError, OSError):
        return None
    data_dir = data.get("data_dir")
    return Path(data_dir) if data_dir else None


def write_configured_data_dir(data_dir: Path, config_path: Path | None = None) -> None:
    path = config_path or get_config_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_name(path.name + ".tmp")
    tmp.write_text(json.dumps({"data_dir": str(data_dir)}, indent=2))
    tmp.replace(path)  # atomic on the same filesystem -- no half-written file on crash


def relocate_data_dir(old_dir: Path, new_dir: Path) -> None:
    """Moves the ENTIRE data folder from old_dir to new_dir.

    A plain Path.rename() when possible -- instant, same-filesystem, same
    approach already used for the SubmissionAppData->AltamiranoBuildersAppData
    migration and the project/increment soft-deletes. Falls back to a real
    copy-then-delete when the two locations are on different drives/
    volumes: rename() raises OSError there (errno EXDEV) since a rename
    can't cross filesystems -- this is an EXPECTED, handled case, not a
    bug path, and is exactly how a user picking a different drive in the
    Settings dialog gets handled correctly.

    No-ops if old_dir doesn't exist (nothing to move -- e.g. a genuinely
    fresh install with no prior data anywhere) or if old_dir == new_dir
    (nothing actually changed). Refuses to proceed if new_dir already
    exists and is non-empty -- never silently merges into or overwrites
    whatever's already there, same "never merge, never overwrite"
    philosophy as this app's soft-delete features.
    """
    if old_dir == new_dir or not old_dir.exists():
        return
    if new_dir.exists() and any(new_dir.iterdir()):
        raise FileExistsError(
            f"{new_dir} already exists and is not empty -- refusing to move data there "
            "to avoid merging with or overwriting whatever's already there"
        )

    new_dir.parent.mkdir(parents=True, exist_ok=True)
    try:
        old_dir.rename(new_dir)
    except OSError:
        # Cross-device (different drive/volume): copy everything over
        # first, then remove the original only once the copy has fully
        # succeeded -- never delete the source before the copy is done.
        if new_dir.exists():
            shutil.rmtree(new_dir)
        shutil.copytree(old_dir, new_dir)
        shutil.rmtree(old_dir)
