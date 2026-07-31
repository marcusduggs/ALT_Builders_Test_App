"""
Sets up file-based logging for the whole app -- attached to the ROOT
logger, so every module's existing `logging.getLogger(__name__)` calls
(and any new ones) are captured here with no per-module changes needed.

Before this module existed, the app had no logging.basicConfig()/
handler configured anywhere. That's invisible, not just unhelpful, once
packaged: the release build runs PyInstaller with --windowed (see
.github/workflows/release.yml), so there is no console for the
default last-resort stderr handler to write to -- a warning/error
logged before this module ran effectively went nowhere. configure_logging()
must be called once, early, at app startup (see ui/app.py's main()) so
that even a failure during first-launch setup gets captured.

The log file lives next to config.json (core.app_config.get_config_path()'s
directory) -- the OS-standard per-user app config location, not the
data directory: logging has to work even if the data directory is
unreadable/misconfigured, which is exactly the kind of thing this log
might need to explain.
"""

from __future__ import annotations

import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path

from core import app_config

LOG_FILENAME = "app.log"
MAX_LOG_BYTES = 2 * 1024 * 1024  # 2 MB
BACKUP_COUNT = 3  # app.log, app.log.1, app.log.2, app.log.3 -- ~8MB ceiling total


def get_log_path() -> Path:
    """The real, OS-specific log file location -- alongside config.json,
    reusing that same platform-detection logic rather than duplicating
    it (core.app_config.get_config_path()'s own directory).
    """
    return app_config.get_config_path().parent / LOG_FILENAME


def configure_logging(log_path: Path | None = None) -> Path:
    """Attaches a rotating file handler to the root logger -- idempotent
    (safe to call more than once; only adds the handler the first time,
    identified by name, so tests or a re-entrant call never end up with
    duplicate handlers double-writing every line). Returns the resolved
    log path.

    log_path defaults to get_log_path() in production; tests should
    pass an explicit path instead, same convention as core.app_config's
    functions, so this never touches a real user's actual config
    directory.
    """
    path = log_path or get_log_path()
    path.parent.mkdir(parents=True, exist_ok=True)

    root_logger = logging.getLogger()
    handler_name = "app_file_handler"
    if any(h.name == handler_name for h in root_logger.handlers):
        return path

    handler = RotatingFileHandler(path, maxBytes=MAX_LOG_BYTES, backupCount=BACKUP_COUNT, encoding="utf-8")
    handler.name = handler_name
    handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(name)s: %(message)s"))
    root_logger.addHandler(handler)
    # INFO, not WARNING -- a silent-hang bug report is exactly the kind
    # of thing where "nothing looked wrong at the time" is the problem;
    # capturing normal operation context (which file, which increment)
    # is what makes a later WARNING/ERROR line actually diagnosable
    # instead of a traceback with no surrounding story.
    root_logger.setLevel(logging.INFO)
    return path
