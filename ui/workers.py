"""
Background execution helper for slow calls into core/ (parsing a real TIO
workbook can take several seconds -- see core/excel_reader.py's
open_workbook docstring). Runs the call on a QThreadPool worker thread so
the UI thread's event loop stays responsive (the window keeps repainting,
menus keep working, etc.) instead of freezing for the whole duration, and
shows a window-modal QProgressDialog so the user gets clear feedback and
can't trigger the same action a second time while one is already running.
"""

from __future__ import annotations

import logging
from typing import Callable

from PySide6.QtCore import QObject, QRunnable, Qt, QThreadPool, Signal, Slot
from PySide6.QtWidgets import QMessageBox, QProgressDialog, QWidget

from core import app_logging

logger = logging.getLogger(__name__)


class _WorkerSignals(QObject):
    finished = Signal(object)
    error = Signal(Exception)


class _Worker(QRunnable):
    def __init__(self, fn: Callable, args: tuple, kwargs: dict, description: str):
        super().__init__()
        self.fn = fn
        self.args = args
        self.kwargs = kwargs
        self.description = description  # run_with_progress's `message` -- what the user was doing, for the log line
        self.signals = _WorkerSignals()
        # Qt's QThreadPool deletes the C++-side QRunnable as soon as run()
        # returns when autoDelete is on (the default) -- that races with
        # the Python object still being referenced (see run_with_progress,
        # which holds a reference until the dialog closes) and is a known
        # source of "wrapped C++ object has been deleted" crashes. Turning
        # it off just leaves cleanup to normal Python garbage collection.
        self.setAutoDelete(False)

    @Slot()
    def run(self):
        try:
            result = self.fn(*self.args, **self.kwargs)
        except Exception as exc:  # noqa: BLE001 -- intentionally broad, forwarded to on_error, never swallowed
            # logger.exception() (not just .error()) captures the full
            # traceback via sys.exc_info() -- has to happen HERE, still
            # inside the except block on the background thread, since
            # that context is gone by the time _error() below runs on
            # the UI thread. Every run_with_progress caller gets this
            # for free -- the whole point is that a caller's on_error
            # can be as simple as showing a dialog; it never has to
            # remember to log anything itself for the failure to end up
            # in the log file (see core/app_logging.py).
            fn_name = getattr(self.fn, "__qualname__", repr(self.fn))
            logger.exception("Background task failed (%r): %s", self.description, fn_name)
            self.signals.error.emit(exc)
        else:
            self.signals.finished.emit(result)


def run_with_progress(
    parent: QWidget,
    message: str,
    fn: Callable,
    on_finished: Callable[[object], None],
    on_error: Callable[[Exception], None],
    *args,
    **kwargs,
) -> None:
    """Runs fn(*args, **kwargs) on a background thread. Shows `message` in
    a window-modal, indeterminate progress dialog (no cancel button --
    these calls read/parse a file on a background thread and aren't
    interruptible mid-way) for as long as it takes. Window-modal blocks
    interaction with `parent`'s window specifically, which is exactly
    "the user can't trigger the same action twice" without having to
    manually track and disable individual buttons.

    on_finished(result) or on_error(exception) is called back on the UI
    thread once the call completes -- safe to touch widgets from there.
    """
    dialog = QProgressDialog(message, "", 0, 0, parent)
    dialog.setWindowModality(Qt.WindowModal)
    dialog.setWindowTitle("Please Wait")
    dialog.setMinimumDuration(0)  # show immediately -- these calls always take multiple seconds
    dialog.setCancelButton(None)
    dialog.show()

    worker = _Worker(fn, args, kwargs, description=message)

    def _report_callback_failure(which: str):
        # fn() itself already succeeded (or its failure already went
        # through _Worker.run()'s own logging) -- this catches on_finished/
        # on_error THEMSELVES raising, e.g. a widget failing to build
        # from data that parsed fine but turned out to be unexpected in
        # some way _Worker.run() had no way to anticipate. Without this,
        # that exception propagates straight into Qt's signal-delivery
        # machinery, which on a --windowed build (no console) swallows
        # it silently: the progress dialog has already closed by this
        # point, so the net effect is the screen just... does nothing,
        # with no error and no clue why -- exactly the "loading
        # indicator then nothing happens" failure mode this exists to
        # prevent, confirmed against a real file that reproduced it.
        logger.exception("%s callback failed after: %r", which, message)
        QMessageBox.critical(
            parent,
            "Something Went Wrong",
            "An unexpected error occurred while handling this. "
            f"Details were written to the log file:\n\n{app_logging.get_log_path()}",
        )

    def _finished(result):
        dialog.close()
        try:
            on_finished(result)
        except Exception:  # noqa: BLE001 -- last-resort catch-all, see _report_callback_failure
            _report_callback_failure("on_finished")

    def _error(exc):
        dialog.close()
        try:
            on_error(exc)
        except Exception:  # noqa: BLE001 -- last-resort catch-all, see _report_callback_failure
            _report_callback_failure("on_error")

    worker.signals.finished.connect(_finished)
    worker.signals.error.connect(_error)
    # Keep the worker (and its signals object) alive for as long as the
    # dialog is -- otherwise Python could garbage-collect it before the
    # background thread finishes.
    dialog._worker = worker

    QThreadPool.globalInstance().start(worker)
