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

from typing import Callable

from PySide6.QtCore import QObject, QRunnable, Qt, QThreadPool, Signal, Slot
from PySide6.QtWidgets import QProgressDialog, QWidget


class _WorkerSignals(QObject):
    finished = Signal(object)
    error = Signal(Exception)


class _Worker(QRunnable):
    def __init__(self, fn: Callable, args: tuple, kwargs: dict):
        super().__init__()
        self.fn = fn
        self.args = args
        self.kwargs = kwargs
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

    worker = _Worker(fn, args, kwargs)

    def _finished(result):
        dialog.close()
        on_finished(result)

    def _error(exc):
        dialog.close()
        on_error(exc)

    worker.signals.finished.connect(_finished)
    worker.signals.error.connect(_error)
    # Keep the worker (and its signals object) alive for as long as the
    # dialog is -- otherwise Python could garbage-collect it before the
    # background thread finishes.
    dialog._worker = worker

    QThreadPool.globalInstance().start(worker)
