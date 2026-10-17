"""Background task adapter with explicit cancellation state."""
from __future__ import annotations
import threading
from PySide6.QtCore import QObject, QRunnable, Signal

class WorkerSignals(QObject):
    succeeded = Signal(int, object)
    failed = Signal(int, str)
    progress = Signal(str)
    cancelled = Signal(int)

class Worker(QRunnable):
    def __init__(self, task_id, function):
        super().__init__()
        self.task_id = task_id
        self.function = function
        self.cancel_event = threading.Event()
        self.signals = WorkerSignals()

    def run(self):
        try:
            if self.cancel_event.is_set():
                self.signals.cancelled.emit(self.task_id)
                return
            result = self.function(self.signals.progress.emit, self.cancel_event)
            if self.cancel_event.is_set():
                self.signals.cancelled.emit(self.task_id)
            else:
                self.signals.succeeded.emit(self.task_id, result)
        except Exception as exc:
            if self.cancel_event.is_set():
                self.signals.cancelled.emit(self.task_id)
            else:
                self.signals.failed.emit(self.task_id, str(exc))

