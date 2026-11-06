import queue
import threading
from typing import Any, Callable
import pythoncom


class OutlookComWorker:
    """
    Runs all pywin32 / Outlook COM calls on a single dedicated STA thread.
    Flask request threads submit callables via `run(fn, *args, **kwargs)`.
    """

    def __init__(self) -> None:
        self._queue: "queue.Queue[tuple[Callable[..., Any], tuple[Any, ...], dict[str, Any], queue.Queue[tuple[bool, Any]]]]" = queue.Queue()
        self._thread = threading.Thread(target=self._loop, daemon=True, name="OutlookCOMWorker")
        self._thread.start()

    def _loop(self) -> None:
        pythoncom.CoInitialize()
        try:
            while True:
                item = self._queue.get()
                if item is None:
                    break
                fn, args, kwargs, reply_q = item
                try:
                    result = fn(*args, **kwargs)
                    reply_q.put((True, result))
                except Exception as exc:
                    reply_q.put((False, exc))
        finally:
            pythoncom.CoUninitialize()

    def run(self, fn: Callable[..., Any], *args: Any, **kwargs: Any) -> Any:
        reply_q: "queue.Queue[tuple[bool, Any]]" = queue.Queue(maxsize=1)
        self._queue.put((fn, args, kwargs, reply_q))
        ok, payload = reply_q.get()
        if ok:
            return payload
        raise payload


com_worker = OutlookComWorker()
