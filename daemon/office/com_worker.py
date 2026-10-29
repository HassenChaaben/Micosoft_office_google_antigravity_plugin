"""
One dedicated STA thread for all Word COM calls.

COM objects are bound to the apartment (thread) that created them, so every Word call
must run on the same thread. The asyncio daemon submits jobs with `await worker.run(fn)`;
the job runs on the worker thread and its result (or exception) comes back as a future.
"""

import asyncio
import concurrent.futures
import queue
import threading
from typing import Any, Callable, Optional
from urllib.parse import unquote, urlparse


def _default_app_factory():
    """Attach to the user's running Word (never starts a new one)."""
    import win32com.client

    return win32com.client.GetActiveObject("Word.Application")


def _norm(path: str) -> str:
    """Lower-case, URL-decoded, forward-slash form of a path or URL."""
    p = unquote(str(path or "")).strip().replace("\\", "/")
    return p.rstrip("/").lower()


def _basename(path: str) -> str:
    p = _norm(path)
    if p.startswith(("http://", "https://")):
        p = urlparse(p).path
    return p.rsplit("/", 1)[-1]


class ComWorker:
    """Runs callables on a single COM-initialised (STA) thread."""

    def __init__(self, app_factory: Optional[Callable[[], Any]] = None, name: str = "word-com-worker"):
        self._app_factory = app_factory or _default_app_factory
        self._name = name
        self._jobs: "queue.Queue" = queue.Queue()
        self._thread: Optional[threading.Thread] = None
        self._thread_id: Optional[int] = None
        self._app = None
        self._ready = threading.Event()

    # --- lifecycle ------------------------------------------------------------------

    def start(self) -> "ComWorker":
        if self._thread is None or not self._thread.is_alive():
            self._ready.clear()
            self._thread = threading.Thread(target=self._loop, name=self._name, daemon=True)
            self._thread.start()
            self._ready.wait()
        return self

    def stop(self, timeout: float = 10.0) -> None:
        if self._thread is not None and self._thread.is_alive():
            self._jobs.put(None)
            self._thread.join(timeout)
        self._thread = None

    def _loop(self) -> None:
        import pythoncom

        pythoncom.CoInitialize()
        self._thread_id = threading.get_ident()
        self._ready.set()
        try:
            while True:
                item = self._jobs.get()
                if item is None:
                    break
                fn, args, kwargs, fut = item
                if not fut.set_running_or_notify_cancel():
                    continue
                try:
                    fut.set_result(fn(*args, **kwargs))
                except BaseException as exc:  # delivered to the awaiting caller
                    fut.set_exception(exc)
        finally:
            self._app = None
            pythoncom.CoUninitialize()

    # --- job submission ---------------------------------------------------------------

    @property
    def on_worker_thread(self) -> bool:
        return threading.get_ident() == self._thread_id

    def submit(self, fn: Callable, *args, **kwargs) -> concurrent.futures.Future:
        """Queue `fn(*args, **kwargs)` on the STA thread; returns a concurrent future."""
        if self._thread is None or not self._thread.is_alive():
            self.start()
        fut: concurrent.futures.Future = concurrent.futures.Future()
        if self.on_worker_thread:  # re-entrant call from a job: run inline, never deadlock
            try:
                fut.set_result(fn(*args, **kwargs))
            except BaseException as exc:
                fut.set_exception(exc)
            return fut
        self._jobs.put((fn, args, kwargs, fut))
        return fut

    async def run(self, fn: Callable, *args, **kwargs) -> Any:
        """Await `fn(*args, **kwargs)` executed on the STA thread."""
        return await asyncio.wrap_future(self.submit(fn, *args, **kwargs))

    def call(self, fn: Callable, *args, timeout: Optional[float] = None, **kwargs) -> Any:
        """Blocking variant of `run` for synchronous callers."""
        return self.submit(fn, *args, **kwargs).result(timeout)

    # --- Word access (only valid inside a job) -----------------------------------------

    def _require_worker(self) -> None:
        if not self.on_worker_thread:
            raise RuntimeError("Word COM objects may only be used inside a ComWorker job")

    def app(self):
        """The Word.Application, re-attached if the cached one died (e.g. Word restarted)."""
        self._require_worker()
        if self._app is not None:
            try:
                _ = self._app.Version
                return self._app
            except Exception:
                self._app = None
        self._app = self._app_factory()
        return self._app

    def get_document(self, doc_url_or_name: Optional[str] = None):
        """Find an open document by full path, https/OneDrive URL or file name; else the active one."""
        app = self.app()
        docs = app.Documents
        if docs.Count == 0:
            raise RuntimeError("No Word document is open")
        if doc_url_or_name:
            wanted = _norm(doc_url_or_name)
            wanted_name = _basename(doc_url_or_name)
            by_name = None
            for i in range(1, docs.Count + 1):
                d = docs(i)
                full = _norm(d.FullName)
                if full == wanted:
                    return d
                if by_name is None and wanted_name and (_basename(d.FullName) == wanted_name
                                                        or _norm(d.Name) == wanted_name):
                    by_name = d
            if by_name is not None:
                return by_name
        return app.ActiveDocument


__all__ = ["ComWorker"]
