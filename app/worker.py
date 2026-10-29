import logging
import queue
import threading
from collections.abc import Callable, Iterable
from dataclasses import dataclass
from pathlib import Path

from app.convert import (
    MSG_FAILED,
    PDF_TO_WORD,
    ConversionError,
    output_path_for,
    part_path,
    pdf_to_docx,
)
from app.office import docx_to_pdf

WAITING = "waiting"
CONVERTING = "converting"
DONE = "done"
FAILED = "failed"


@dataclass
class Job:
    id: int
    source: Path
    mode: str
    status: str = WAITING
    message: str = ""
    page: int = 0
    pages: int = 0
    output: Path | None = None

    def to_dict(self) -> dict:
        return {
            "id": self.id,
            "name": self.source.name,
            "status": self.status,
            "message": self.message,
            "page": self.page,
            "pages": self.pages,
        }


class ConversionWorker:
    def __init__(self, on_update: Callable[[Job], None]):
        self._on_update = on_update
        self._queue: queue.Queue[Job] = queue.Queue()
        self._lock = threading.Lock()
        self._thread: threading.Thread | None = None
        self._stopping = False
        self._current: Job | None = None
        self._next_id = 1
        self.jobs: dict[int, Job] = {}
        self.last_output: Path | None = None

    def add(self, paths: Iterable[str | Path], mode: str) -> list[Job]:
        new_jobs = []
        with self._lock:
            for raw in paths:
                path = Path(raw)
                if path.is_dir():
                    continue
                job = Job(id=self._next_id, source=path, mode=mode)
                self._next_id += 1
                self.jobs[job.id] = job
                new_jobs.append(job)
            if new_jobs and self._thread is None:
                self._thread = threading.Thread(target=self._run, daemon=True)
                self._thread.start()
        for job in new_jobs:
            self._notify(job)
            self._queue.put(job)
        return new_jobs

    @property
    def current_mode(self) -> str | None:
        """Mode of the job being converted right now, or None."""
        with self._lock:
            return self._current.mode if self._current is not None else None

    def join(self, timeout: float) -> None:
        """Wait up to timeout seconds for the worker thread to finish."""
        thread = self._thread
        if thread is not None:
            thread.join(timeout)

    def stop(self) -> None:
        self._stopping = True
        while True:
            try:
                self._queue.get_nowait()
            except queue.Empty:
                break
        with self._lock:
            current = self._current
        if current is not None and current.output is not None:
            part_path(current.output).unlink(missing_ok=True)

    def _notify(self, job: Job) -> None:
        try:
            self._on_update(job)
        except Exception:
            logging.exception("on_update callback failed")

    def _run(self) -> None:
        while not self._stopping:
            try:
                job = self._queue.get(timeout=0.2)
            except queue.Empty:
                continue
            if self._stopping:
                break
            self._process(job)

    def _process(self, job: Job) -> None:
        with self._lock:
            self._current = job
        try:
            job.output = output_path_for(job.source, job.mode)
            job.status = CONVERTING
            self._notify(job)
            if job.mode == PDF_TO_WORD:
                written = pdf_to_docx(job.source, job.output, lambda page, pages: self._progress(job, page, pages))
            else:
                written = docx_to_pdf(job.source, job.output)
            if self._stopping:
                return
            job.output = written
            job.status = DONE
            self.last_output = job.output
        except ConversionError as error:
            job.status = FAILED
            job.message = error.message
        except Exception:
            logging.exception("Conversion of %s failed", job.source)
            job.status = FAILED
            job.message = MSG_FAILED
        finally:
            if self._stopping and job.output is not None:
                part_path(job.output).unlink(missing_ok=True)
            with self._lock:
                self._current = None
        self._notify(job)

    def _progress(self, job: Job, page: int, pages: int) -> None:
        job.page = page
        job.pages = pages
        self._notify(job)
