import json
import logging
import os
import subprocess
import threading
from pathlib import Path

import webview
from webview.dom import DOMEventHandler

from app.convert import PDF_TO_WORD, WORD_TO_PDF
from app.office import WORD_TIMEOUT_SECONDS, office_available
from app.startup import (
    close_splash,
    configure_environment,
    log_path,
    show_missing_webview2,
    show_startup_failed,
    webview2_version,
)
from app.worker import ConversionWorker

FILE_TYPES = {
    PDF_TO_WORD: ("PDF files (*.pdf)",),
    WORD_TO_PDF: ("Word files (*.docx;*.doc)",),
}


def _explorer_select_command(path: Path) -> str:
    """Explorer needs '/select,' and the quoted path as one argument."""
    return f'explorer /select,"{path}"'


def drop_paths(event: dict) -> list[str]:
    files = (event.get("dataTransfer") or {}).get("files") or []
    return [f["pywebviewFullPath"] for f in files if f.get("pywebviewFullPath")]


class Api:
    def __init__(self):
        self._mode = PDF_TO_WORD
        self._worker: ConversionWorker | None = None
        self._window = None

    def set_mode(self, mode: str) -> None:
        if mode in (PDF_TO_WORD, WORD_TO_PDF):
            self._mode = mode

    def choose_files(self) -> None:
        result = self._window.create_file_dialog(
            webview.FileDialog.OPEN,
            allow_multiple=True,
            file_types=FILE_TYPES[self._mode],
        )
        if result:
            self._worker.add(result, self._mode)

    def open_folder(self) -> None:
        output = self._worker.last_output
        if output is not None:
            os.startfile(output.parent)

    def show_file(self, job_id) -> None:
        job = self._worker.jobs.get(int(job_id))
        if job is not None and job.output is not None:
            subprocess.Popen(_explorer_select_command(job.output))

    def office_available(self) -> bool:
        return office_available()


def _configure_logging() -> None:
    logging.basicConfig(
        filename=log_path(),
        level=logging.WARNING,
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
    )
    # A conversion temporarily raises the root level to INFO; keep the file quiet.
    for handler in logging.getLogger().handlers:
        handler.setLevel(logging.WARNING)


def _startup_log_location() -> str:
    try:
        return str(log_path())
    except Exception:
        return r"%LOCALAPPDATA%\PDF Converter\app.log"


def main() -> None:
    try:
        _configure_logging()
        configure_environment()
        if webview2_version() is None:
            close_splash()
            show_missing_webview2()
            return
        _run()
    except Exception:
        logging.exception("PDF Converter could not start")
        close_splash()
        show_startup_failed(_startup_log_location())


def _run() -> None:
    api = Api()
    window_gone = threading.Event()

    def push(job) -> None:
        # After closing starts, the window may be gone or going; evaluate_js
        # could then wait forever, so updates are dropped.
        if window_gone.is_set():
            return
        try:
            api._window.evaluate_js(f"window.app.onJob({json.dumps(job.to_dict())})")
        except Exception:
            logging.exception("Could not push job update")

    worker = ConversionWorker(on_update=push)
    api._worker = worker
    window = webview.create_window(
        "PDF Converter",
        url=str(Path(__file__).parent / "ui" / "index.html"),
        js_api=api,
        width=520,
        height=600,
        min_size=(420, 480),
    )
    api._window = window

    def on_drop(event) -> None:
        try:
            worker.add(drop_paths(event), api._mode)
        except Exception:
            logging.exception("Drop failed")

    def on_loaded() -> None:
        try:
            window.dom.get_element("#drop").on(
                "drop", DOMEventHandler(on_drop, prevent_default=True)
            )
        except Exception:
            logging.exception("Could not register drop handler")

    def on_shown() -> None:
        close_splash()

    def on_closing() -> None:
        # Runs on the UI thread: never wait for the worker here, its pushes
        # need the UI thread.
        window_gone.set()
        try:
            worker.stop()
        except Exception:
            logging.exception("Stopping worker failed")

    window.events.loaded += on_loaded
    window.events.shown += on_shown
    window.events.closing += on_closing
    webview.start()
    window_gone.set()
    # The worker is a daemon thread; let a Word conversion finish so Word is
    # quit instead of left running hidden. The watchdog bounds the wait.
    if worker.current_mode == WORD_TO_PDF:
        worker.join(WORD_TIMEOUT_SECONDS + 10)
