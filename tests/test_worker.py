import threading
import time

from app.convert import MSG_NOT_PDF, PDF_TO_WORD, part_path
from app.worker import ConversionWorker


def wait_for(pred, timeout=60):
    end = time.time() + timeout
    while time.time() < end:
        if pred():
            return
        time.sleep(0.05)
    raise AssertionError("timeout")


def test_same_pdf_twice_gets_numbered_output(text_pdf):
    updates = []
    w = ConversionWorker(on_update=lambda j: updates.append(j.to_dict()))
    jobs = w.add([text_pdf, text_pdf], PDF_TO_WORD)
    wait_for(lambda: all(j.status in ("done", "failed") for j in jobs))
    assert [j.status for j in jobs] == ["done", "done"]
    assert jobs[0].output.name == "sample.docx" and jobs[1].output.name == "sample (1).docx"
    assert w.last_output == jobs[1].output
    assert any(u["status"] == "converting" and u["pages"] == 3 for u in updates)


def test_wrong_type_fails_with_message(sample_docx):
    w = ConversionWorker(on_update=lambda j: None)
    [job] = w.add([sample_docx], PDF_TO_WORD)
    wait_for(lambda: job.status == "failed")
    assert job.message == MSG_NOT_PDF


def test_directories_are_ignored(tmp_path):
    w = ConversionWorker(on_update=lambda j: None)
    assert w.add([tmp_path], PDF_TO_WORD) == []


def test_stop_removes_partial_output(text_pdf, monkeypatch):
    started, release = threading.Event(), threading.Event()

    def fake(src, dst, progress=None):
        part_path(dst).touch()
        started.set()
        release.wait()

    monkeypatch.setattr("app.worker.pdf_to_docx", fake)
    w = ConversionWorker(on_update=lambda j: None)
    [job] = w.add([text_pdf], PDF_TO_WORD)
    started.wait(5)
    w.stop()
    assert not part_path(job.output).exists()
    release.set()
