import pytest

from app.convert import ConversionError, MSG_FAILED, MSG_NO_OFFICE, MSG_PASSWORD_WORD, part_path
from app.office import word_installed, find_libreoffice, office_available, docx_to_pdf

@pytest.fixture(autouse=True)
def _quiet_com_teardown():
    """Word's COM teardown after Quit() raises a handled RPC error
    (0x800706be) that faulthandler prints as a fatal exception; silence it for
    these tests only."""
    import faulthandler

    was_enabled = faulthandler.is_enabled()
    faulthandler.disable()
    yield
    if was_enabled:
        faulthandler.enable()


needs_word = pytest.mark.skipif(not word_installed(), reason="Word not installed")
needs_lo = pytest.mark.skipif(find_libreoffice() is None, reason="LibreOffice not installed")


def test_office_available_matches_parts():
    assert office_available() == (word_installed() or find_libreoffice() is not None)


@needs_word
def test_word_converts_docx(sample_docx, tmp_path):
    dst = tmp_path / "out.pdf"
    assert docx_to_pdf(sample_docx, dst) == dst
    assert not part_path(dst).exists()
    import pymupdf
    with pymupdf.open(dst) as pdf:
        assert "Hello from PDF Converter" in pdf[0].get_text()


def _save_with_word(src, dst, file_format=None, **protection):
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    try:
        word = win32com.client.DispatchEx("Word.Application")
        try:
            word.Visible = False
            word.DisplayAlerts = 0
            doc = word.Documents.Open(str(src))
            for name, value in protection.items():
                setattr(doc, name, value)
            if file_format is None:
                doc.SaveAs2(str(dst))
            else:
                doc.SaveAs2(str(dst), FileFormat=file_format)
            doc.Close(0)
            doc = None
        finally:
            word.Quit()
    finally:
        pythoncom.CoUninitialize()


@needs_word
def test_password_protected_docx_reports_message(sample_docx, tmp_path):
    protected = tmp_path / "locked.docx"
    _save_with_word(sample_docx, protected, Password="secret")
    with pytest.raises(ConversionError) as e:
        docx_to_pdf(protected, tmp_path / "out.pdf")
    assert e.value.message == MSG_PASSWORD_WORD


@needs_word
def test_write_protected_docx_reports_message(sample_docx, tmp_path):
    protected = tmp_path / "readonly.docx"
    _save_with_word(sample_docx, protected, WritePassword="secret")
    with pytest.raises(ConversionError) as e:
        docx_to_pdf(protected, tmp_path / "out.pdf")
    assert e.value.message == MSG_PASSWORD_WORD


@needs_word
def test_unprotected_legacy_doc_converts(sample_docx, tmp_path):
    legacy = tmp_path / "plain.doc"
    _save_with_word(sample_docx, legacy, file_format=0)
    dst = tmp_path / "out.pdf"
    docx_to_pdf(legacy, dst)
    import pymupdf
    with pymupdf.open(dst) as pdf:
        assert "Hello from PDF Converter" in pdf[0].get_text()


@needs_word
@pytest.mark.parametrize("protection", [{"Password": "secret"}, {"WritePassword": "secret"}])
def test_protected_legacy_doc_reports_message(sample_docx, tmp_path, protection):
    protected = tmp_path / "locked.doc"
    _save_with_word(sample_docx, protected, file_format=0, **protection)
    with pytest.raises(ConversionError) as e:
        docx_to_pdf(protected, tmp_path / "out.pdf")
    assert e.value.message == MSG_PASSWORD_WORD


@needs_word
def test_unreadable_ole_file_fails_closed(sample_docx, tmp_path, monkeypatch):
    import pythoncom

    protected = tmp_path / "locked.docx"
    _save_with_word(sample_docx, protected, Password="secret")

    def share_violation(*args):
        raise pythoncom.com_error(-2147287008, "sharing violation", None, None)

    monkeypatch.setattr(pythoncom, "StgOpenStorage", share_violation)
    with pytest.raises(ConversionError) as e:
        docx_to_pdf(protected, tmp_path / "out.pdf")
    assert e.value.message == MSG_FAILED


@needs_lo
def test_libreoffice_converts_docx(sample_docx, tmp_path, monkeypatch):
    monkeypatch.setattr("app.office.word_installed", lambda: False)
    docx_to_pdf(sample_docx, tmp_path / "out.pdf")
    assert (tmp_path / "out.pdf").exists()


def test_no_office_raises(sample_docx, tmp_path, monkeypatch):
    monkeypatch.setattr("app.office.word_installed", lambda: False)
    monkeypatch.setattr("app.office.find_libreoffice", lambda: None)
    with pytest.raises(ConversionError) as e:
        docx_to_pdf(sample_docx, tmp_path / "out.pdf")
    assert e.value.message == MSG_NO_OFFICE


class _FakeWord:
    def __init__(self):
        self.quit_calls = []

    def Quit(self, **kwargs):
        self.quit_calls.append(kwargs)
        raise RuntimeError("Word is gone")  # as after a terminate


def test_watchdog_terminates_hung_word(sample_docx, tmp_path, monkeypatch):
    import threading

    import pywintypes
    import win32com.client

    import app.office

    fake_word = _FakeWord()
    terminated = []
    killed = threading.Event()

    def record_terminate(pid):
        terminated.append(pid)
        killed.set()

    def hung_export(word, src, out_pdf):
        assert killed.wait(10), "watchdog did not fire"
        raise pywintypes.com_error(-2147023170, "The remote procedure call failed.", None, None)

    monkeypatch.setattr(app.office, "word_installed", lambda: True)
    monkeypatch.setattr(win32com.client, "DispatchEx", lambda name: fake_word)
    monkeypatch.setattr(app.office, "_word_pid", lambda word: 4242)
    monkeypatch.setattr(app.office, "WORD_TIMEOUT_SECONDS", 1)
    monkeypatch.setattr(app.office, "_terminate_word", record_terminate)
    monkeypatch.setattr(app.office, "_export_with_word", hung_export)
    dst = tmp_path / "out.pdf"
    with pytest.raises(ConversionError) as e:
        docx_to_pdf(sample_docx, dst)
    assert e.value.message == MSG_FAILED
    assert terminated == [4242]
    assert fake_word.quit_calls == [{"SaveChanges": 0}]
    assert not dst.exists() and not part_path(dst).exists()


@needs_word
def test_watchdog_terminates_real_word(sample_docx, tmp_path, monkeypatch):
    """A real hidden Word is terminated by the watchdog; the file fails."""
    import threading

    import win32api
    import win32event

    import app.office

    real_export = app.office._export_with_word
    real_terminate = app.office._terminate_word
    terminated = []
    killed = threading.Event()

    def terminate(pid):
        real_terminate(pid)
        terminated.append(pid)
        killed.set()

    def slow_export(word, src, out_pdf):
        # Stand-in for a hidden modal: do nothing until the watchdog fires,
        # then talk to the (now dead) Word, as a blocked call would return.
        assert killed.wait(15), "watchdog did not fire"
        real_export(word, src, out_pdf)

    monkeypatch.setattr(app.office, "WORD_TIMEOUT_SECONDS", 1)
    monkeypatch.setattr(app.office, "_terminate_word", terminate)
    monkeypatch.setattr(app.office, "_export_with_word", slow_export)
    dst = tmp_path / "out.pdf"
    with pytest.raises(ConversionError) as e:
        docx_to_pdf(sample_docx, dst)
    assert e.value.message == MSG_FAILED
    assert len(terminated) == 1
    assert not dst.exists() and not part_path(dst).exists()
    synchronize = 0x00100000
    try:
        handle = win32api.OpenProcess(synchronize, False, terminated[0])
    except win32api.error:
        return  # process already gone
    try:
        assert win32event.WaitForSingleObject(handle, 5000) == win32event.WAIT_OBJECT_0
    finally:
        win32api.CloseHandle(handle)


def _fake_libreoffice(monkeypatch, calls):
    """Pretend only LibreOffice is installed; soffice writes a small PDF."""
    from pathlib import Path

    monkeypatch.setattr("app.office.word_installed", lambda: False)
    monkeypatch.setattr("app.office.find_libreoffice", lambda: Path(r"C:\LO\soffice.exe"))

    def run(args, **kwargs):
        calls.append(args)
        outdir = Path(args[args.index("--outdir") + 1])
        (outdir / f"{Path(args[-1]).stem}.pdf").write_bytes(b"%PDF-1.4 converted")

    monkeypatch.setattr("app.office.subprocess.run", run)


def test_libreoffice_uses_private_profile(sample_docx, tmp_path, monkeypatch):
    from pathlib import Path
    from urllib.parse import unquote, urlparse

    calls = []
    _fake_libreoffice(monkeypatch, calls)
    dst = tmp_path / "out.pdf"
    assert docx_to_pdf(sample_docx, dst) == dst
    assert dst.read_bytes() == b"%PDF-1.4 converted"
    [args] = calls
    [option] = [a for a in args if a.startswith("-env:UserInstallation=")]
    url = option.split("=", 1)[1]
    assert url.startswith("file:///") and "\\" not in url
    profile = Path(unquote(urlparse(url).path).lstrip("/"))
    outdir = Path(args[args.index("--outdir") + 1])
    assert profile.parent == outdir


def test_word_to_pdf_never_overwrites_appeared_file(sample_docx, tmp_path, monkeypatch):
    _fake_libreoffice(monkeypatch, [])
    dst = tmp_path / "out.pdf"
    dst.write_bytes(b"someone else's file")
    written = docx_to_pdf(sample_docx, dst)
    assert written == tmp_path / "out (1).pdf"
    assert written.read_bytes() == b"%PDF-1.4 converted"
    assert dst.read_bytes() == b"someone else's file"
    assert not part_path(dst).exists()


def test_protected_docx_is_refused_before_libreoffice(sample_docx, tmp_path, monkeypatch):
    import zipfile

    protected = tmp_path / "readonly.docx"
    with zipfile.ZipFile(sample_docx) as src, zipfile.ZipFile(protected, "w") as out:
        for item in src.infolist():
            data = src.read(item.filename)
            if item.filename == "word/settings.xml":
                data = data.replace(
                    b"<w:zoom ",
                    b'<w:writeProtection w:cryptProviderType="rsaAES" '
                    b'w:hash="abc" w:salt="def"/><w:zoom ',
                    1,
                )
                assert b"writeProtection" in data
            out.writestr(item, data)
    calls = []
    _fake_libreoffice(monkeypatch, calls)
    with pytest.raises(ConversionError) as e:
        docx_to_pdf(protected, tmp_path / "out.pdf")
    assert e.value.message == MSG_PASSWORD_WORD
    assert calls == []
