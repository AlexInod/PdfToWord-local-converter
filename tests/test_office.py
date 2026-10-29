import pytest

from app.convert import ConversionError, MSG_NO_OFFICE, MSG_PASSWORD_WORD
from app.office import word_installed, find_libreoffice, office_available, docx_to_pdf

needs_word = pytest.mark.skipif(not word_installed(), reason="Word not installed")
needs_lo = pytest.mark.skipif(find_libreoffice() is None, reason="LibreOffice not installed")


def test_office_available_matches_parts():
    assert office_available() == (word_installed() or find_libreoffice() is not None)


@needs_word
def test_word_converts_docx(sample_docx, tmp_path):
    dst = tmp_path / "out.pdf"
    docx_to_pdf(sample_docx, dst)
    import pymupdf
    with pymupdf.open(dst) as pdf:
        assert "Hello from PDF Converter" in pdf[0].get_text()


@needs_word
def test_password_protected_docx_reports_message(sample_docx, tmp_path):
    import pythoncom
    import win32com.client

    protected = tmp_path / "locked.docx"
    pythoncom.CoInitialize()
    try:
        word = win32com.client.DispatchEx("Word.Application")
        try:
            word.Visible = False
            word.DisplayAlerts = 0
            doc = word.Documents.Open(str(sample_docx))
            doc.Password = "secret"
            doc.SaveAs2(str(protected))
            doc.Close(0)
        finally:
            word.Quit()
    finally:
        pythoncom.CoUninitialize()
    with pytest.raises(ConversionError) as e:
        docx_to_pdf(protected, tmp_path / "out.pdf")
    assert e.value.message == MSG_PASSWORD_WORD


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
