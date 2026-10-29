import shutil

import pytest

from app.convert import inspect_pdf, ConversionError, MSG_SCANNED, MSG_PASSWORD_PDF, MSG_FAILED


def test_text_pdf_page_count(text_pdf):
    assert inspect_pdf(text_pdf) == 3


def test_scanned_pdf_raises(scanned_pdf):
    with pytest.raises(ConversionError) as e:
        inspect_pdf(scanned_pdf)
    assert e.value.message == MSG_SCANNED


def test_mixed_pdf_is_not_scanned(mixed_pdf):
    assert inspect_pdf(mixed_pdf) == 3


def test_encrypted_pdf_raises(encrypted_pdf):
    with pytest.raises(ConversionError) as e:
        inspect_pdf(encrypted_pdf)
    assert e.value.message == MSG_PASSWORD_PDF


def test_corrupt_pdf_raises_failed(corrupt_pdf):
    with pytest.raises(ConversionError) as e:
        inspect_pdf(corrupt_pdf)
    assert e.value.message == MSG_FAILED


def test_unicode_name_is_fine(text_pdf, tmp_path):
    src = tmp_path / "Planificare română (test).PDF"
    shutil.copy(text_pdf, src)
    assert inspect_pdf(src) == 3
