import logging
import shutil

import pytest

from app.convert import (
    pdf_to_docx, part_path, output_path_for, ConversionError,
    PDF_TO_WORD, MSG_SCANNED, MSG_CANNOT_SAVE, _PageProgressHandler,
)


def test_converts_and_reports_progress(text_pdf, tmp_path):
    dst = tmp_path / "out.docx"
    seen = []
    pdf_to_docx(text_pdf, dst, progress=lambda p, n: seen.append((p, n)))
    assert dst.exists()
    assert not part_path(dst).exists()
    assert seen[-1] == (3, 3)
    assert seen == sorted(seen)           # never goes backwards


def test_scanned_pdf_is_rejected_before_writing(scanned_pdf, tmp_path):
    dst = tmp_path / "out.docx"
    with pytest.raises(ConversionError) as e:
        pdf_to_docx(scanned_pdf, dst)
    assert e.value.message == MSG_SCANNED
    assert not dst.exists() and not part_path(dst).exists()


def test_unwritable_destination_reports_cannot_save(text_pdf, tmp_path):
    dst = tmp_path / "missing-folder" / "out.docx"
    with pytest.raises(ConversionError) as e:
        pdf_to_docx(text_pdf, dst)
    assert e.value.message == MSG_CANNOT_SAVE


def test_unicode_and_spaces_round_trip(text_pdf, tmp_path):
    src = tmp_path / "Planificare română (test).pdf"
    shutil.copy(text_pdf, src)
    dst = output_path_for(src, PDF_TO_WORD)
    pdf_to_docx(src, dst)
    assert dst == tmp_path / "Planificare română (test).docx" and dst.exists()


def test_root_logger_level_is_restored(text_pdf, tmp_path):
    before = logging.getLogger().level
    pdf_to_docx(text_pdf, tmp_path / "out.docx")
    assert logging.getLogger().level == before


def _assert_logging_clean(level_before):
    root = logging.getLogger()
    assert root.level == level_before
    assert not [h for h in root.handlers if isinstance(h, _PageProgressHandler)]


def test_logging_state_restored_after_success_with_progress(text_pdf, tmp_path):
    before = logging.getLogger().level
    pdf_to_docx(text_pdf, tmp_path / "out.docx", progress=lambda p, n: None)
    _assert_logging_clean(before)


def test_logging_state_restored_after_failure_with_progress(text_pdf, tmp_path):
    before = logging.getLogger().level
    with pytest.raises(ConversionError):
        pdf_to_docx(
            text_pdf,
            tmp_path / "missing-folder" / "out.docx",
            progress=lambda p, n: None,
        )
    _assert_logging_clean(before)
