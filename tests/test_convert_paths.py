from pathlib import Path

import pytest

from app.convert import (
    MSG_NOT_PDF,
    MSG_NOT_WORD,
    PDF_TO_WORD,
    WORD_TO_PDF,
    ConversionError,
    output_path_for,
    part_path,
    unique_path,
)


def test_unique_path_returns_same_when_free(tmp_path):
    assert unique_path(tmp_path / "Report.docx") == tmp_path / "Report.docx"


def test_unique_path_appends_counter(tmp_path):
    (tmp_path / "Report.docx").touch()
    (tmp_path / "Report (1).docx").touch()
    assert unique_path(tmp_path / "Report.docx") == tmp_path / "Report (2).docx"


def test_output_path_pdf_mode(tmp_path):
    assert output_path_for(tmp_path / "Report.PDF", PDF_TO_WORD) == tmp_path / "Report.docx"


def test_output_path_word_mode_accepts_doc_and_docx(tmp_path):
    assert output_path_for(tmp_path / "a.doc", WORD_TO_PDF) == tmp_path / "a.pdf"
    assert output_path_for(tmp_path / "b.docx", WORD_TO_PDF) == tmp_path / "b.pdf"


def test_wrong_extension_raises_mode_message(tmp_path):
    with pytest.raises(ConversionError) as e:
        output_path_for(tmp_path / "a.docx", PDF_TO_WORD)
    assert e.value.message == MSG_NOT_PDF
    with pytest.raises(ConversionError) as e:
        output_path_for(tmp_path / "a.pdf", WORD_TO_PDF)
    assert e.value.message == MSG_NOT_WORD


def test_part_path():
    assert part_path(Path("D:/x/Report.docx")) == Path("D:/x/Report.docx.part")
