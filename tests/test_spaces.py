import pytest
from docx import Document
from pdf2docx.text.Spans import Spans

from app.convert import _merge_blank_spans, _patch_pdf2docx_blank_spans, pdf_to_docx


def span(text, x0, x1, **extra):
    chars = [{"c": c, "bbox": (x0, 0.0, x1, 10.0)} for c in text]
    return {"chars": chars, "bbox": (x0, 0.0, x1, 10.0), "font": "Helvetica", "size": 12, **extra}


def text_of(s):
    return "".join(c["c"] for c in s["chars"])


def test_blank_span_merges_into_previous_text():
    raws = [span("Hello", 0, 30), span(" ", 30, 34), span("there", 34, 60)]
    merged = _merge_blank_spans(raws)
    assert [text_of(s) for s in merged] == ["Hello ", "there"]
    assert merged[0]["bbox"] == (0, 0.0, 34, 10.0)
    assert merged[1] is raws[2]
    assert text_of(raws[0]) == "Hello"


def test_styled_blank_span_is_kept():
    raws = [span("Hello", 0, 30), span(" ", 30, 34, style=[{"type": 1}])]
    assert _merge_blank_spans(raws) == raws


def test_leading_blank_span_is_kept():
    raws = [span(" ", 0, 4), span("Hello", 4, 30)]
    merged = _merge_blank_spans(raws)
    assert merged[0] is raws[0] and merged[1] is raws[1]


def test_image_span_is_untouched_and_not_merged_into():
    image = {"image": b"x", "bbox": (0, 0, 5, 5)}
    raws = [image, span(" ", 5, 9), span("a", 9, 12)]
    merged = _merge_blank_spans(raws)
    assert merged == raws and merged[0] is image


def test_patch_is_idempotent():
    original = Spans.restore
    _patch_pdf2docx_blank_spans()
    _patch_pdf2docx_blank_spans()
    assert Spans.restore._merges_blank_spans is True
    inner = Spans.restore.__wrapped__
    assert not getattr(inner, "_merges_blank_spans", False)
    if not getattr(original, "_merges_blank_spans", False):
        assert inner is original


def test_pdf_to_docx_keeps_spaces_between_words(word_spans_pdf, tmp_path):
    dst = tmp_path / "out.docx"
    pdf_to_docx(word_spans_pdf, dst)
    text = "\n".join(p.text for p in Document(dst).paragraphs)
    assert "Hello there world" in text
    assert "Another line here" in text
