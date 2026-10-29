import docx
import pymupdf
import pytest

PARAGRAPH = (
    "This is a sample paragraph used to test the PDF converter. "
    "It contains enough real text to be extracted from the page."
)


def _add_text_page(doc, number):
    page = doc.new_page()
    page.insert_text((72, 100), f"Page {number}. {PARAGRAPH}", fontsize=11)


def _add_scanned_page(doc):
    page = doc.new_page()
    pix = pymupdf.Pixmap(pymupdf.csRGB, pymupdf.IRect(0, 0, 100, 140), False)
    pix.set_rect(pix.irect, (255, 255, 255))
    page.insert_image(page.rect, pixmap=pix)


@pytest.fixture
def text_pdf(tmp_path):
    path = tmp_path / "sample.pdf"
    doc = pymupdf.open()
    for number in range(1, 4):
        _add_text_page(doc, number)
    doc.save(path)
    doc.close()
    return path


@pytest.fixture
def scanned_pdf(tmp_path):
    path = tmp_path / "scanned.pdf"
    doc = pymupdf.open()
    for _ in range(2):
        _add_scanned_page(doc)
    doc.save(path)
    doc.close()
    return path


@pytest.fixture
def mixed_pdf(tmp_path):
    path = tmp_path / "mixed.pdf"
    doc = pymupdf.open()
    _add_scanned_page(doc)
    _add_text_page(doc, 2)
    _add_text_page(doc, 3)
    doc.save(path)
    doc.close()
    return path


@pytest.fixture
def encrypted_pdf(tmp_path, text_pdf):
    path = tmp_path / "encrypted.pdf"
    doc = pymupdf.open(text_pdf)
    doc.save(
        path,
        encryption=pymupdf.PDF_ENCRYPT_AES_256,
        user_pw="secret",
        owner_pw="secret",
    )
    doc.close()
    return path


@pytest.fixture
def corrupt_pdf(tmp_path):
    path = tmp_path / "corrupt.pdf"
    path.write_bytes(b"not a pdf")
    return path


@pytest.fixture
def sample_docx(tmp_path):
    path = tmp_path / "sample.docx"
    document = docx.Document()
    document.add_paragraph("Hello from PDF Converter")
    document.save(path)
    return path


# Lines of near-equal width, so pdf2docx joins them into one paragraph. No
# line ends in a space, as in LaTeX output. Avoid "fi"/"fl": the font ligates them.
WRAPPED_LINES = [
    "The audience would soon get to meet Peter Parker",
    "„at the beginning of a story about a hero who is more",
    "„spider” than „man”, and this paper gives a short pre-",
    "sentation of the movie to the readers of a whole clasă",
    "întreagă of students.",
]


@pytest.fixture
def wrapped_lines_pdf(tmp_path):
    """One paragraph whose line ends need a space or lose a hyphen."""
    path = tmp_path / "wrapped.pdf"
    doc = pymupdf.open()
    page = doc.new_page()
    html = "<p style='font-family:serif;font-size:12pt'>" + "<br>".join(WRAPPED_LINES) + "</p>"
    page.insert_htmlbox(pymupdf.Rect(72, 72, 523, 400), html)
    doc.save(path)
    doc.close()
    return path


TOC_ENTRIES = [
    (72, "Foreword", "3"),
    (72, "1 Presentation of the Film", "4"),
    (90, "1.1 General presentation", "4"),
    (90, "1.2 The plot in short", "5"),
    (72, "2 The Harmony Between the Heroic and the Emotional, and the Actors Behind It", "8"),
    (90, "2.1 Heroism as a choice of the heart", "8"),
]


@pytest.fixture
def toc_pdf(tmp_path):
    """Table of contents with typed dot leaders, as LaTeX and Word draw them."""
    path = tmp_path / "toc.pdf"
    doc = pymupdf.open()
    page = doc.new_page()
    y = 100.0
    for x, title, number in TOC_ENTRIES:
        page.insert_text((x, y), title, fontsize=12, fontname="tiro")
        end = x + pymupdf.get_text_length(title, fontname="tiro", fontsize=12) + 8
        count = int((480 - end) / 6)
        if count >= 4:  # a long title leaves no room for dots
            page.insert_text((end, y), " ".join("." * count), fontsize=12, fontname="tiro")
        page.insert_text((515, y), number, fontsize=12, fontname="tiro")
        y += 30
    doc.save(path)
    doc.close()
    return path


BIBLIOGRAPHY = [
    ("1. Bryant, Jacob. The Punisher Joins the Cast of the New Movie With Tom",
     "Holland. TheWrap, 20 June 2025."),
    ("2. Couprie, Quentin. The most grounded hero yet: reviews land two days",
     "before release. 3DVF, July 2026."),
]


@pytest.fixture
def hanging_indent_pdf(tmp_path):
    """Numbered entries whose wrapped lines are indented under the text."""
    path = tmp_path / "bibliography.pdf"
    doc = pymupdf.open()
    page = doc.new_page()
    y = 100.0
    for first, second in BIBLIOGRAPHY:
        page.insert_text((72, y), first, fontsize=11, fontname="tiro")
        page.insert_text((88, y + 13), second, fontsize=11, fontname="tiro")
        y += 30
    doc.save(path)
    doc.close()
    return path


@pytest.fixture
def word_spans_pdf(tmp_path):
    path = tmp_path / "words.pdf"
    doc = pymupdf.open()
    page = doc.new_page()
    size = 12
    for y, words in ((100.0, ["Hello", "there", "world"]), (140.0, ["Another", "line", "here"])):
        x = 72.0
        for w in words:
            page.insert_text((x, y), w, fontsize=size, fontname="helv")
            x += pymupdf.get_text_length(w, fontname="helv", fontsize=size) + 0.33 * size
    doc.save(path)
    doc.close()
    return path
