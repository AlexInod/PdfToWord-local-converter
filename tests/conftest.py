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
