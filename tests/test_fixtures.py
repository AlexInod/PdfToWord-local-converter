def test_text_pdf_has_three_pages_with_text(text_pdf):
    import pymupdf
    doc = pymupdf.open(text_pdf)
    assert doc.page_count == 3
    assert len(doc[0].get_text().strip()) > 20

def test_scanned_pdf_has_no_text(scanned_pdf):
    import pymupdf
    doc = pymupdf.open(scanned_pdf)
    assert doc[0].get_text().strip() == ""
    assert doc[0].get_images()

def test_encrypted_pdf_needs_password(encrypted_pdf):
    import pymupdf
    assert pymupdf.open(encrypted_pdf).needs_pass
