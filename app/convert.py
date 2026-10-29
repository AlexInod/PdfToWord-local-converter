"""PDF <-> Word conversion helpers. Heavy libraries are imported lazily."""
import logging
import os
from pathlib import Path
from typing import Callable

PDF_TO_WORD = "pdf_to_word"
WORD_TO_PDF = "word_to_pdf"

MSG_SCANNED = "Scanned PDF, not supported yet"
MSG_PASSWORD_PDF = "Password-protected PDF"
MSG_PASSWORD_WORD = "Password-protected Word file"
MSG_NO_OFFICE = "Word or LibreOffice must be installed for Word to PDF"
MSG_NOT_PDF = "Not a PDF file"
MSG_NOT_WORD = "Not a Word file"
MSG_CANNOT_SAVE = "Could not save next to this file"
MSG_FAILED = "Could not convert this file"


class ConversionError(Exception):
    """Conversion failed; `message` is the user-facing text."""

    def __init__(self, message: str):
        super().__init__(message)
        self.message = message


def unique_path(path: Path) -> Path:
    """Return path, or 'stem (N).suffix' for the first N that does not exist."""
    if not path.exists():
        return path
    n = 1
    while True:
        candidate = path.with_name(f"{path.stem} ({n}){path.suffix}")
        if not candidate.exists():
            return candidate
        n += 1


def output_path_for(source: Path, mode: str) -> Path:
    """Validate the source extension for `mode`; return a free output path."""
    ext = source.suffix.lower()
    if mode == PDF_TO_WORD:
        if ext != ".pdf":
            raise ConversionError(MSG_NOT_PDF)
        return unique_path(source.with_suffix(".docx"))
    if mode == WORD_TO_PDF:
        if ext not in (".docx", ".doc"):
            raise ConversionError(MSG_NOT_WORD)
        return unique_path(source.with_suffix(".pdf"))
    raise ValueError(f"unknown mode: {mode}")


def part_path(dst: Path) -> Path:
    """Temporary name used while writing dst."""
    return dst.with_name(dst.name + ".part")


def _is_scanned_page(page) -> bool:
    """True when the page has almost no text and one image covers most of it."""
    if len(page.get_text().strip()) >= 20:
        return False
    page_area = page.rect.get_area()
    if page_area <= 0:
        return False
    for image in page.get_images(full=True):
        for rect in page.get_image_rects(image[0]):
            if rect.get_area() >= 0.8 * page_area:
                return True
    return False


def inspect_pdf(path: Path) -> int:
    """Return the page count; raise ConversionError for unusable PDFs."""
    import pymupdf

    try:
        doc = pymupdf.open(path)
    except Exception:
        raise ConversionError(MSG_FAILED) from None
    with doc:
        if doc.needs_pass:
            raise ConversionError(MSG_PASSWORD_PDF)
        count = doc.page_count
        if count == 0:
            raise ConversionError(MSG_FAILED)
        scanned = sum(1 for page in doc if _is_scanned_page(page))
    if scanned * 2 > count:
        raise ConversionError(MSG_SCANNED)
    return count


_PAGE_RECORD = "(%d/%d) Page %d"
_CREATING_MARKER = "[4/4] Creating pages"


class _PageProgressHandler(logging.Handler):
    """Forward pdf2docx's per-page log records from the parsing pass only."""

    def __init__(self, progress: Callable[[int, int], None]):
        super().__init__(level=logging.INFO)
        self._progress = progress
        self._parsing_done = False

    def emit(self, record: logging.LogRecord) -> None:
        if self._parsing_done:
            return
        if isinstance(record.msg, str) and _CREATING_MARKER in record.msg:
            self._parsing_done = True
            return
        if record.msg != _PAGE_RECORD or len(record.args) < 2:
            return
        page, pages = record.args[0], record.args[1]
        if page == pages:
            self._parsing_done = True
        self._progress(page, pages)


def pdf_to_docx(
    src: Path,
    dst: Path,
    progress: Callable[[int, int], None] | None = None,
) -> None:
    """Convert a PDF to DOCX at dst, reporting progress(page, pages)."""
    inspect_pdf(src)
    from pdf2docx import Converter

    part = part_path(dst)
    root = logging.getLogger()
    previous_level = root.level
    handler = _PageProgressHandler(progress) if progress else None
    cv = None
    try:
        if handler:
            root.addHandler(handler)
            root.setLevel(logging.INFO)
        cv = Converter(str(src))
        cv.convert(str(part))
        os.replace(part, dst)
    except ConversionError:
        part.unlink(missing_ok=True)
        raise
    except OSError as exc:
        part.unlink(missing_ok=True)
        raise ConversionError(MSG_CANNOT_SAVE) from exc
    except Exception as exc:
        part.unlink(missing_ok=True)
        raise ConversionError(MSG_FAILED) from exc
    finally:
        if cv is not None:
            cv.close()
        if handler:
            root.removeHandler(handler)
            root.setLevel(previous_level)
