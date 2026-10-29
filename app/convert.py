"""PDF <-> Word conversion helpers. Heavy libraries are imported lazily."""
from pathlib import Path

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
