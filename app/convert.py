"""PDF <-> Word conversion helpers. Heavy libraries are imported lazily."""
import logging
import os
import re
import zipfile
from functools import lru_cache
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
    """Temporary name used while writing dst (per process, in dst's folder)."""
    return dst.with_name(f"{dst.name}.{os.getpid()}.part")


_RENAME_ATTEMPTS = 5


def finish_part(part: Path, dst: Path) -> Path:
    """Rename part to dst without ever overwriting; return the final path.

    os.rename refuses an existing target on Windows. If dst appeared while
    converting (e.g. another instance wrote it), the next free name is used.
    """
    for _ in range(_RENAME_ATTEMPTS):
        try:
            os.rename(part, dst)
            return dst
        except FileExistsError:
            dst = unique_path(dst)
    raise ConversionError(MSG_CANNOT_SAVE)


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


FONT_KEY = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"
FONT_TAG = re.compile(r"<w:rFonts\b[^>]*>")
FONT_ATTR = re.compile(r'(w:(?:ascii|hAnsi|eastAsia|cs))="([^"]*)"')

# Metric-compatible stand-ins, matched against the upper-cased name with
# spaces and dashes removed. The longest matching key wins.
FONT_TABLE = {
    "Times New Roman": (
        "NIMBUSROM", "TERMES", "TIMES", "LMROMAN", "CMR", "CMBX", "CMTI",
        "CMSL", "CMCSC", "SFRM", "SFBX", "SFTI",
    ),
    "Arial": ("NIMBUSSAN", "HEROS", "HELVETICA", "ARIAL", "LMSANS", "CMSS", "SFSS"),
    "Courier New": ("NIMBUSMON", "CURSOR", "COURIER", "LMMONO", "CMTT", "SFTT"),
    "Palatino Linotype": ("PALLADIO", "PALATINO", "PAGELLA"),
    "Bookman Old Style": ("BOOKMAN", "BONUM"),
    "Century Schoolbook": ("CENTURYSCH", "SCHOLA"),
    "Monotype Corsiva": ("CHANCERY", "CHORUS"),
    "Cambria Math": ("CMSY", "CMMI", "CMEX", "MSAM", "MSBM"),
}
CLASS_FALLBACK = {"mono": "Courier New", "sans": "Arial", "serif": "Times New Roman"}


@lru_cache(maxsize=1)
def installed_font_names() -> set[str]:
    """Lower-cased font names from the Windows registry, e.g. 'arial (truetype)'."""
    import winreg

    names: set[str] = set()
    for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
        try:
            with winreg.OpenKey(hive, FONT_KEY) as key:
                for index in range(winreg.QueryInfoKey(key)[1]):
                    names.add(winreg.EnumValue(key, index)[0].lower())
        except OSError:
            continue
    return names


def font_is_installed(name: str, installed: set[str]) -> bool:
    prefix = name.lower()
    return any(entry.startswith(prefix) for entry in installed)


def _pdf2docx_name(name: str) -> str:
    """The font name pdf2docx writes for a PDF font name."""
    return name.split("+")[-1].split("-")[0]


def pdf_font_classes(pdf: Path) -> dict[str, str]:
    """Map each normalised, lower-cased PDF font name to mono, serif or sans."""
    import pymupdf

    classes: dict[str, str] = {}
    with pymupdf.open(pdf) as doc:
        for page in doc:
            for block in page.get_text("dict")["blocks"]:
                for line in block.get("lines", ()):
                    for span in line["spans"]:
                        flags = span["flags"]
                        cls = "mono" if flags & 8 else "serif" if flags & 4 else "sans"
                        classes[_pdf2docx_name(span["font"]).lower()] = cls
    return classes


def substitute_for(name: str, cls: str | None, installed: set[str]) -> str | None:
    """Installed metric-compatible replacement for a missing font, else None."""
    if font_is_installed(name, installed):
        return None
    key = name.upper().replace(" ", "").replace("-", "")
    best, best_len = None, 0
    for substitute, keys in FONT_TABLE.items():
        for candidate in keys:
            if candidate in key and len(candidate) > best_len:
                best, best_len = substitute, len(candidate)
    if best is None:
        best = CLASS_FALLBACK.get(cls or "serif", "Times New Roman")
    return best if font_is_installed(best, installed) else None


def _word_parts(names: list[str]) -> list[str]:
    return [n for n in names if n.startswith("word/") and n.endswith(".xml")]


def docx_font_names(docx: Path) -> set[str]:
    found: set[str] = set()
    with zipfile.ZipFile(docx) as zf:
        for part in _word_parts(zf.namelist()):
            text = zf.read(part).decode("utf-8")
            for tag in FONT_TAG.findall(text):
                found.update(m.group(2) for m in FONT_ATTR.finditer(tag))
    return found


def substitute_fonts(docx: Path, mapping: dict[str, str]) -> None:
    """Rewrite font names in every word/*.xml part, replacing the file."""
    if not mapping:
        return

    def swap(match: re.Match) -> str:
        return f'{match.group(1)}="{mapping.get(match.group(2), match.group(2))}"'

    tmp = docx.with_name(f"{docx.name}.fonts")
    try:
        with zipfile.ZipFile(docx) as zin, zipfile.ZipFile(tmp, "w") as zout:
            parts = set(_word_parts(zin.namelist()))
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename in parts:
                    text = FONT_TAG.sub(
                        lambda tag: FONT_ATTR.sub(swap, tag.group(0)),
                        data.decode("utf-8"),
                    )
                    data = text.encode("utf-8")
                zout.writestr(item, data)
        os.replace(tmp, docx)
    finally:
        tmp.unlink(missing_ok=True)


def font_substitutions(pdf: Path, docx: Path) -> dict[str, str]:
    """Map DOCX font names Windows lacks to installed look-alikes."""
    installed = installed_font_names()
    classes = pdf_font_classes(pdf)
    mapping: dict[str, str] = {}
    for name in docx_font_names(docx):
        cls = classes.get(_pdf2docx_name(name).lower())
        substitute = substitute_for(name, cls, installed)
        if substitute:
            mapping[name] = substitute
    return mapping


def pdf_to_docx(
    src: Path,
    dst: Path,
    progress: Callable[[int, int], None] | None = None,
) -> Path:
    """Convert a PDF to DOCX, reporting progress(page, pages).

    Returns the path written: dst, or a numbered name if dst appeared meanwhile.
    """
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
        # A page that fails to parse would otherwise be dropped silently.
        cv.convert(str(part), ignore_page_error=False)
        try:
            mapping = font_substitutions(src, part)
            if mapping:
                logging.info("Font substitutions: %s", mapping)
                substitute_fonts(part, mapping)
        except Exception as exc:
            logging.warning("Font substitution skipped: %s", exc)
        return finish_part(part, dst)
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
