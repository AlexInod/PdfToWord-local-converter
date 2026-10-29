"""PDF <-> Word conversion helpers. Heavy libraries are imported lazily."""
import functools
import logging
import os
import re
import unicodedata
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
GLYPH_FONTS = ("DINGBAT", "SYMBOL", "WINGDING", "WEBDING")
CLASS_FALLBACK = {"mono": "Courier New", "sans": "Arial", "serif": "Times New Roman"}


@lru_cache(maxsize=1)
def installed_font_names() -> frozenset[str]:
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
    return frozenset(names)


def _registry_families(entry: str) -> set[str]:
    """Family names in a registry entry: 'arial bold (truetype)' -> {'arial'}."""
    entry = re.sub(r"\s*\([^)]*\)\s*$", "", entry)
    families = set()
    for part in entry.split(" & "):
        part = part.strip()
        families.add(part)
        while True:
            stripped = re.sub(r"\s+(bold|italic|regular)$", "", part)
            if stripped == part:
                break
            part = stripped
            families.add(part)
    return families


def font_is_installed(name: str, installed: set[str]) -> bool:
    """True when name equals an installed family (style words ignored)."""
    wanted = name.lower()
    return any(wanted in _registry_families(entry) for entry in installed)


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
    if any(word in key for word in GLYPH_FONTS):
        return None
    key = key.removeprefix("TEXGYRE")
    best, best_len = None, 0
    for substitute, keys in FONT_TABLE.items():
        for candidate in keys:
            if key.startswith(candidate) and len(candidate) > best_len:
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


def _is_blank_span(span: dict) -> bool:
    return (
        "image" not in span
        and not span.get("style")
        and not "".join(c["c"] for c in span["chars"]).strip()
    )


def _merge_blank_spans(raws: list[dict]) -> list[dict]:
    """Fold space-only spans into the preceding text span.

    pdf2docx drops unstyled blank spans, which glues words together when the
    PDF puts each inter-word space in its own span.
    """
    merged: list[dict] = []
    previous = None
    for span in raws:
        if previous is not None and _is_blank_span(span):
            a, b = previous["bbox"], span["bbox"]
            previous = dict(
                previous,
                chars=[*previous["chars"], *span["chars"]],
                bbox=(min(a[0], b[0]), min(a[1], b[1]), max(a[2], b[2]), max(a[3], b[3])),
            )
            merged[-1] = previous
            continue
        merged.append(span)
        previous = span if "image" not in span and not _is_blank_span(span) else None
    return merged


def _patch_pdf2docx_blank_spans() -> None:
    """Make pdf2docx keep the spaces that _merge_blank_spans folds in (once)."""
    # Depends on pdf2docx 0.5.13 internals (Spans.restore): re-verify on a pin bump.
    from pdf2docx.text.Spans import Spans

    original = Spans.restore
    if getattr(original, "_merges_blank_spans", False):
        return

    @functools.wraps(original)
    def restore(self, raws):
        try:
            merged = _merge_blank_spans(raws)
        except Exception:
            logging.warning("blank-span merge skipped for one line", exc_info=True)
            merged = raws
        return original(self, merged)

    restore._merges_blank_spans = True
    Spans.restore = restore


LINE_END_HYPHENS = ("-", "\u00ad")


def _is_word_edge(c: str) -> bool:
    """Letter, digit, symbol or non-dash punctuation in a script that spaces words."""
    if len(c) != 1:
        return False
    if c.isalnum():
        return unicodedata.east_asian_width(c) not in ("W", "F")
    category = unicodedata.category(c)
    return category.startswith(("P", "S")) and category != "Pd"


def _needs_line_end_space(end: str, start: str) -> bool:
    """True when the last char of a line and the first of the next are two words."""
    return _is_word_edge(end) and _is_word_edge(start)


def _adjust_last_word(lines, delete_end_line_hyphen: bool) -> None:
    """pdf2docx's Lines.adjust_last_word without its ASCII-only word test.

    The original tests c.encode().isalnum(), so a line ending next to „ ” or
    a letter such as ă glued onto the following line.
    """
    from pdf2docx.text.TextSpan import TextSpan

    instances = lines._instances
    for line, next_line in zip(instances, instances[1:]):
        end_span, start_span = line.spans[-1], next_line.spans[0]
        if not isinstance(end_span, TextSpan) or not isinstance(start_span, TextSpan):
            continue
        if not end_span.chars or not start_span.chars:
            continue
        end_char, start = end_span.chars[-1], start_span.chars[0].c[:1]
        if (
            delete_end_line_hyphen
            and end_char.c.endswith(LINE_END_HYPHENS)
            and start.islower()
        ):
            end_char.c = end_char.c[:-1]
        if _needs_line_end_space(end_char.c[-1:], start):
            end_char.c += " "


def _patch_pdf2docx_line_ends() -> None:
    """Install _adjust_last_word in pdf2docx (once)."""
    # Depends on pdf2docx 0.5.13 internals (Lines.adjust_last_word): re-verify on a pin bump.
    from pdf2docx.text.Lines import Lines

    original = Lines.adjust_last_word
    if getattr(original, "_unicode_line_ends", False):
        return

    @functools.wraps(original)
    def adjust_last_word(self, delete_end_line_hyphen):
        try:
            _adjust_last_word(self, delete_end_line_hyphen)
        except Exception:
            logging.warning("line-end adjustment fell back for one block", exc_info=True)
            original(self, delete_end_line_hyphen)

    adjust_last_word._unicode_line_ends = True
    Lines.adjust_last_word = adjust_last_word


def _drop_wrapped_line_tabs(lines) -> None:
    """Remove the TAB pdf2docx puts before a line that continues the previous one.

    The TAB only reproduces the PDF indent of a wrapped line (e.g. a hanging
    bibliography entry). Word wraps at different places, so it lands mid-line.
    """
    instances = lines._instances
    for previous, line in zip(instances, instances[1:]):
        if line.tab_stop and not previous.line_break and not previous.in_same_row(line):
            line.tab_stop = 0


def _patch_pdf2docx_wrapped_tabs() -> None:
    """Run _drop_wrapped_line_tabs after pdf2docx decides line breaks (once)."""
    # Depends on pdf2docx 0.5.13 internals (Lines.parse_line_break runs after
    # Lines.parse_tab_stop): re-verify on a pin bump.
    from pdf2docx.text.Lines import Lines

    original = Lines.parse_line_break
    if getattr(original, "_drops_wrapped_tabs", False):
        return

    @functools.wraps(original)
    def parse_line_break(self, *args, **kwargs):
        result = original(self, *args, **kwargs)
        try:
            _drop_wrapped_line_tabs(self)
        except Exception:
            logging.warning("wrapped-line tabs kept for one block", exc_info=True)
        return result

    parse_line_break._drops_wrapped_tabs = True
    Lines.parse_line_break = parse_line_break


# "Title . . . . . 12": a title, at least four leader dots, a page number.
TOC_LINE = re.compile(
    r"^(?P<head>.*?\S)\s*(?:\.\s*){4,}(?P<page>\d{1,4}|[ivxlcdm]{1,7})\s*$", re.IGNORECASE
)
TOC_LEADER = re.compile(r"(?:\.\s*){4,}")
# "A title too long for dots<TAB>12": only trusted next to a TOC_LINE.
TOC_TAIL = re.compile(r"^(?P<head>.*?\S)\s*\t\s*(?P<page>\d{1,4}|[ivxlcdm]{1,7})\s*$", re.IGNORECASE)


def _section_widths(body):
    """Yield (paragraph element, text width of its section) for body paragraphs."""
    from docx.oxml.ns import qn
    from docx.section import Section

    pending = []
    for child in body.iterchildren():
        if child.tag == qn("w:p"):
            pending.append(child)
            sect = child.find(qn("w:pPr") + "/" + qn("w:sectPr"))
        elif child.tag == qn("w:sectPr"):
            sect = child
        else:
            continue
        if sect is not None:
            section = Section(sect, None)
            width = section.page_width - section.left_margin - section.right_margin
            yield from ((p, width) for p in pending)
            pending = []


def _fix_toc_paragraph(p, width, pattern) -> bool:
    """Replace typed leader dots with one right-aligned dot-leader TAB."""
    from copy import deepcopy

    from docx.enum.text import WD_TAB_ALIGNMENT, WD_TAB_LEADER
    from docx.text.run import Run

    match = pattern.match(p.text)
    if not match or "\n" in p.text or len(TOC_LEADER.findall(p.text)) > 1:
        return False  # one entry per paragraph only; merged entries are left alone
    head, page = match["head"], match["page"]
    runs = [Run(r, p) for r in p._p.xpath("./w:r | ./w:hyperlink/w:r")]
    if not runs:
        return False
    source, done = runs[-1], 0
    for run in runs:
        text, start = run.text, done
        done += len(text)
        if start >= len(head):
            if page in text:
                source = run
            run._r.getparent().remove(run._r)
        elif done > len(head):
            run.text = text[: len(head) - start]
    for link in p._p.xpath("./w:hyperlink[not(w:r)]"):
        p._p.remove(link)
    for text in ("\t", page):
        r = deepcopy(source._r)
        p._p.append(r)
        Run(r, p).text = text

    stops = p.paragraph_format.tab_stops
    kept = sorted(stops, key=lambda s: s.position)[: head.count("\t")]
    kept = [(s.position, s.alignment, s.leader) for s in kept]
    stops.clear_all()
    for position, alignment, leader in kept:
        stops.add_tab_stop(position, alignment, leader)
    edge = width - (p.paragraph_format.right_indent or 0)
    stops.add_tab_stop(edge, WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
    return True


def fix_toc_leaders(docx: Path) -> int:
    """Line up table-of-contents page numbers; return how many lines changed."""
    from docx import Document
    from docx.text.paragraph import Paragraph

    document = Document(str(docx))
    lines = [
        (Paragraph(p, document._body), width)
        for p, width in _section_widths(document.element.body)
    ]
    fixed = {i for i, (p, width) in enumerate(lines) if _fix_toc_paragraph(p, width, TOC_LINE)}
    grown = bool(fixed)
    while grown:
        grown = False
        for i, (p, width) in enumerate(lines):
            if i in fixed or not ({i - 1, i + 1} & fixed):
                continue
            if _fix_toc_paragraph(p, width, TOC_TAIL):
                fixed.add(i)
                grown = True
    changed = len(fixed)
    if changed:
        tmp = docx.with_name(f"{docx.name}.toc")
        try:
            document.save(str(tmp))
            os.replace(tmp, docx)
        finally:
            tmp.unlink(missing_ok=True)
    return changed


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
    try:
        _patch_pdf2docx_blank_spans()
    except Exception as exc:
        logging.warning("Blank-span patch skipped: %s", exc)
    try:
        _patch_pdf2docx_line_ends()
    except Exception as exc:
        logging.warning("Line-end patch skipped: %s", exc)
    try:
        _patch_pdf2docx_wrapped_tabs()
    except Exception as exc:
        logging.warning("Wrapped-line tab patch skipped: %s", exc)

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
        # Line-end hyphens are mostly word breaks (LaTeX hyphenates by default).
        cv.convert(str(part), ignore_page_error=False, delete_end_line_hyphen=True)
        try:
            mapping = font_substitutions(src, part)
            if mapping:
                logging.info("Font substitutions: %s", mapping)
                substitute_fonts(part, mapping)
        except Exception as exc:
            logging.warning("Font substitution skipped: %s", exc)
        try:
            fix_toc_leaders(part)
        except Exception as exc:
            logging.warning("Table-of-contents fix skipped: %s", exc)
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
