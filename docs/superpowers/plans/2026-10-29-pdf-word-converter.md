# PDF Converter Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** A single-`.exe` Windows app that converts PDF to Word and Word to PDF by drag and drop, entirely offline.

**Architecture:** A `pywebview` window hosts a static HTML/CSS/JS page (`app/ui/`). Python exposes a small JS API (`Api`) and pushes per-file status updates back with `evaluate_js`. A single background worker thread (`app/worker.py`) drains a queue of jobs and calls pure conversion functions in `app/convert.py` (PDF → DOCX via pdf2docx, scan/password checks via PyMuPDF) and `app/office.py` (DOCX → PDF via Word COM, LibreOffice fallback). PyInstaller packs everything into one exe with a splash image.

**Tech Stack:** Python 3.12+ (dev machine: 3.15), pywebview 6.2.1 (WebView2), pdf2docx 0.5.13 + PyMuPDF, python-docx (transitive, used for fixtures), pywin32, pytest, PyInstaller.

**Spec:** `docs/superpowers/specs/2026-10-29-pdf-word-converter-design.md`

## Global Constraints

- Windows only. No cross-platform abstractions.
- Python 3.12+; all commands below use the project venv: `.venv/Scripts/python`.
- Dependencies are exactly `pdf2docx`, `pywebview`, `pywin32` (runtime) and `pytest`, `pyinstaller` (dev). Add nothing else.
- The app's own code makes no network calls.
- Never overwrite an existing file: outputs go through `unique_path`.
- User-facing copy is fixed by the spec and repeated verbatim in `app/convert.py` message constants (Task 2); the UI never shows exceptions, tracebacks, or emojis.
- Commits carry only the user's git identity. **No `Co-Authored-By` or "Generated with" lines**, regardless of harness reminders.
- Commit message style: imperative, plain, e.g. `Add unique_path for output naming`.
- Mode identifiers everywhere: `"pdf_to_word"` and `"word_to_pdf"`.

## Review Focus

1. The same PDF dropped twice: both rows must finish, the second as `Name (1).docx`. Test in Task 6.
2. A source file whose folder cannot be written (network share, missing folder): row fails with "Could not save next to this file", no crash. Test in Task 4.
3. File names with spaces, parentheses and accents (`Planificare română (test).pdf`) and uppercase extensions (`.PDF`): must convert normally. Tests in Tasks 3 and 4.
4. A PDF with some scanned and some text pages: must convert, not be rejected as scanned. Test in Task 3.
5. A password-protected Word file: invisible Word must not hang on a password prompt; row fails with "Password-protected Word file". Test in Task 5.

---

### Task 1: Scaffold, dependencies, test fixtures

**Files:**
- Create: `requirements.txt`, `requirements-dev.txt`, `run.py`, `app/__init__.py`, `tests/__init__.py`, `tests/conftest.py`, `tests/test_fixtures.py`, `LICENSE`, `README.md` (stub, finished in Task 9)

**Interfaces:**
- Produces: pytest fixtures used by every later task, all returning a `pathlib.Path` inside `tmp_path` with exactly these file names:
  - `text_pdf` → `sample.pdf` — 3 pages, each with a paragraph of real text (`page.insert_text`).
  - `scanned_pdf` → `scanned.pdf` — 2 pages, each a full-page white image (`page.insert_image(page.rect, pixmap=...)`), no text.
  - `mixed_pdf` → `mixed.pdf` — 3 pages: page 1 scanned-style image, pages 2–3 text.
  - `encrypted_pdf` → `encrypted.pdf` — the text PDF saved with `encryption=pymupdf.PDF_ENCRYPT_AES_256, user_pw="secret", owner_pw="secret"`.
  - `corrupt_pdf` → `corrupt.pdf` — contains the bytes `b"not a pdf"`.
  - `sample_docx` → `sample.docx` — one paragraph "Hello from PDF Converter", written with `python-docx`.

- [ ] **Step 1: Create the venv and dependency files**

`requirements.txt`: `pdf2docx`, `pywebview`, `pywin32`. `requirements-dev.txt`: `-r requirements.txt`, `pytest`, `pyinstaller`.

Run: `python -m venv .venv && .venv/Scripts/python -m pip install -r requirements-dev.txt`
Expected: ends with `Successfully installed ...` and no error.

- [ ] **Step 2: Write `tests/conftest.py` with the six fixtures above, and `tests/test_fixtures.py`**

```python
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
```

- [ ] **Step 3: Run the fixture tests**

Run: `.venv/Scripts/python -m pytest tests/test_fixtures.py -v`
Expected: 3 passed.

- [ ] **Step 4: Add `run.py`, `app/__init__.py`, `LICENSE`, `README.md` stub**

`run.py` is the dev and PyInstaller entry point: `from app.main import main` inside `if __name__ == "__main__":` then `main()` (Task 8 creates `app/main.py`; until then `run.py` is not executed). `LICENSE` is the full AGPL-3.0 text (`gh api /licenses/agpl-3.0 --jq .body > LICENSE`). `README.md` stub: title "PDF Converter" and one sentence from the spec's Intent.

- [ ] **Step 5: Commit**

```bash
git add requirements.txt requirements-dev.txt run.py app/__init__.py tests LICENSE README.md
git commit -m "Scaffold project with dependencies and test fixtures"
```

---

### Task 2: Output naming and input validation

**Files:**
- Create: `app/convert.py`
- Test: `tests/test_convert_paths.py`

**Interfaces:**
- Produces, in `app/convert.py`:
  - `PDF_TO_WORD = "pdf_to_word"`, `WORD_TO_PDF = "word_to_pdf"`
  - Message constants (exact copy): `MSG_SCANNED = "Scanned PDF, not supported yet"`, `MSG_PASSWORD_PDF = "Password-protected PDF"`, `MSG_PASSWORD_WORD = "Password-protected Word file"`, `MSG_NO_OFFICE = "Word or LibreOffice must be installed for Word to PDF"`, `MSG_NOT_PDF = "Not a PDF file"`, `MSG_NOT_WORD = "Not a Word file"`, `MSG_CANNOT_SAVE = "Could not save next to this file"`, `MSG_FAILED = "Could not convert this file"`
  - `class ConversionError(Exception)` — `.message: str` is the user-facing text (one of the constants).
  - `unique_path(path: Path) -> Path` — returns `path` if it does not exist, else `"<stem> (1)<suffix>"`, `"<stem> (2)<suffix>"`, … the first that does not exist.
  - `output_path_for(source: Path, mode: str) -> Path` — validates the extension for the mode (case-insensitive; `pdf_to_word` accepts `.pdf`; `word_to_pdf` accepts `.docx` and `.doc`), raises `ConversionError(MSG_NOT_PDF | MSG_NOT_WORD)`, and returns `unique_path(source.with_suffix(".docx" | ".pdf"))`.
  - `part_path(dst: Path) -> Path` — `dst` with `.part` appended to the name (`Report.docx.part`).

- [ ] **Step 1: Write the failing tests**

```python
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
```

- [ ] **Step 2: Run to verify failure**

Run: `.venv/Scripts/python -m pytest tests/test_convert_paths.py -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'app.convert'`.

- [ ] **Step 3: Implement the constants, `ConversionError`, `unique_path`, `output_path_for`, `part_path` in `app/convert.py`**

Keep heavy imports (`pymupdf`, `pdf2docx`) out of module scope; this file is imported at startup.

- [ ] **Step 4: Run to verify pass**

Run: `.venv/Scripts/python -m pytest tests/test_convert_paths.py -v`
Expected: 6 passed.

- [ ] **Step 5: Commit**

```bash
git add app/convert.py tests/test_convert_paths.py
git commit -m "Add output naming and input validation"
```

---

### Task 3: PDF inspection (page count, scanned, password)

**Files:**
- Modify: `app/convert.py`
- Test: `tests/test_inspect_pdf.py`

**Interfaces:**
- Produces: `inspect_pdf(path: Path) -> int` — opens with `pymupdf` (imported inside the function) and returns the page count. Raises `ConversionError(MSG_PASSWORD_PDF)` if `doc.needs_pass`; `ConversionError(MSG_SCANNED)` if more than half of the pages are scanned; `ConversionError(MSG_FAILED)` if the file cannot be opened or has 0 pages.
  A page is scanned when `len(page.get_text().strip()) < 20` and some image (`page.get_images()` + `page.get_image_rects(xref)`) covers at least 80 % of `page.rect` area.

- [ ] **Step 1: Write the failing tests**

```python
def test_text_pdf_page_count(text_pdf):
    assert inspect_pdf(text_pdf) == 3

def test_scanned_pdf_raises(scanned_pdf):
    with pytest.raises(ConversionError) as e:
        inspect_pdf(scanned_pdf)
    assert e.value.message == MSG_SCANNED

def test_mixed_pdf_is_not_scanned(mixed_pdf):
    assert inspect_pdf(mixed_pdf) == 3

def test_encrypted_pdf_raises(encrypted_pdf):
    with pytest.raises(ConversionError) as e:
        inspect_pdf(encrypted_pdf)
    assert e.value.message == MSG_PASSWORD_PDF

def test_corrupt_pdf_raises_failed(corrupt_pdf):
    with pytest.raises(ConversionError) as e:
        inspect_pdf(corrupt_pdf)
    assert e.value.message == MSG_FAILED

def test_unicode_name_is_fine(text_pdf, tmp_path):
    src = tmp_path / "Planificare română (test).PDF"
    shutil.copy(text_pdf, src)
    assert inspect_pdf(src) == 3
```

- [ ] **Step 2: Run to verify failure**

Run: `.venv/Scripts/python -m pytest tests/test_inspect_pdf.py -v`
Expected: FAIL with `ImportError: cannot import name 'inspect_pdf'`.

- [ ] **Step 3: Implement `inspect_pdf` in `app/convert.py`**

- [ ] **Step 4: Run to verify pass**

Run: `.venv/Scripts/python -m pytest tests/test_inspect_pdf.py -v`
Expected: 6 passed.

- [ ] **Step 5: Commit**

```bash
git add app/convert.py tests/test_inspect_pdf.py
git commit -m "Detect scanned and password-protected PDFs"
```

---

### Task 4: PDF to DOCX conversion with page progress

**Files:**
- Modify: `app/convert.py`
- Test: `tests/test_pdf_to_docx.py`

**Interfaces:**
- Consumes: `inspect_pdf`, `part_path`, `ConversionError` (Tasks 2–3).
- Produces: `pdf_to_docx(src: Path, dst: Path, progress: Callable[[int, int], None] | None = None) -> None`.
  Calls `inspect_pdf(src)` first. Converts with `pdf2docx.Converter(str(src)).convert(str(part_path(dst)))` then `os.replace(part, dst)`; always `cv.close()`. `progress(page, pages)` is called from pdf2docx's per-page log records: attach a `logging.Handler` to the root logger for the duration of the call (root level set to `INFO` and restored after), and forward records whose `msg == "(%d/%d) Page %d"` and that come from the parsing pass (the first `(1/n)` … `(n/n)` run; the second "Creating pages" pass is ignored, so progress never goes backwards). `OSError` while writing → `ConversionError(MSG_CANNOT_SAVE)`; any other exception → `ConversionError(MSG_FAILED)`. The part file is removed on any failure.

- [ ] **Step 1: Write the failing tests**

```python
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
```

- [ ] **Step 2: Run to verify failure**

Run: `.venv/Scripts/python -m pytest tests/test_pdf_to_docx.py -v`
Expected: FAIL with `ImportError: cannot import name 'pdf_to_docx'`.

- [ ] **Step 3: Implement `pdf_to_docx` and the private progress log handler in `app/convert.py`**

- [ ] **Step 4: Run to verify pass**

Run: `.venv/Scripts/python -m pytest tests/test_pdf_to_docx.py -v`
Expected: 5 passed.

- [ ] **Step 5: Commit**

```bash
git add app/convert.py tests/test_pdf_to_docx.py
git commit -m "Convert PDF to DOCX with page progress"
```

---

### Task 5: Word to PDF via Word COM, LibreOffice fallback

**Files:**
- Create: `app/office.py`
- Test: `tests/test_office.py`

**Interfaces:**
- Consumes: `ConversionError`, `MSG_NO_OFFICE`, `MSG_PASSWORD_WORD`, `MSG_CANNOT_SAVE`, `MSG_FAILED` (Task 2).
- Produces, in `app/office.py`:
  - `word_installed() -> bool` — `winreg` lookup of `HKEY_CLASSES_ROOT\Word.Application\CurVer`.
  - `find_libreoffice() -> Path | None` — first existing of `C:\Program Files\LibreOffice\program\soffice.exe`, `C:\Program Files (x86)\LibreOffice\program\soffice.exe`, `shutil.which("soffice")`.
  - `office_available() -> bool` — `word_installed() or find_libreoffice() is not None`.
  - `docx_to_pdf(src: Path, dst: Path) -> None` — Word first, else LibreOffice, else `ConversionError(MSG_NO_OFFICE)`. Both engines write into a `tempfile.TemporaryDirectory()` and the result is `shutil.move`d to `dst` (`OSError` on the move → `MSG_CANNOT_SAVE`).
    - Word: `pythoncom.CoInitialize()`; `word = win32com.client.DispatchEx("Word.Application")`; `word.Visible = False`; `word.DisplayAlerts = 0`; `doc = word.Documents.Open(FileName=str(src), ReadOnly=True, AddToRecentFiles=False, PasswordDocument="\x01")` (a wrong password makes Word raise instead of prompting; an unprotected file ignores it); `doc.ExportAsFixedFormat(OutputFileName=<tmp>/<stem>.pdf, ExportFormat=17)`; `doc.Close(SaveChanges=0)`; `word.Quit()` in `finally`; `pythoncom.CoUninitialize()`. A `pywintypes.com_error` whose text contains "password" → `MSG_PASSWORD_WORD`, any other → `MSG_FAILED`.
    - LibreOffice: `subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", tmp, str(src)], timeout=120, creationflags=subprocess.CREATE_NO_WINDOW)`; missing output → `MSG_FAILED`.

- [ ] **Step 1: Write the failing tests**

```python
needs_word = pytest.mark.skipif(not word_installed(), reason="Word not installed")
needs_lo = pytest.mark.skipif(find_libreoffice() is None, reason="LibreOffice not installed")

def test_office_available_matches_parts():
    assert office_available() == (word_installed() or find_libreoffice() is not None)

@needs_word
def test_word_converts_docx(sample_docx, tmp_path):
    dst = tmp_path / "out.pdf"
    docx_to_pdf(sample_docx, dst)
    import pymupdf
    assert "Hello from PDF Converter" in pymupdf.open(dst)[0].get_text()

@needs_word
def test_password_protected_docx_reports_message(sample_docx, tmp_path):
    # create a protected copy with Word itself
    protected = tmp_path / "locked.docx"
    ...  # DispatchEx Word, Documents.Open(sample_docx), doc.Password = "secret", doc.SaveAs2(str(protected)), close, quit
    with pytest.raises(ConversionError) as e:
        docx_to_pdf(protected, tmp_path / "out.pdf")
    assert e.value.message == MSG_PASSWORD_WORD

@needs_lo
def test_libreoffice_converts_docx(sample_docx, tmp_path, monkeypatch):
    monkeypatch.setattr("app.office.word_installed", lambda: False)
    docx_to_pdf(sample_docx, tmp_path / "out.pdf")
    assert (tmp_path / "out.pdf").exists()

def test_no_office_raises(sample_docx, tmp_path, monkeypatch):
    monkeypatch.setattr("app.office.word_installed", lambda: False)
    monkeypatch.setattr("app.office.find_libreoffice", lambda: None)
    with pytest.raises(ConversionError) as e:
        docx_to_pdf(sample_docx, tmp_path / "out.pdf")
    assert e.value.message == MSG_NO_OFFICE
```

- [ ] **Step 2: Run to verify failure**

Run: `.venv/Scripts/python -m pytest tests/test_office.py -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'app.office'`.

- [ ] **Step 3: Implement `app/office.py`**

- [ ] **Step 4: Run to verify pass**

Run: `.venv/Scripts/python -m pytest tests/test_office.py -v`
Expected: 4 passed, 1 skipped (LibreOffice is not on the dev machine; Word 16 is). No Word window may appear during the run.

- [ ] **Step 5: Commit**

```bash
git add app/office.py tests/test_office.py
git commit -m "Convert Word to PDF with Word COM and LibreOffice fallback"
```

---

### Task 6: Job queue and worker thread

**Files:**
- Create: `app/worker.py`
- Test: `tests/test_worker.py`

**Interfaces:**
- Consumes: `output_path_for`, `part_path`, `pdf_to_docx`, `ConversionError`, mode constants (Tasks 2–4); `docx_to_pdf` (Task 5).
- Produces, in `app/worker.py`:
  - `@dataclass class Job`: `id: int`, `source: Path`, `mode: str`, `status: str = "waiting"` (one of `waiting | converting | done | failed`), `message: str = ""`, `page: int = 0`, `pages: int = 0`, `output: Path | None = None`; `to_dict() -> dict` with keys `id, name (source.name), status, message, page, pages`.
  - `class ConversionWorker`:
    - `__init__(self, on_update: Callable[[Job], None])` — `on_update` is called (from the worker thread) after every change to a job.
    - `add(self, paths: Iterable[str | Path], mode: str) -> list[Job]` — directories are skipped; each file becomes a `waiting` job (ids increase from 1), `on_update` fires once per new job, jobs are queued in order. The daemon thread starts lazily on first `add`.
    - `jobs: dict[int, Job]`, `last_output: Path | None` (most recent `done` job's output).
    - `stop() -> None` — sets a stop flag, clears the queue, and deletes `part_path(job.output)` of the job currently converting if it exists. Does not join.
    - Processing one job: `output_path_for` → status `converting` (update) → `pdf_to_docx(src, out, progress)` or `docx_to_pdf(src, out)` → status `done`, `output` set, `last_output` set. Progress callbacks set `page/pages` and fire `on_update`. `ConversionError` → `failed` with its `.message`; any other exception → `failed` with `MSG_FAILED` and the traceback logged via `logging.exception`.

- [ ] **Step 1: Write the failing tests**

```python
def wait_for(pred, timeout=60):
    end = time.time() + timeout
    while time.time() < end:
        if pred(): return
        time.sleep(0.05)
    raise AssertionError("timeout")

def test_same_pdf_twice_gets_numbered_output(text_pdf):
    updates = []
    w = ConversionWorker(on_update=lambda j: updates.append(j.to_dict()))
    jobs = w.add([text_pdf, text_pdf], PDF_TO_WORD)
    wait_for(lambda: all(j.status in ("done", "failed") for j in jobs))
    assert [j.status for j in jobs] == ["done", "done"]
    assert jobs[0].output.name == "sample.docx" and jobs[1].output.name == "sample (1).docx"
    assert w.last_output == jobs[1].output
    assert any(u["status"] == "converting" and u["pages"] == 3 for u in updates)

def test_wrong_type_fails_with_message(sample_docx):
    w = ConversionWorker(on_update=lambda j: None)
    [job] = w.add([sample_docx], PDF_TO_WORD)
    wait_for(lambda: job.status == "failed")
    assert job.message == MSG_NOT_PDF

def test_directories_are_ignored(tmp_path):
    w = ConversionWorker(on_update=lambda j: None)
    assert w.add([tmp_path], PDF_TO_WORD) == []

def test_stop_removes_partial_output(text_pdf, monkeypatch):
    started, release = threading.Event(), threading.Event()
    def fake(src, dst, progress=None):
        part_path(dst).touch(); started.set(); release.wait()
    monkeypatch.setattr("app.worker.pdf_to_docx", fake)
    w = ConversionWorker(on_update=lambda j: None)
    [job] = w.add([text_pdf], PDF_TO_WORD)
    started.wait(5)
    w.stop()
    assert not part_path(job.output).exists()
    release.set()
```

`app/worker.py` imports `pdf_to_docx` and `docx_to_pdf` by name into its own namespace (`from app.convert import ...`) so the test above can monkeypatch `app.worker.pdf_to_docx`.

- [ ] **Step 2: Run to verify failure**

Run: `.venv/Scripts/python -m pytest tests/test_worker.py -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'app.worker'`.

- [ ] **Step 3: Implement `app/worker.py`**

- [ ] **Step 4: Run to verify pass**

Run: `.venv/Scripts/python -m pytest tests/test_worker.py -v`
Expected: 4 passed.

- [ ] **Step 5: Commit**

```bash
git add app/worker.py tests/test_worker.py
git commit -m "Add conversion job queue and worker thread"
```

---

### Task 7: UI page

**Files:**
- Create: `app/ui/index.html`, `app/ui/style.css`, `app/ui/app.js`

**Interfaces:**
- Consumes (Task 8 provides these on `window.pywebview.api`, all returning Promises): `set_mode(mode)`, `choose_files()`, `open_folder()`, `show_file(id)`, `office_available()`.
- Produces: `window.app.onJob(job)` — `job` is `Job.to_dict()` from Task 6; creates or updates the row with that `id`. `window.app.mode` holds the current mode string. Element ids: `#drop` (drop area; Task 8 attaches the Python drop handler to it), `#choose` (link), `#mode-pdf` / `#mode-word` (segments), `#list`, `#open-folder`, `#banner`.

- [ ] **Step 1: Write `index.html`, `style.css`, `app.js` to the spec's "User interface" section**

Copy, verbatim from the spec: title `PDF Converter`; segments `PDF to Word` / `Word to PDF`; drop heading `Drop PDF files here` / `Drop Word files here`; sub-line `or choose files` / `.docx or .doc, or choose files` (with "choose files" as the link); statuses `Waiting`, `Page {page} of {pages}` (when `pages > 0`) else `Converting…`, `Done` + `Show` link, failure `message`; footer caption `Saved next to the original files`; button `Open folder` (disabled until a job is `done`); banner text `Word or LibreOffice must be installed for Word to PDF` shown only in `word_to_pdf` mode when `office_available()` resolves false. Style: Segoe UI, accent `#0071e3`, failure colour `#b25000`, radius 10–14 px, thin separators, `@media (prefers-color-scheme: dark)` variant, no emojis, no images other than one inline SVG outline document icon. JS: `dragover`/`drop` on `#drop` call `preventDefault()` and toggle a `.hover` class (the file paths themselves arrive in Python, Task 8); everything that calls the API waits for the `pywebviewready` event; when `window.pywebview` is absent (plain browser) the page still renders.

- [ ] **Step 2: Verify in a browser**

Run: `start msedge "D:\PDF\app\ui\index.html"`, then in DevTools console:
`app.onJob({id:1,name:"Lesson plan.pdf",status:"done",message:"",page:0,pages:0}); app.onJob({id:2,name:"Report Q3.pdf",status:"converting",page:4,pages:12}); app.onJob({id:3,name:"Worksheet scan.pdf",status:"failed",message:"Scanned PDF, not supported yet"})`
Expected: three rows matching the spec mockup (Done + Show, `Page 4 of 12` with a blue bar, orange failure text), `Open folder` enabled, switching to Word to PDF changes the drop copy, dark mode follows the OS setting (toggle Windows dark mode to check).

- [ ] **Step 3: Commit**

```bash
git add app/ui
git commit -m "Add converter window page"
```

---

### Task 8: Window, JS API, drop handling, startup checks

**Files:**
- Create: `app/startup.py`, `app/main.py`
- Test: `tests/test_startup.py`

**Interfaces:**
- Consumes: `ConversionWorker`, `Job` (Task 6); `office_available` (Task 5); UI ids and `app.onJob` (Task 7).
- Produces, in `app/startup.py`:
  - `WEBVIEW2_CLIENT_KEY = r"Software\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}"`
  - `webview2_version() -> str | None` — reads value `pv` from `HKLM\SOFTWARE\WOW6432Node\...`, `HKLM\SOFTWARE\...`, `HKCU\...` (in that order); returns `None` if missing or equal to `"0.0.0.0"`.
  - `MISSING_WEBVIEW2_TEXT = "This app needs Microsoft Edge WebView2, which is free from Microsoft:\nhttps://developer.microsoft.com/microsoft-edge/webview2/"`
  - `show_missing_webview2() -> None` — `ctypes.windll.user32.MessageBoxW(0, MISSING_WEBVIEW2_TEXT, "PDF Converter", 0x10)`.
  - `BROWSER_ARGS = "--disable-features=msSmartScreenProtection --disable-background-networking"`
  - `configure_environment() -> None` — `os.environ.setdefault("WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS", BROWSER_ARGS)`.
  - `log_path() -> Path` — `%LOCALAPPDATA%\PDF Converter\app.log` (parent created). `main()` configures `logging.basicConfig(filename=log_path(), level=logging.WARNING)` and also sets the file handler's own level to `WARNING` (Task 4 temporarily raises the root level to `INFO` during a conversion; pdf2docx's page logs must not flood the file).
- Produces, in `app/main.py`:
  - `class Api` with `mode: str` (initially `PDF_TO_WORD`), `set_mode(mode: str) -> None`, `choose_files() -> None` (opens `window.create_file_dialog(webview.FileDialog.OPEN, allow_multiple=True, file_types=("PDF files (*.pdf)",) | ("Word files (*.docx;*.doc)",))` and passes the result to `worker.add`), `open_folder() -> None` (`os.startfile(worker.last_output.parent)`), `show_file(job_id: int) -> None` (`subprocess.Popen(["explorer", "/select,", str(job.output)])`), `office_available() -> bool`.
  - `main() -> None`: `configure_environment()`; if `webview2_version()` is `None` → `show_missing_webview2()` and return; create `ConversionWorker(on_update=push)` where `push` does `window.evaluate_js(f"window.app.onJob({json.dumps(job.to_dict())})")`; `webview.create_window("PDF Converter", url=str(Path(__file__).parent / "ui" / "index.html"), js_api=api, width=520, height=600, min_size=(420, 480))`; on `window.events.loaded`: `window.dom.get_element("#drop").on("drop", DOMEventHandler(on_drop, prevent_default=True))` where `on_drop(event)` collects `f["pywebviewFullPath"]` for each file in `event["dataTransfer"]["files"]` that has it and calls `worker.add(paths, api.mode)`; on `window.events.shown`: `import pyi_splash; pyi_splash.close()` inside `try/except Exception`; on `window.events.closing`: `worker.stop()`; `webview.start()`.

- [ ] **Step 1: Write the failing tests**

```python
def test_webview2_present_on_dev_machine():
    v = webview2_version()
    assert v and v != "0.0.0.0"

def test_configure_environment_sets_browser_args(monkeypatch):
    monkeypatch.delenv("WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS", raising=False)
    configure_environment()
    assert os.environ["WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS"] == BROWSER_ARGS

def test_log_path_is_under_localappdata():
    assert log_path().parts[-2:] == ("PDF Converter", "app.log")
    assert log_path().parent.exists()
```

- [ ] **Step 2: Run to verify failure**

Run: `.venv/Scripts/python -m pytest tests/test_startup.py -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'app.startup'`.

- [ ] **Step 3: Implement `app/startup.py`**

- [ ] **Step 4: Run to verify pass**

Run: `.venv/Scripts/python -m pytest tests/test_startup.py -v`
Expected: 3 passed.

- [ ] **Step 5: Implement `app/main.py`**

- [ ] **Step 6: Verify the running app**

Run: `.venv/Scripts/python run.py`
Expected, checked by hand: window titled "PDF Converter" appears; dragging `tests` fixture PDFs (save one with `.venv/Scripts/python -c "..."` or use any PDF on disk) onto the drop area adds rows that go `Waiting` → `Page x of y` → `Done`; `Show` selects the `.docx` in Explorer; `Open folder` opens its folder; a `.docx` dropped in PDF to Word mode shows `Not a PDF file`; switching to Word to PDF and dropping a `.docx` produces a `.pdf` with no Word window flashing; `choose files` opens a filtered dialog; closing the window exits the process.

- [ ] **Step 7: Run the whole suite and commit**

Run: `.venv/Scripts/python -m pytest -v`
Expected: all passed, 1 skipped.

```bash
git add app/startup.py app/main.py tests/test_startup.py
git commit -m "Add window, JS API and startup checks"
```

---

### Task 9: Assets, build script, README

**Files:**
- Create: `tools/make_assets.py`, `assets/icon.ico`, `assets/splash.png`, `build.ps1`
- Modify: `README.md`

**Interfaces:**
- Consumes: `run.py` entry (Task 1), `app/ui` data files (Task 7), `pyi_splash.close()` call (Task 8).
- Produces: `dist/PDF Converter.exe`.

- [ ] **Step 1: Write `tools/make_assets.py` and run it**

With PyMuPDF only (no Pillow): draw a 256×256 page with a rounded document outline in `#0071e3` on white (`page.draw_rect`, `page.draw_line`), render with `page.get_pixmap()`, save as PNG; wrap that PNG in an ICO container (6-byte header, one 16-byte directory entry, PNG payload — Windows accepts PNG-in-ICO) → `assets/icon.ico`. Splash: 420×240 white page with the text `PDF Converter` (Helvetica 28 pt, `#1d1d1f`) and `Starting…` (14 pt, `#86868b`) → `assets/splash.png`.

Run: `.venv/Scripts/python tools/make_assets.py && ls assets`
Expected: `icon.ico  splash.png`; both open in Windows Photos / show as the file icon in Explorer.

- [ ] **Step 2: Write `build.ps1`**

```powershell
$ErrorActionPreference = "Stop"
.\.venv\Scripts\python -m pip install -r requirements-dev.txt
.\.venv\Scripts\python tools\make_assets.py
.\.venv\Scripts\python -m PyInstaller --noconfirm --clean --onefile --windowed `
  --name "PDF Converter" --icon assets\icon.ico --splash assets\splash.png `
  --add-data "app\ui;app\ui" --collect-all webview run.py
```

- [ ] **Step 3: Build and run the exe**

Run: `powershell -ExecutionPolicy Bypass -File build.ps1`
Expected: ends with `Building EXE from EXE-00.toc completed successfully.` and `dist/PDF Converter.exe` exists (60–100 MB).

Run: `start "" "D:\PDF\dist\PDF Converter.exe"`
Expected: splash image appears within a second, then the window; the splash closes when the window shows; repeat Task 8 Step 6's drop checks on the exe; no console window; `%LOCALAPPDATA%\PDF Converter\app.log` exists. If the exe fails with a missing module or file, add the matching `--hidden-import` / `--collect-all` to `build.ps1` and rebuild.

- [ ] **Step 4: Write `README.md`**

Sections, plain English for non-technical readers first: what it does and that files never leave the computer (wording from the spec's Network note, including the WebView2 sentence); Download and run (release page; the SmartScreen "Windows protected your PC" dialog → "More info" → "Run anyway", with why it appears); Requirements (Windows 10/11, WebView2 — preinstalled on Windows 11; Word or LibreOffice only for Word to PDF); Known limits (scanned PDFs not yet, layout fidelity varies); Building from source (`build.ps1`); Running tests (`.venv/Scripts/python -m pytest`); License (AGPL-3.0, why: PyMuPDF and pdf2docx).

- [ ] **Step 5: Commit**

```bash
git add tools/make_assets.py assets build.ps1 README.md
git commit -m "Add build script, app assets and README"
```
