# PDF Converter – design

Date: 2026-10-29
Status: approved in conversation; written form for the builder

## Intent

A free Windows desktop app that converts PDF to Word and Word to PDF entirely on the
user's own computer. Nothing is ever uploaded. It exists because everyday users
(office workers, teachers) currently paste private documents into online converter
sites to do this.

The code lives at https://github.com/AlexInod/PdfToWord-local-converter (private for
now, will likely be published).

## Goals

- One `.exe`, nothing to install, no Python knowledge needed.
- Drag and drop PDF files in, get `.docx` files out, saved next to the originals.
- Same for Word to PDF, via a switch at the top of the window.
- Plain-English status per file. Nobody should ever think the app is broken or wonder
  where the output went.
- Fully offline. The app makes no network requests at all.

## Non-goals (version 1)

- OCR for scanned PDFs (planned for a later version; v1 only detects and explains).
- macOS or Linux.
- Code signing. The `.exe` will trigger Windows SmartScreen's "unknown publisher"
  warning; the README explains how to get past it.
- Settings screen, output-folder picker, page-range selection, cancel button.
- Localisation. English only.
- Auto-update.

## Anti-goals

- Any network call, telemetry, or "check for updates" in the app's own code. If the
  user's documents leave the machine, it failed. (See the WebView2 note under
  Technology.)
- Overwriting a user's existing file. Ever.
- A visible Word window flashing up while converting Word to PDF.
- A UI that looks like a developer tool: no terminal, no logs on screen, no emojis.

## Audience

Non-technical office workers and teachers on Windows 10/11. They will download the
`.exe` from the GitHub release page.

## User interface

Mockup chosen: "option A" — one small window with a drop area and a file list.
Reference mockup: `.superpowers/brainstorm/1720-1793268823/content/layout-v2.html`
(not committed; the description below is authoritative).

Style: minimal, Apple Human Interface Guidelines inspired. Lots of white space,
gray secondary text, one accent blue (`#0071e3`) for links and progress, soft rounded
corners (10–14 px), almost no borders. Font: Segoe UI (Windows system font; SF Pro
cannot be shipped). No emojis anywhere. Light and dark mode follow the Windows
setting (`prefers-color-scheme`).

Window title: "PDF Converter". Default size about 520×600, resizable with a sensible
minimum.

Layout, top to bottom:

1. **Segmented control** with two segments: `PDF to Word` | `Word to PDF`. Switching
   modes clears nothing; files already in the list keep their status. Dropped files
   are interpreted in the mode active at drop time.
2. **Drop area**: dashed rounded rectangle, outline document icon, heading
   "Drop PDF files here" (or "Drop Word files here"), sub-line "or *choose files*"
   (link opens a native file dialog filtered to the mode's extensions; Word mode says
   ".docx or .doc, or choose files"). The area highlights while a drag hovers over it.
3. **File list**: one row per file, newest at the bottom, thin separators.
   Left: file name. Right: status text in gray, plus:
   - Queued: "Waiting"
   - Converting: "Page 4 of 12" with a thin blue progress bar when page progress is
     available, otherwise "Converting…" with an indeterminate bar.
   - Done: "Done" + a blue "Show" link that opens Explorer with the output selected.
   - Failed: a plain-English reason in a warm orange/brown (`#b25000`-ish), e.g.
     "Scanned PDF, not supported yet", "Password-protected PDF",
     "Word is not installed", "Not a PDF file", "Could not convert this file".
4. **Footer**: left, small gray caption "Saved next to the original files"; right,
   an "Open folder" button that opens the folder of the most recently converted file.
   Disabled (dimmed) until something has been converted.

Dropping a file whose extension does not match the active mode adds a failed row
("Not a PDF file" / "Not a Word file") rather than silently ignoring it.
Dropping a folder is ignored in v1.

## Behaviour

- Output file goes next to the input: `Report.pdf` → `Report.docx`. If
  `Report.docx` already exists, use `Report (1).docx`, then `Report (2).docx`, etc.
  Never overwrite.
- Files are converted one at a time in a background worker; the window stays
  responsive. Multiple drops append to the queue.
- Closing the window while converting stops the app; a partially written output file
  is deleted.
- Before converting a PDF, detect scanned documents: if the document's pages contain
  (almost) no extractable text and are dominated by one large image each, mark the
  row "Scanned PDF, not supported yet" and skip it. Encrypted PDFs that need a
  password → "Password-protected PDF".
- Word to PDF: use the Microsoft Word installed on the machine, invisibly, via COM
  automation (`Word.Application`, `Visible = False`, `ExportAsFixedFormat` to PDF).
  If Word is not installed, fall back to LibreOffice (`soffice --headless
  --convert-to pdf`) if found in the usual install locations. If neither exists, the
  row fails with "Word or LibreOffice must be installed for Word to PDF" and the
  message appears once as a banner under the segmented control while in that mode.
- Startup must feel instant; heavy imports (PyMuPDF, pdf2docx) load lazily or in a
  background thread so the window appears first.

## Technology (decided)

- Python 3.12+. Verified 2026-10-29: `pdf2docx`, `pymupdf` 1.28.2, `pywebview`
  6.2.1 and `pywin32` 312 all install on the dev machine's Python 3.15.
- UI: `pywebview` window hosting local HTML/CSS/JS (`app/ui/`). On Windows this
  uses the WebView2 runtime, which ships with Windows 11 and most Windows 10 machines.
  Dropped files expose their full path through pywebview's DOM drag-and-drop
  support: a Python handler on the drop element's `events.drop` receives
  `event['dataTransfer']['files'][i]['pywebviewFullPath']` (verified in pywebview
  6.2.1, `webview/util.py`). The "choose files" link uses
  `window.create_file_dialog`.
- PDF to Word engine: `pdf2docx` (which uses PyMuPDF and python-docx).
  Page progress: hook pdf2docx's per-page log messages if practical, otherwise
  indeterminate.
- Scan detection: PyMuPDF text and image extraction (already a dependency).
- Word to PDF: `pywin32` COM; LibreOffice via `subprocess`.
- Packaging: PyInstaller, `--onefile --windowed`, custom icon, build script
  `build.ps1`. Expect 60–100 MB. A one-file exe unpacks itself to `%TEMP%` on every
  launch (roughly 1–4 s before any Python runs), so the build includes a PyInstaller
  splash screen (`--splash`, a small static "PDF Converter" image) that shows
  immediately and closes when the window appears. "One `.exe`" wins over a faster
  `--onedir` zip, which non-technical users would have to extract.
- WebView2 runtime check: at startup, before creating the window, check the registry
  for the WebView2 runtime (`HKLM\SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\
  {F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}` or the same key under `HKCU`). If absent,
  show a native message box ("This app needs Microsoft Edge WebView2, which is free
  from Microsoft") with the download URL, then exit. Never fall back to pywebview's
  MSHTML (IE) renderer; the UI would look broken.
- Network: the *app code* makes no network calls, and the user's files never leave
  the machine. The embedded WebView2 engine (Edge) may do its own background
  traffic; the app sets `WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS` with
  `--disable-features=msSmartScreenProtection --disable-background-networking` to
  minimise it. The README states exactly this, honestly.
- CI: none in v1. The exe is built locally with `build.ps1` and uploaded to a GitHub
  release by hand. A GitHub Actions build is a planned follow-up.
- Tests: `pytest` for the conversion core (output naming, scan detection on generated
  fixture PDFs, mode/extension validation, Word/LibreOffice discovery). COM tests are
  skipped when Word is absent. UI is checked manually.
- License: AGPL-3.0 (required because PyMuPDF and pdf2docx are AGPL/GPL).
- Repo layout:

  ```
  app/
    main.py        # pywebview window + JS-facing API
    convert.py     # engine: pdf_to_docx, docx_to_pdf, detect_scanned, unique_path
    office.py      # Word COM + LibreOffice discovery/fallback
    ui/index.html, ui/style.css, ui/app.js
  tests/
  build.ps1
  requirements.txt
  README.md, LICENSE
  ```

## Constraints

- No network access of any kind in the shipped app.
- Commits carry only the user's git identity (`AlexInod`); no AI attribution lines.
- Windows only; do not add cross-platform abstractions for v1.
- Keep the dependency list short; every extra package grows the `.exe`.

## Left to the builder

- Exact CSS values, icon design, animation timings.
- How the pywebview JS API is shaped (e.g. `api.add_files(paths)` with events pushed
  via `window.evaluate_js`).
- Scan-detection thresholds (suggested: average < 20 text characters per page and an
  image covering > 80 % of the page area on most pages).
- Whether page progress is practical with pdf2docx; indeterminate is acceptable.
- PyInstaller hidden-import and data-file details.
