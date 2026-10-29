# PDF Converter

## What it does

PDF Converter turns PDF files into Word documents and Word documents into PDF files. Drop your files onto the window and the converted copies appear next to the originals. It works completely offline: the app itself makes no network calls, and your files never leave your computer. The window is drawn by Microsoft Edge WebView2, which may do some background network traffic of its own; the app asks it to keep that to a minimum.

## Download and run

1. Go to the Releases page of this project on GitHub and download `PDF Converter.exe`.
2. Double-click it. There is nothing to install.

The first time, Windows may show a blue "Windows protected your PC" (SmartScreen) dialog. Click **More info**, then **Run anyway**. This appears because the program is not code-signed, so Windows does not recognise the publisher. It does not mean anything is wrong with the file.

## Requirements

- Windows 10 or 11.
- Microsoft Edge WebView2 runtime. It is preinstalled on Windows 11 and most Windows 10 machines. If it is missing, the app tells you and you can get it free from https://developer.microsoft.com/microsoft-edge/webview2/
- Microsoft Word or LibreOffice, only for Word to PDF.

## Known limits

- Scanned PDFs (pages that are only pictures) are not supported yet.
- Password-protected PDFs are refused.
- Password-protected or write-protected Word files are refused.
- Layout fidelity varies: complex documents may not look identical after conversion.

## Where files go

The converted file is saved next to the original. An existing file is never overwritten: a number is added instead, for example `Report (1).docx`.

## If something goes wrong

The app writes problems to a log file at `%LOCALAPPDATA%\PDF Converter\app.log`. Include it when you report a problem.

## Building from source

```
py -m venv .venv
.venv\Scripts\python -m pip install -r requirements-dev.txt
powershell -ExecutionPolicy Bypass -File build.ps1
```

The result is `dist\PDF Converter.exe`.

## Running tests

```
.venv\Scripts\python -m pytest
```

When Word is installed, the suite drives a hidden Word instance and takes about 40 seconds.

## License

AGPL-3.0, because PyMuPDF and pdf2docx are AGPL/GPL licensed. See [LICENSE](LICENSE).
