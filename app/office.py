"""Word -> PDF via Microsoft Word (COM), falling back to LibreOffice."""
import logging
import re
import shutil
import subprocess
import tempfile
import threading
import time
import uuid
from pathlib import Path

from app.convert import (
    ConversionError,
    MSG_CANNOT_SAVE,
    MSG_FAILED,
    MSG_NO_OFFICE,
    MSG_PASSWORD_WORD,
    finish_part,
    part_path,
)

# A hidden modal dialog in Word (first run, activation, encoding, a protection
# the pre-check misses) would block the worker forever; past this many seconds
# the Word process is terminated and the file fails.
WORD_TIMEOUT_SECONDS = 180

_LIBREOFFICE_PATHS = (
    Path(r"C:\Program Files\LibreOffice\program\soffice.exe"),
    Path(r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"),
)
_WD_EXPORT_FORMAT_PDF = 17
_OLE_SIGNATURE = bytes.fromhex("d0cf11e0a1b11ae1")
_FIB_PASSWORD_BITS = 0x01 | 0x08  # fEncrypted | fWriteReservation
_MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3
_STGM_READ = 0x00
_STGM_SHARE_DENY_WRITE = 0x20
_STGM_SHARE_EXCLUSIVE = 0x10  # streams must be opened exclusively
_WORD_WINDOW_CLASS = "OpusApp"
_FIND_WORD_WINDOW_SECONDS = 5.0
_PROCESS_TERMINATE = 0x0001


def word_installed() -> bool:
    import winreg

    try:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Word.Application\CurVer"):
            return True
    except OSError:
        return False


def find_libreoffice() -> Path | None:
    for path in _LIBREOFFICE_PATHS:
        if path.exists():
            return path
    found = shutil.which("soffice")
    return Path(found) if found else None


def office_available() -> bool:
    return word_installed() or find_libreoffice() is not None


_WRITE_PROTECTION = re.compile(
    rb"<w:writeProtection\s[^>]*(?:w:hash|w:cryptProviderType|w:algorithmName)="
)


def _has_write_password(src: Path) -> bool:
    """True for a .docx with a password to modify (Word prompts for it)."""
    import zipfile

    try:
        with zipfile.ZipFile(src) as package:
            settings = package.read("word/settings.xml")
    except (OSError, KeyError, zipfile.BadZipFile):
        return False
    return _WRITE_PROTECTION.search(settings) is not None


def _is_encrypted(src: Path) -> bool:
    """True for a Word file that needs a password, judged without opening Word.

    Word shows a modal password prompt even when hidden with alerts off, which
    would hang the worker, so protection is detected from the file itself.
    Encrypted .docx files are OLE containers holding an EncryptedPackage
    stream; legacy .doc files set fEncrypted or fWriteReservation in the FIB.
    If the file cannot be inspected reliably, raise ConversionError(MSG_FAILED)
    rather than let Word prompt.
    """
    import pythoncom

    try:
        with open(src, "rb") as handle:
            signature = handle.read(len(_OLE_SIGNATURE))
    except OSError as exc:
        raise ConversionError(MSG_FAILED) from exc
    if signature != _OLE_SIGNATURE:
        return _has_write_password(src)

    flags = _STGM_READ | _STGM_SHARE_DENY_WRITE
    pythoncom.CoInitialize()
    storage = stream = None
    try:
        storage = pythoncom.StgOpenStorage(str(src), None, flags)
        names = {element[0] for element in storage.EnumElements()}
        if "EncryptedPackage" in names:
            return True
        if "WordDocument" in names:
            stream = storage.OpenStream(
                "WordDocument", None, _STGM_READ | _STGM_SHARE_EXCLUSIVE
            )
            fib = bytes(stream.Read(12))
            return len(fib) >= 12 and bool(fib[11] & _FIB_PASSWORD_BITS)
    except pythoncom.com_error as exc:
        raise ConversionError(MSG_FAILED) from exc
    finally:
        storage = stream = None  # release COM objects before uninitializing
        pythoncom.CoUninitialize()
    return False


def _word_pid(word) -> int | None:
    """Find a Word instance's process by giving its window a unique caption.

    FindWindow also sees hidden windows. Returns None if the window does not
    show up within a few seconds.
    """
    import win32gui
    import win32process

    caption = uuid.uuid4().hex
    word.Caption = caption
    deadline = time.monotonic() + _FIND_WORD_WINDOW_SECONDS
    while True:
        try:
            hwnd = win32gui.FindWindow(_WORD_WINDOW_CLASS, caption)
        except win32gui.error:
            hwnd = 0
        if hwnd:
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            return pid or None
        if time.monotonic() >= deadline:
            return None
        time.sleep(0.1)


def _terminate_word(pid: int) -> None:
    import win32api

    handle = win32api.OpenProcess(_PROCESS_TERMINATE, False, pid)
    try:
        win32api.TerminateProcess(handle, 1)
    finally:
        win32api.CloseHandle(handle)


def _on_word_timeout(pid: int, timed_out: threading.Event) -> None:
    logging.warning(
        "Word did not finish within %s s; terminating it", WORD_TIMEOUT_SECONDS
    )
    timed_out.set()
    try:
        _terminate_word(pid)
    except Exception:
        logging.exception("Could not terminate Word (pid %s)", pid)


def _export_with_word(word, src: Path, out_pdf: Path) -> None:
    """Open src read-only in the given Word instance and export it to out_pdf."""
    # A wrong PasswordDocument does not make Word 16 raise (it prompts), so
    # protected files must be refused by the pre-check (_is_encrypted) before
    # Word is started; for unprotected files these values are ignored.
    doc = word.Documents.Open(
        FileName=str(src),
        ConfirmConversions=False,
        ReadOnly=True,
        AddToRecentFiles=False,
        PasswordDocument="\x01",
        PasswordTemplate="\x01",
        OpenAndRepair=False,
        NoEncodingDialog=True,
    )
    try:
        doc.ExportAsFixedFormat(
            OutputFileName=str(out_pdf), ExportFormat=_WD_EXPORT_FORMAT_PDF
        )
    finally:
        doc.Close(SaveChanges=0)
        doc = None


def _word_to_pdf(src: Path, out_pdf: Path) -> None:
    import pythoncom
    import pywintypes
    import win32com.client

    pythoncom.CoInitialize()
    word = None
    timer = None
    timed_out = threading.Event()
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        word.AutomationSecurity = _MSO_AUTOMATION_SECURITY_FORCE_DISABLE
        pid = _word_pid(word)
        if pid is None:
            logging.warning("Could not find the Word process; conversion has no timeout")
        else:
            timer = threading.Timer(
                WORD_TIMEOUT_SECONDS, _on_word_timeout, (pid, timed_out)
            )
            timer.daemon = True
            timer.start()
        _export_with_word(word, src, out_pdf)
        if timed_out.is_set():
            raise ConversionError(MSG_FAILED)
    except ConversionError:
        raise
    except pywintypes.com_error as exc:
        text = " ".join(str(part) for part in (*exc.args, exc)).lower()
        password = "password" in text and not timed_out.is_set()
        raise ConversionError(MSG_PASSWORD_WORD if password else MSG_FAILED) from exc
    except Exception as exc:
        raise ConversionError(MSG_FAILED) from exc
    finally:
        if word is not None:
            try:
                word.Quit(SaveChanges=0)
            except Exception:
                pass  # expected when the watchdog has terminated Word
        if timer is not None:
            timer.cancel()
        word = None  # drop COM proxies before uninitializing
        pythoncom.CoUninitialize()


def _libreoffice_to_pdf(soffice: Path, src: Path, outdir: Path) -> Path:
    # A fresh profile keeps an already running LibreOffice from taking the job.
    profile = (outdir / "lo-profile").as_uri()
    try:
        subprocess.run(
            [str(soffice), f"-env:UserInstallation={profile}", "--headless",
             "--convert-to", "pdf", "--outdir", str(outdir), str(src)],
            timeout=120,
            creationflags=subprocess.CREATE_NO_WINDOW,
            check=False,
        )
    except (OSError, subprocess.SubprocessError) as exc:
        raise ConversionError(MSG_FAILED) from exc
    result = outdir / f"{src.stem}.pdf"
    if not result.exists():
        raise ConversionError(MSG_FAILED)
    return result


def docx_to_pdf(src: Path, dst: Path) -> Path:
    """Convert a Word file to PDF using Word, else LibreOffice.

    Returns the path written: dst, or a numbered name if dst appeared meanwhile.
    """
    if _is_encrypted(src):
        raise ConversionError(MSG_PASSWORD_WORD)
    use_word = word_installed()
    soffice = None if use_word else find_libreoffice()
    if not use_word and soffice is None:
        raise ConversionError(MSG_NO_OFFICE)

    part = part_path(dst)
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as tmp:
        tmp_dir = Path(tmp)
        if use_word:
            result = tmp_dir / f"{src.stem}.pdf"
            _word_to_pdf(src, result)
        else:
            result = _libreoffice_to_pdf(soffice, src, tmp_dir)
        try:
            shutil.move(str(result), str(part))
            return finish_part(part, dst)
        except ConversionError:
            part.unlink(missing_ok=True)
            raise
        except OSError as exc:
            part.unlink(missing_ok=True)
            raise ConversionError(MSG_CANNOT_SAVE) from exc
