"""Word -> PDF via Microsoft Word (COM), falling back to LibreOffice."""
import re
import shutil
import subprocess
import tempfile
from pathlib import Path

from app.convert import (
    ConversionError,
    MSG_CANNOT_SAVE,
    MSG_FAILED,
    MSG_NO_OFFICE,
    MSG_PASSWORD_WORD,
)

_LIBREOFFICE_PATHS = (
    Path(r"C:\Program Files\LibreOffice\program\soffice.exe"),
    Path(r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"),
)
_WD_EXPORT_FORMAT_PDF = 17
_STGM_READ = 0x00
_STGM_SHARE_DENY_WRITE = 0x20


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
    would hang the worker, so encryption is detected from the file itself.
    Encrypted .docx files are OLE containers holding an EncryptedPackage
    stream; legacy .doc files set the fEncrypted bit in the FIB.
    """
    import pythoncom

    if _has_write_password(src):
        return True
    flags = _STGM_READ | _STGM_SHARE_DENY_WRITE
    try:
        storage = pythoncom.StgOpenStorage(str(src), None, flags)
    except pythoncom.com_error:
        return False  # not an OLE container: plain .docx (zip) or unreadable
    names = {element[0] for element in storage.EnumElements()}
    if "EncryptedPackage" in names:
        return True
    if "WordDocument" in names:
        fib = bytes(storage.OpenStream("WordDocument", None, flags).Read(12))
        return len(fib) >= 12 and bool(fib[11] & 0x01)
    return False


def _word_to_pdf(src: Path, out_pdf: Path) -> None:
    import pythoncom
    import pywintypes
    import win32com.client

    pythoncom.CoInitialize()
    word = doc = None
    try:
        if _is_encrypted(src):
            raise ConversionError(MSG_PASSWORD_WORD)
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        # A wrong password makes Word raise instead of prompting; an
        # unprotected file ignores it.
        doc = word.Documents.Open(
            FileName=str(src),
            ReadOnly=True,
            AddToRecentFiles=False,
            PasswordDocument="\x01",
            PasswordTemplate="\x01",
        )
        try:
            doc.ExportAsFixedFormat(
                OutputFileName=str(out_pdf), ExportFormat=_WD_EXPORT_FORMAT_PDF
            )
        finally:
            doc.Close(SaveChanges=0)
            doc = None
    except ConversionError:
        raise
    except pywintypes.com_error as exc:
        text = " ".join(str(part) for part in (*exc.args, exc)).lower()
        raise ConversionError(
            MSG_PASSWORD_WORD if "password" in text else MSG_FAILED
        ) from exc
    except Exception as exc:
        raise ConversionError(MSG_FAILED) from exc
    finally:
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        word = doc = None  # drop COM proxies before uninitializing
        pythoncom.CoUninitialize()


def _libreoffice_to_pdf(soffice: Path, src: Path, outdir: Path) -> Path:
    try:
        subprocess.run(
            [str(soffice), "--headless", "--convert-to", "pdf",
             "--outdir", str(outdir), str(src)],
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


def docx_to_pdf(src: Path, dst: Path) -> None:
    """Convert a Word file to PDF at dst using Word, else LibreOffice."""
    use_word = word_installed()
    soffice = None if use_word else find_libreoffice()
    if not use_word and soffice is None:
        raise ConversionError(MSG_NO_OFFICE)

    with tempfile.TemporaryDirectory() as tmp:
        tmp_dir = Path(tmp)
        if use_word:
            result = tmp_dir / f"{src.stem}.pdf"
            _word_to_pdf(src, result)
        else:
            result = _libreoffice_to_pdf(soffice, src, tmp_dir)
        try:
            shutil.move(str(result), str(dst))
        except OSError as exc:
            raise ConversionError(MSG_CANNOT_SAVE) from exc
