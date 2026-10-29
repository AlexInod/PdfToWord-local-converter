import ctypes
import os
import winreg
from pathlib import Path

WEBVIEW2_CLIENT_KEY = r"Software\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}"
MISSING_WEBVIEW2_TEXT = (
    "This app needs Microsoft Edge WebView2, which is free from Microsoft:\n"
    "https://developer.microsoft.com/microsoft-edge/webview2/"
)
STARTUP_FAILED_TEXT = "PDF Converter could not start. Details were written to:\n{log}"
BROWSER_ARGS = "--disable-features=msSmartScreenProtection --disable-background-networking"


def webview2_version() -> str | None:
    locations = (
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node" + WEBVIEW2_CLIENT_KEY[len("Software"):]),
        (winreg.HKEY_LOCAL_MACHINE, WEBVIEW2_CLIENT_KEY),
        (winreg.HKEY_CURRENT_USER, WEBVIEW2_CLIENT_KEY),
    )
    for root, subkey in locations:
        try:
            with winreg.OpenKey(root, subkey) as key:
                value, _ = winreg.QueryValueEx(key, "pv")
        except OSError:
            continue
        if value and value != "0.0.0.0":
            return value
    return None


# MB_ICONERROR | MB_TOPMOST | MB_SETFOREGROUND: nothing else is on screen yet.
MESSAGE_BOX_FLAGS = 0x10 | 0x40000 | 0x10000


def _message_box(text: str) -> None:
    ctypes.windll.user32.MessageBoxW(0, text, "PDF Converter", MESSAGE_BOX_FLAGS)


def show_missing_webview2() -> None:
    _message_box(MISSING_WEBVIEW2_TEXT)


def show_startup_failed(log: Path | str) -> None:
    _message_box(STARTUP_FAILED_TEXT.format(log=log))


def close_splash() -> None:
    """Close the PyInstaller splash screen, if there is one."""
    try:
        import pyi_splash

        pyi_splash.close()
    except Exception:
        pass


def configure_environment() -> None:
    os.environ.setdefault("WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS", BROWSER_ARGS)


def log_path() -> Path:
    base = os.environ.get("LOCALAPPDATA") or str(Path.home() / "AppData" / "Local")
    folder = Path(base) / "PDF Converter"
    folder.mkdir(parents=True, exist_ok=True)
    return folder / "app.log"
