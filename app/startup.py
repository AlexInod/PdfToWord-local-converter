import ctypes
import os
import winreg
from pathlib import Path

WEBVIEW2_CLIENT_KEY = r"Software\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}"
MISSING_WEBVIEW2_TEXT = (
    "This app needs Microsoft Edge WebView2, which is free from Microsoft:\n"
    "https://developer.microsoft.com/microsoft-edge/webview2/"
)
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


def show_missing_webview2() -> None:
    ctypes.windll.user32.MessageBoxW(0, MISSING_WEBVIEW2_TEXT, "PDF Converter", 0x10)


def configure_environment() -> None:
    os.environ.setdefault("WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS", BROWSER_ARGS)


def log_path() -> Path:
    base = os.environ.get("LOCALAPPDATA") or str(Path.home() / "AppData" / "Local")
    folder = Path(base) / "PDF Converter"
    folder.mkdir(parents=True, exist_ok=True)
    return folder / "app.log"
