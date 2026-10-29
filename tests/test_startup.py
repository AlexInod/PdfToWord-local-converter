import ctypes
import os

from app.startup import (
    BROWSER_ARGS,
    MISSING_WEBVIEW2_TEXT,
    STARTUP_FAILED_TEXT,
    configure_environment,
    log_path,
    show_missing_webview2,
    show_startup_failed,
    webview2_version,
)


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


def test_message_boxes_are_topmost_errors(monkeypatch):
    shown = []
    monkeypatch.setattr(
        ctypes.windll.user32,
        "MessageBoxW",
        lambda hwnd, text, title, flags: shown.append((text, title, flags)),
    )
    log = r"C:\Users\x\AppData\Local\PDF Converter\app.log"
    show_missing_webview2()
    show_startup_failed(log)
    flags = 0x10 | 0x40000 | 0x10000
    assert shown == [
        (MISSING_WEBVIEW2_TEXT, "PDF Converter", flags),
        ("PDF Converter could not start. Details were written to:\n" + log,
         "PDF Converter", flags),
    ]
    assert STARTUP_FAILED_TEXT.format(log=log) == shown[1][0]
