import os

from app.startup import BROWSER_ARGS, configure_environment, log_path, webview2_version


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
