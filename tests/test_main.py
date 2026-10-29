from pathlib import Path

import app.main
from app.main import Api, _explorer_select_command, drop_paths
from app.startup import log_path


def test_drop_paths_keeps_only_entries_with_full_path():
    event = {
        "dataTransfer": {
            "files": [
                {"name": "a.pdf", "pywebviewFullPath": "D:\\x\\a.pdf"},
                {"name": "b.txt"},
            ]
        }
    }
    assert drop_paths(event) == ["D:\\x\\a.pdf"]


def test_drop_paths_tolerates_missing_data():
    assert drop_paths({}) == []


def test_api_public_names_are_exactly_the_five_methods():
    assert {n for n in dir(Api()) if not n.startswith("_")} == {
        "set_mode", "choose_files", "open_folder", "show_file", "office_available"
    }


def test_explorer_select_keeps_path_with_comma_in_one_argument():
    command = _explorer_select_command(Path(r"D:\a,b (1).docx"))
    assert command == r'explorer /select,"D:\a,b (1).docx"'


def _patch_startup(monkeypatch, calls):
    monkeypatch.setattr(app.main, "_configure_logging", lambda: None)
    monkeypatch.setattr(app.main, "configure_environment", lambda: None)
    monkeypatch.setattr(app.main, "close_splash", lambda: calls.append("splash"))
    monkeypatch.setattr(
        app.main, "show_missing_webview2", lambda: calls.append("webview2")
    )
    monkeypatch.setattr(
        app.main, "show_startup_failed", lambda log: calls.append(("failed", log))
    )


def test_missing_webview2_closes_splash_before_message(monkeypatch):
    calls = []
    _patch_startup(monkeypatch, calls)
    monkeypatch.setattr(app.main, "webview2_version", lambda: None)
    app.main.main()
    assert calls == ["splash", "webview2"]


def test_unexpected_startup_error_is_logged_and_shown(monkeypatch, caplog):
    calls = []
    _patch_startup(monkeypatch, calls)

    def broken():
        raise RuntimeError("boom")

    monkeypatch.setattr(app.main, "webview2_version", broken)
    app.main.main()
    assert calls == ["splash", ("failed", str(log_path()))]
    assert "boom" in caplog.text
