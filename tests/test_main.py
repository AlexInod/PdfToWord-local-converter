from app.main import drop_paths


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
