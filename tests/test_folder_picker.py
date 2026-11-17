"""services/folder_picker.py's contents(): the files of a chosen folder AND of every folder under it, any depth -
approved 2026-11-02, so choosing e.g. "Bpr-irembo" also reads its "July 2026", "August 2026", ... folders."""
from pathlib import Path

from services import folder_picker as fp


def _write(path: Path, data: bytes = b"x"):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(data)


def test_files_of_nested_folders_are_found_too(tmp_path):
    _write(tmp_path / "top.xlsx")
    _write(tmp_path / "July 2026" / "1.xlsx")
    _write(tmp_path / "July 2026" / "2.csv")
    _write(tmp_path / "August 2026" / "Week 1" / "deep.xlsx")       # two levels down
    found = fp.contents(tmp_path, ["xlsx", "csv"])
    assert sorted(n for n, _ in found) == ["1.xlsx", "2.csv", "deep.xlsx", "top.xlsx"]


def test_only_the_asked_extensions_are_kept(tmp_path):
    _write(tmp_path / "statement.xlsx")
    _write(tmp_path / "notes.txt")
    found = fp.contents(tmp_path, ["xlsx"])
    assert [n for n, _ in found] == ["statement.xlsx"]


def test_a_hidden_folder_is_not_looked_into(tmp_path):
    _write(tmp_path / "seen.xlsx")
    _write(tmp_path / ".git" / "hidden.xlsx")
    found = fp.contents(tmp_path, ["xlsx"])
    assert [n for n, _ in found] == ["seen.xlsx"]


def test_an_excel_lock_file_is_skipped():
    import tempfile
    with tempfile.TemporaryDirectory() as d:
        d = Path(d)
        _write(d / "report.xlsx")
        _write(d / "~$report.xlsx")                     # Excel's own lock file for an open workbook
        found = fp.contents(d, ["xlsx"])
        assert [n for n, _ in found] == ["report.xlsx"]


def test_two_folders_can_each_have_a_file_of_the_same_name(tmp_path):
    _write(tmp_path / "A" / "1.xlsx", b"from A")
    _write(tmp_path / "B" / "1.xlsx", b"from B")
    found = fp.contents(tmp_path, ["xlsx"])
    assert sorted(content for _, content in found) == [b"from A", b"from B"]
