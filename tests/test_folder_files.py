"""Finding the statement and IremboPay export files for a period in a local folder (dianah branch): the usual
acceptance window, the 45-day re-send lookback, caching, and "file of another bank"."""
import io

import pandas as pd
from openpyxl import Workbook

import config
import engine as e
from boa_fixtures import payments, statement
from services import folder_files as ff

SEP = e.make_period("month", "2026-09-01")


def spec(path):
    return config.FolderSpec("local", str(path))


class FakeDrive:
    """A drive=... good enough for folder_files: find_folder()/list_files()/download(), no real Drive API call -
    the shared Statements folder's own layout (a '<Month Year>' sub-folder per month)."""
    def __init__(self, months):
        self.months = months                           # {"August 2026": [(DriveFile-like, content bytes), ...]}
        self.ids = {name: f"folder-{name}" for name in months}

    def find_folder(self, parent_id, name):
        return self.ids.get(name)

    def list_files(self, folder_id, extensions=()):
        ends = tuple(f".{x.lower()}" for x in extensions)
        name = next((n for n, i in self.ids.items() if i == folder_id), None)
        files = self.months.get(name, [])
        return [f for f, _ in files if not ends or f.name.lower().endswith(ends)]

    def download(self, file_id, mime=""):
        return next(c for files in self.months.values() for f, c in files if f.id == file_id)


def drive_file(id, name, modified="2026-09-10T00:00:00.000Z"):
    from integrations.google_drive import DriveFile
    return DriveFile(id, name, modified=modified)


def im_statement(rows):
    """rows: [(Transaction Date 'dd/mm/yyyy', Value Date 'dd/mm/yyyy', Description, Tran Id, Credit)]."""
    wb = Workbook()
    ws = wb.active
    ws.append(["Transaction Date", "Value Date", "Description", "Tran Id", "Credit", "Balance (RWF)"])
    for row in rows:
        ws.append(list(row) + [0])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def write(path, name, content):
    path.mkdir(parents=True, exist_ok=True)
    (path / name).write_bytes(content)
    return path / name


def test_statements_split_into_current_and_reference(tmp_path):
    stm_dir, exp_dir = tmp_path / "stm", tmp_path / "exp"
    write(stm_dir, "august.xlsx", statement([("20/08/2026", "R1", "ePay irembo 880000000001", None, 1000)]))
    write(stm_dir, "september.xlsx", statement([("05/09/2026", "R2", "ePay irembo 880000000002", None, 2000)]))
    write(exp_dir, "export.csv", payments([
        dict(invoice="880000000001", amount=1000, paid_at="2026-09-05 09:00:00", ref="R1"),
        dict(invoice="880000000002", amount=2000, paid_at="2026-09-05 09:00:00", ref="R2"),
    ]))
    folders = {"BOA": {"statements": spec(stm_dir), "exports": spec(exp_dir)}}
    found = ff.find_files(SEP, "BOA", folders)
    assert not found.problems
    assert [f.name for f in found.statements] == ["september.xlsx"]
    assert [f.name for f in found.reference] == ["august.xlsx"]
    assert found.reference[0].reference is True
    assert [f.name for f in found.exports] == ["export.csv"]
    assert len(ff.all_files(found)) == 3
    lines = ff.describe(found)
    assert lines[0].startswith("Statements found: 05 Sep")
    assert lines[1].startswith("IremboPay export: taken")


def test_a_file_too_old_for_even_the_lookback_window_is_left_out(tmp_path):
    stm_dir = tmp_path / "stm"
    write(stm_dir, "ancient.xlsx", statement([("01/01/2026", "R9", "ePay irembo 880000000009", None, 500)]))
    folders = {"BOA": {"statements": spec(stm_dir), "exports": None}}
    found = ff.find_files(SEP, "BOA", folders)
    assert found.statements == [] and found.reference == []
    assert "No IremboPay exports folder is set" in found.problems[0]


def test_no_folder_set_is_a_problem_not_a_crash():
    found = ff.find_files(SEP, "BOA", {})
    assert len(found.problems) == 2 and all("No " in p for p in found.problems)


def test_missing_local_folder_is_reported(tmp_path):
    folders = {"BOA": {"statements": spec(tmp_path / "missing"), "exports": None}}
    found = ff.find_files(SEP, "BOA", folders)
    assert "was not found" in found.problems[0]


def test_a_file_of_another_bank_is_named(tmp_path):
    stm_dir = tmp_path / "stm"
    write(stm_dir, "im_file.xlsx", im_statement([("05/09/2026", "05/09/2026", "IREMBO PAYMENT 8800000000001",
                                                  "T1", 1000)]))
    folders = {"BOA": {"statements": spec(stm_dir), "exports": None}}
    found = ff.find_files(SEP, "BOA", folders)
    assert found.statements == []
    assert any("this is a I&M statement, not Bank of Africa" in p for p in found.problems)


def test_reading_is_cached_until_the_file_changes(tmp_path):
    stm_dir = tmp_path / "stm"
    f = write(stm_dir, "sept.xlsx", statement([("05/09/2026", "R1", "ePay irembo 880000000001", None, 1000)]))
    folders = {"BOA": {"statements": spec(stm_dir), "exports": None}}
    first = ff.find_files(SEP, "BOA", folders).statements[0]
    again = ff.find_files(SEP, "BOA", folders).statements[0]
    assert first is again                                          # same object: not re-read

    f.write_bytes(statement([("06/09/2026", "R2", "ePay irembo 880000000002", None, 2000)]))
    import os, time
    os.utime(f, (time.time() + 5, time.time() + 5))                 # force a different mtime
    changed = ff.find_files(SEP, "BOA", folders).statements[0]
    assert changed is not first and changed.first == pd.Timestamp("2026-09-06")


def test_a_local_gap_is_filled_from_the_shared_drive_statements_folder(tmp_path, monkeypatch):
    # local has September only; the shared Drive "Statements" folder also has an older file the 45-day lookback
    # needs (reference only, never required) - approved 2026-10-26
    monkeypatch.setenv("BANK_RECON_STATEMENTS_FOLDER_ID", "ROOT123")
    stm_dir = tmp_path / "stm"
    write(stm_dir, "september.xlsx", statement([("05/09/2026", "R2", "ePay irembo 880000000002", None, 2000)]))
    content = statement([("20/08/2026", "R1", "ePay irembo 880000000001", None, 1000)])
    drive = FakeDrive({"August 2026": [(drive_file("D1", "august.xlsx"), content)]})
    folders = {"BOA": {"statements": spec(stm_dir), "exports": None}}
    found = ff.find_files(SEP, "BOA", folders, drive=drive)
    assert [f.name for f in found.statements] == ["september.xlsx"]
    assert [f.name for f in found.reference] == ["august.xlsx"]
    assert found.reference[0].reference is True
    assert found.reference[0].from_drive is True
    assert not found.statements[0].from_drive                      # the local file is not marked as Drive's
    assert "(1 from Drive)" in ff.describe(found)[0]


def test_drive_is_checked_every_time_not_only_when_local_is_short(tmp_path, monkeypatch):
    # local already has a September file; the shared Drive folder has another one, for a different day - it is
    # picked up too, the way the notebooks always check Drive for a new day's statement, approved 2026-10-26
    monkeypatch.setenv("BANK_RECON_STATEMENTS_FOLDER_ID", "ROOT123")
    stm_dir = tmp_path / "stm"
    write(stm_dir, "sept5.xlsx", statement([("05/09/2026", "R1", "ePay irembo 880000000001", None, 1000)]))
    content = statement([("10/09/2026", "R2", "ePay irembo 880000000002", None, 2000)])
    drive = FakeDrive({"September 2026": [(drive_file("D2", "sept10.xlsx"), content)]})
    folders = {"BOA": {"statements": spec(stm_dir), "exports": None}}
    found = ff.find_files(SEP, "BOA", folders, drive=drive)
    assert sorted(f.name for f in found.statements) == ["sept10.xlsx", "sept5.xlsx"]


def test_a_drive_file_already_found_locally_is_not_duplicated(tmp_path, monkeypatch):
    monkeypatch.setenv("BANK_RECON_STATEMENTS_FOLDER_ID", "ROOT123")
    stm_dir = tmp_path / "stm"
    content = statement([("05/09/2026", "R1", "ePay irembo 880000000001", None, 1000)])
    write(stm_dir, "sept5.xlsx", content)
    drive = FakeDrive({"September 2026": [(drive_file("D3", "sept5_copy.xlsx"), content)]})   # same content, on Drive too
    folders = {"BOA": {"statements": spec(stm_dir), "exports": None}}
    found = ff.find_files(SEP, "BOA", folders, drive=drive)
    assert [f.name for f in found.statements] == ["sept5.xlsx"]        # the local copy, not a second one from Drive


def test_drive_is_not_asked_when_the_shared_folder_is_not_set(tmp_path):
    # the default (no statements_root_folder_id): behaves exactly as before this fallback existed
    stm_dir = tmp_path / "stm"
    write(stm_dir, "september.xlsx", statement([("05/09/2026", "R1", "ePay irembo 880000000001", None, 1000)]))

    class Boom(FakeDrive):
        def find_folder(self, *a, **k):
            raise AssertionError("Drive should not have been asked")
    folders = {"BOA": {"statements": spec(stm_dir), "exports": None}}
    found = ff.find_files(SEP, "BOA", folders, drive=Boom({}))
    assert [f.name for f in found.statements] == ["september.xlsx"]
