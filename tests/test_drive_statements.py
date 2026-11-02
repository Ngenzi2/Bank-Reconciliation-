"""Ecobank statements read from Google Drive (26/10/2026): picked by the days INSIDE each file, IREMBO PAY files
only, at most 10 downloads started per second, the days remembered between two reads."""
import time

import pytest

import pandas as pd

import engine
from integrations.google_drive import DriveFile
from services import drive_statements as ds


@pytest.fixture(autouse=True)
def nothing_kept_between_tests():
    """27/10/2026: the statements read and the folder lists are kept in memory - each test starts with none."""
    ds._CONTENT.clear(); ds._LISTING.clear(); ds.JOBS.clear()
    yield


class FakeDrive:
    """A Drive folder: {name: day inside the file}; the content is the day, read back by a fake read_bank_file."""

    def __init__(self, days):
        self.files = [DriveFile(f"id{i}", n, modified="2026-10-26T08:00:00Z") for i, n in enumerate(days)]
        self.days, self.started = days, []

    def list_files(self, folder_id, extensions=(), depth=3):
        return list(self.files)

    def thread_http(self):
        return object()

    def download(self, file_id, mime="", http=None):
        self.started.append(time.monotonic())
        name = next(f.name for f in self.files if f.id == file_id)
        return self.days[name].encode()


def fake_read(name, content, bank=None):
    text = content.decode()
    if text == "bad":
        raise engine.ReconError(f"{name}: not a statement")
    first, last = (pd.Timestamp(d) for d in (text.split("|") * 2)[:2])
    return engine.UploadedFile(name, "bank", engine.file_digest(content), pd.DataFrame(), "", first, last, 1)


def test_picked_by_the_days_inside_not_the_name(tmp_path, monkeypatch):
    monkeypatch.setattr(engine, "read_bank_file", fake_read)
    drive = FakeDrive({"IREMBO PAY2026-08-04-00-01-04.xlsx": "2026-08-03",        # the name says 04/08: holds 03/08
                       "IREMBO PAY2026-08-05-00-00-36.xlsx": "2026-08-04",
                       "IREMBO PAY2026-08-11-00-00-49.xlsx": "2026-08-09|2026-08-10",   # touches the period's last day
                       "IREMBO PAY2026-08-12-00-00-58.xlsx": "2026-08-11",
                       "IREMBO_TACCT_TRANS_050826.XLSX": "2026-08-05",                # not an IREMBO PAY file
                       "IREMBO_MONTHLY_TRAN_17082026.XLSX": "2026-08-05",
                       "IREMBO PAY broken.xlsx": "bad"})
    got = ds.pick(drive, "folder", "ECO", "2026-08-03", "2026-08-09", tmp_path / "c.pkl")
    assert [n for n, _ in got.files] == ["IREMBO PAY2026-08-04-00-01-04.xlsx", "IREMBO PAY2026-08-05-00-00-36.xlsx",
                                         "IREMBO PAY2026-08-11-00-00-49.xlsx"]
    assert got.considered == 5 and got.unreadable == ["IREMBO PAY broken.xlsx"]

    # the second time only the files of the period are downloaded (the others' days are remembered)
    drive.started.clear()
    again = ds.pick(drive, "folder", "ECO", "2026-08-03", "2026-08-09", tmp_path / "c.pkl")
    assert [n for n, _ in again.files] == [n for n, _ in got.files] and again.downloaded == 3


def test_at_most_ten_downloads_start_per_second(tmp_path, monkeypatch):
    monkeypatch.setattr(engine, "read_bank_file", fake_read)
    drive = FakeDrive({f"IREMBO PAY2026-08-{d:02d}.xlsx": f"2026-08-{d:02d}" for d in range(1, 26)})
    ds.pick(drive, "folder", "ECO", "2026-08-01", "2026-08-31", tmp_path / "c.pkl")
    s = sorted(drive.started)
    assert len(s) == 25
    assert all(b - a >= 0.9 for a, b in zip(s, s[10:]))        # any 11 starts span about one second (timing margin)


def test_a_file_slow_to_come_is_asked_again(tmp_path, monkeypatch):
    """27/10/2026: Google is sometimes slow to answer one file (20-40 s): after 3 s the same file is asked again on
    a new connection and the first answer is used."""
    monkeypatch.setattr(engine, "read_bank_file", fake_read)
    drive = FakeDrive({"IREMBO PAY slow.xlsx": "2026-08-03", "IREMBO PAY fast.xlsx": "2026-08-04"})
    calls = {}
    real = drive.download

    def download(file_id, mime="", http=None):
        calls[file_id] = calls.get(file_id, 0) + 1
        if file_id == "id0" and calls[file_id] == 1:
            time.sleep(6)                                         # the first request for this file hangs
        return real(file_id, mime, http)
    drive.download = download
    t = time.monotonic()
    got = ds.pick(drive, "folder", "ECO", "2026-08-03", "2026-08-09", tmp_path / "c.pkl")
    assert time.monotonic() - t < 5.5 and len(got.files) == 2 and calls["id0"] == 2


def test_a_statement_read_once_is_not_downloaded_again(tmp_path, monkeypatch):
    """27/10/2026: switching back to a bank and period already read is instant - nothing is downloaded again."""
    monkeypatch.setattr(engine, "read_bank_file", fake_read)
    drive = FakeDrive({"IREMBO PAY2026-08-04.xlsx": "2026-08-03", "IREMBO PAY2026-08-05.xlsx": "2026-08-04"})
    ds.pick(drive, "folder", "ECO", "2026-08-03", "2026-08-09", tmp_path / "c.pkl")
    drive.started.clear()
    again = ds.pick(drive, "folder", "ECO", "2026-08-03", "2026-08-09", tmp_path / "c.pkl")
    assert len(again.files) == 2 and drive.started == []


def test_the_reading_runs_in_the_background(tmp_path, monkeypatch):
    """27/10/2026: start_job returns at once; the job says how far it is and gives the result when done."""
    monkeypatch.setattr(engine, "read_bank_file", fake_read)
    drive = FakeDrive({f"IREMBO PAY2026-08-{d:02d}.xlsx": f"2026-08-{d:02d}" for d in range(1, 21)})
    t = time.monotonic()
    job = ds.start_job(("ECO", "k"), lambda: drive, "folder", "ECO", "2026-08-01", "2026-08-31", tmp_path / "c.pkl")
    assert time.monotonic() - t < 0.5 and ds.start_job(("ECO", "k"), None, "", "ECO", 0, 0, tmp_path) is job
    while not job["done"]:
        time.sleep(0.05)
    assert job["error"] is None and len(job["result"].files) == 20 and job["n"] == 20
