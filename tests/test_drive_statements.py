"""Ecobank statements read from Google Drive (26/10/2026): picked by the days INSIDE each file, IREMBO PAY files
only, at most 10 downloads started per second, the days remembered between two reads."""
import time

import pandas as pd

import engine
from integrations.google_drive import DriveFile
from services import drive_statements as ds


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
    assert all(b - a >= 0.99 for a, b in zip(s, s[10:]))       # any 11 starts span at least one second
