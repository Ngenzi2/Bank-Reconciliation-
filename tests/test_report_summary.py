"""The small summary file next to each report on Drive (services/report_summary.py, 11/11/2026): the dashboard's
figures of a report, written when it is saved, read instead of the Excel when it belongs to the Excel on Drive."""
import hashlib
import json

import numpy as np
import pandas as pd

import engine
from integrations.google_drive import DriveError, DriveFile
from services import dashboard, facts_cache, report_history, report_summary

XLSX = "IM_recon_summary_01-09-2026_to_30-09-2026.xlsx"


class Rec:
    def __init__(self, md5="abc"):
        self.md5, self.name, self.file_id = md5, XLSX, "X"


def facts(record=None):
    month = dashboard.Fact("IM", "month", pd.Timestamp("2026-09-01"), pd.Timestamp("2026-09-30"),
                           {"processed": {"RWF": np.int64(10), "USD": 2}}, {"processed": {"RWF": np.float64(1500.5)}},
                           {("Rwanda Revenue Authority", "RWF"): (np.int64(3), 900.25)}, {"RWF": 1}, True, record,
                           "month")
    day = dashboard.Fact("IM", "day", pd.Timestamp("2026-09-07"), pd.Timestamp("2026-09-07"),
                         {"processed": {"RWF": 4}}, {"processed": {"RWF": 600.0}}, None, {}, False, record, "month")
    return [month, day]


def test_the_figures_come_back_the_same():
    rec = Rec()
    data = report_summary.to_bytes(facts(), rec.md5)
    assert len(data) < 5000                                               # a few KB, not a report
    assert report_summary.from_bytes(data, rec) == facts(rec)             # the same facts, attached to the report
    assert report_summary.name_for(XLSX) == "IM_recon_summary_01-09-2026_to_30-09-2026.summary.json"


def test_a_summary_of_another_content_version_or_a_broken_one_is_not_used(monkeypatch):
    data = report_summary.to_bytes(facts(), "abc")
    assert report_summary.from_bytes(data, Rec("other")) is None          # the Excel was replaced since
    assert report_summary.from_bytes(data, Rec("")) is None               # no checksum on Drive: not trusted
    assert report_summary.from_bytes(b"{not json", Rec()) is None
    assert report_summary.from_bytes(DriveError("not downloaded"), Rec()) is None
    monkeypatch.setattr(facts_cache, "VERSION", facts_cache.VERSION + 1)  # the figures are read another way now
    assert report_summary.from_bytes(data, Rec()) is None


class FakeDrive:
    def __init__(self, fail=()):
        self.files, self.fail = {}, set(fail)

    def find_folder(self, parent, name):
        return f"{parent}/{name}"

    find_or_create_folder = find_folder

    def forget_folders(self):
        pass

    def thread_http(self):
        return None

    def list_children(self, parent, name=None, folders_only=False):
        return [DriveFile(fid, n) for fid, (n, _, p) in self.files.items() if p == parent and name in (None, n)]

    def upload(self, parent, name, data, mime, replace_id=None, http=None):
        if any(name.endswith(x) for x in self.fail):
            raise DriveError(f"{name}: refused")
        fid = replace_id or f"id{len(self.files)}"
        self.files[fid] = (name, data, parent)
        return DriveFile(fid, name, mime, len(data), "2026-10-01T10:00:00Z", hashlib.md5(data).hexdigest(), "", parent)


def test_saving_a_report_writes_its_summary_next_to_it_and_replaces_it(monkeypatch):
    seen = []
    monkeypatch.setattr(report_history, "facts_from_bytes", lambda record, data, session: seen.append(data) or facts())
    bank, period = engine.get_bank("IM"), engine.make_period("month", "2026-09-01")
    drive = FakeDrive()
    files = {"xlsx": b"excel one", "pdf": b"pdf"}
    results = report_history.save_report(drive, "root", bank, period, files, session=object())
    assert all(r.ok for r in results) and seen == [b"excel one"]             # read from the Excel just saved
    summary = [(fid, n, d) for fid, (n, d, _) in drive.files.items() if n.endswith(".summary.json")]
    assert [n for _, n, _ in summary] == [report_summary.name_for(XLSX)]
    assert json.loads(summary[0][2])["xlsx_md5"] == hashlib.md5(b"excel one").hexdigest()

    existing = report_history.existing_reports(drive, "root", bank, period)  # saved again: replaced, not added
    report_history.save_report(drive, "root", bank, period, {"xlsx": b"excel two"}, "replace", existing,
                               session=object())
    again = [(fid, d) for fid, (n, d, _) in drive.files.items() if n.endswith(".summary.json")]
    assert [fid for fid, _ in again] == [summary[0][0]]
    assert json.loads(again[0][1])["xlsx_md5"] == hashlib.md5(b"excel two").hexdigest()


def test_a_summary_that_cannot_be_written_never_fails_the_save(monkeypatch):
    monkeypatch.setattr(report_history, "facts_from_bytes", lambda record, data, session: facts())
    bank, period = engine.get_bank("IM"), engine.make_period("month", "2026-09-01")
    drive = FakeDrive(fail=[".summary.json"])
    results = report_history.save_report(drive, "root", bank, period, {"xlsx": b"x", "pdf": b"p"}, session=object())
    assert [r.ok for r in results] == [True, True]
    assert not any(n.endswith(".summary.json") for n, _, _ in drive.files.values())
    assert report_summary.write(drive, "", XLSX, facts(), "abc") is None    # no folder known: nothing written


def test_the_summary_is_not_a_report_in_the_listing():
    assert report_history.parse_name(report_summary.name_for(XLSX)) is None
    assert report_history.parse_name(report_summary.bank_name("IM")) is None


# ---- one file per bank (13/11/2026) ---------------------------------------------------------------------------------
class Drive(FakeDrive):
    """FakeDrive that can also be read; refuse_replace: a file made by another account cannot be replaced."""
    refuse_replace = False

    def download(self, fid, mime="", http=None, as_xlsx=False):
        return self.files[fid][1]

    def upload(self, parent, name, data, mime, replace_id=None, http=None):
        if replace_id and self.refuse_replace:
            raise DriveError(f"{name}: not available to this app")
        return super().upload(parent, name, data, mime, replace_id, http)


def bank_files(drive):
    return [(fid, d) for fid, (n, d, _) in drive.files.items() if n == report_summary.bank_name("IM")]


def test_the_bank_file_holds_every_report_and_new_ones_are_added():
    drive = Drive()
    assert report_summary.update_bank(drive, "folder", "IM", {"a|1": facts()}) is not None
    assert report_summary.update_bank(drive, "folder", "IM", {"b|1": facts()[:1]}) is not None
    (fid, data), = bank_files(drive)                                      # one file, replaced: both reports in it
    saved = report_summary.bank_from_bytes(data)
    assert set(saved) == {"a|1", "b|1"}
    back = report_summary.facts_of(saved["a|1"], Rec())
    assert [(f.kind, f.counts, f.amounts, f.merchants) for f in back] == \
           [(f.kind, {m: {c: int(n) for c, n in v.items()} for m, v in f.counts.items()},
             {m: {c: float(a) for c, a in v.items()} for m, v in f.amounts.items()},
             None if f.merchants is None else {k: (int(n), a) for k, (n, a) in f.merchants.items()}) for f in facts()]
    assert all(isinstance(f.record, Rec) for f in back)

    report_summary.update_bank(drive, "folder", "IM", {"c|1": facts()}, keep=["b|1", "c|1"])   # a|1: no longer on Drive
    assert set(report_summary.bank_from_bytes(bank_files(drive)[0][1])) == {"b|1", "c|1"}
    assert report_summary.update_bank(drive, "", "IM", {"d|1": facts()}) is None               # no folder: nothing


def test_a_bank_file_of_another_version_or_a_broken_one_is_not_used(monkeypatch):
    drive = Drive()
    report_summary.update_bank(drive, "folder", "IM", {"a|1": facts()})
    data = bank_files(drive)[0][1]
    assert report_summary.bank_from_bytes(b"not json") == {} and report_summary.bank_from_bytes(None) == {}
    assert report_summary.facts_of([{"bank": "IM"}], Rec()) is None
    monkeypatch.setattr(facts_cache, "VERSION", facts_cache.VERSION + 1)
    assert report_summary.bank_from_bytes(data) == {}


def test_a_bank_file_that_cannot_be_replaced_gets_a_second_one_with_everything():
    drive = Drive()
    report_summary.update_bank(drive, "folder", "IM", {"a|1": facts()})
    drive.refuse_replace = True                                           # written by another account
    assert report_summary.update_bank(drive, "folder", "IM", {"b|1": facts()}) is not None
    files = bank_files(drive)
    assert len(files) == 2
    assert set(report_summary.bank_from_bytes(files[1][1])) == {"a|1", "b|1"}


def test_saving_a_report_adds_it_to_its_bank_file(monkeypatch):
    monkeypatch.setattr(report_history, "facts_from_bytes", lambda record, data, session: facts())
    bank, period = engine.get_bank("IM"), engine.make_period("month", "2026-09-01")
    drive = Drive()
    results = report_history.save_report(drive, "root", bank, period, {"xlsx": b"excel", "pdf": b"pdf"},
                                         session=object())
    xlsx = next(r.file for r in results if r.name.endswith(".xlsx"))
    record = report_history._record(xlsx, report_history.parse_name(xlsx.name))
    (fid, data), = bank_files(drive)
    assert drive.files[fid][2] == "root/I&M Bank"                         # in the bank's folder
    assert set(report_summary.bank_from_bytes(data)) == {facts_cache.key(record)}
