"""The dashboard's summary on Drive (services/report_summary.py): one file per bank holding the figures of every report
of the bank (13/11/2026); 18/11/2026: the only summary (no file next to each report), one file kept per bank."""
import hashlib

import numpy as np
import pytest
import pandas as pd

import engine
from integrations.google_drive import DriveError, DriveFile
from services import dashboard, facts_cache, report_history, report_summary

XLSX = "IM_recon_summary_01-09-2026_to_30-09-2026.xlsx"


@pytest.fixture(autouse=True)
def no_keepers(monkeypatch):
    """Anyone writes a bank's file, unless a test sets the keepers (report_summary.KEEPERS)."""
    monkeypatch.setattr(report_summary, "KEEPERS", {})


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


def test_saving_a_report_adds_it_to_its_bank_file_and_nothing_next_to_it(monkeypatch):
    # 18/11/2026: no <report>.summary.json any more: the Excel, the PDF and the bank's file only
    monkeypatch.setattr(report_history, "facts_from_bytes", lambda record, data, session: facts())
    bank, period = engine.get_bank("IM"), engine.make_period("month", "2026-09-01")
    drive = Drive()
    report_history.save_report(drive, "root", bank, period, {"xlsx": b"excel", "pdf": b"pdf"}, session=object())
    assert sorted(n for n, _, _ in drive.files.values()) == sorted([XLSX, XLSX.replace(".xlsx", ".pdf"),
                                                                     report_summary.bank_name("IM")])


def test_a_bank_file_that_cannot_be_written_never_fails_the_save(monkeypatch):
    monkeypatch.setattr(report_history, "facts_from_bytes", lambda record, data, session: facts())
    bank, period = engine.get_bank("IM"), engine.make_period("month", "2026-09-01")
    drive = Drive(fail=[report_summary.BANK_SUFFIX])
    results = report_history.save_report(drive, "root", bank, period, {"xlsx": b"x", "pdf": b"p"}, session=object())
    assert [r.ok for r in results] == [True, True] and not bank_files(drive)


def test_the_summary_is_not_a_report_in_the_listing():
    assert report_history.parse_name(report_summary.bank_name("IM")) is None


# ---- one file per bank (13/11/2026) ---------------------------------------------------------------------------------
class Drive(FakeDrive):
    """FakeDrive that can also be read; refuse_replace: a file made by another account cannot be replaced."""
    refuse_replace = False

    def download(self, fid, mime="", http=None, as_xlsx=False):
        return self.files[fid][1]

    refuse_trash = False

    def upload(self, parent, name, data, mime, replace_id=None, http=None):
        if replace_id and self.refuse_replace:
            raise DriveError(f"{name}: not available to this app")
        return super().upload(parent, name, data, mime, replace_id, http)

    def trash(self, fid, what=""):
        if self.refuse_trash:
            raise DriveError("not allowed")
        self.files.pop(fid)


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


def test_a_file_that_cannot_be_written_makes_no_copy():
    # 18/11/2026: Drive refuses to replace the file: nothing written, no second file (the dashboard adds the report
    # the next time, from its Excel)
    drive = Drive()
    report_summary.update_bank(drive, "folder", "IM", {"a|1": facts()})
    drive.refuse_replace = True
    assert report_summary.update_bank(drive, "folder", "IM", {"b|1": facts()}) is None
    (fid, data), = bank_files(drive)
    assert set(report_summary.bank_from_bytes(data)) == {"a|1"}
    drive.refuse_replace = False
    report_summary.update_bank(drive, "folder", "IM", {"b|1": facts()})
    assert set(report_summary.bank_from_bytes(bank_files(drive)[0][1])) == {"a|1", "b|1"}


def test_a_slow_drive_is_tried_again_before_a_second_file_is_made():
    drive = Drive()
    report_summary.update_bank(drive, "folder", "IM", {"a|1": facts()})
    real, calls = drive.upload, []

    def once_slow(parent, name, data, mime, replace_id=None, http=None):
        calls.append(replace_id)
        if len(calls) == 1:
            raise DriveError("no answer from Google Drive (TimeoutError)")
        return real(parent, name, data, mime, replace_id, http)
    drive.upload = once_slow
    report_summary.update_bank(drive, "folder", "IM", {"b|1": facts()})
    assert len(bank_files(drive)) == 1 and calls[0] == calls[1] is not None   # replaced on the second try


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


def test_the_dashboard_merges_a_banks_copies_without_adding_anything():
    # 18/11/2026: the dashboard's tidying: no new report, only the copies merged into one (the others to the bin)
    first, second = Drive(), Drive()
    report_summary.update_bank(first, "folder", "IM", {"a|1": facts()})
    report_summary.update_bank(second, "folder", "IM", {"b|1": facts()})
    drive = Drive()                                                     # two copies of the IM file, as made before
    drive.files = {"c1": first.files[bank_files(first)[0][0]], "c2": second.files[bank_files(second)[0][0]]}
    assert report_summary.update_bank(drive, "folder", "IM", {}) is not None
    (fid, data), = bank_files(drive)
    assert set(report_summary.bank_from_bytes(data)) == {"a|1", "b|1"}            # both copies' reports, one file
    assert report_summary.update_bank(drive, "folder", "IM", {}) is None          # one file left: nothing written


def test_only_the_old_per_report_summaries_go_to_the_bin():
    drive = Drive()
    keep = [DriveFile("x", XLSX), DriveFile("b", report_summary.bank_name("IM"))]
    old = [DriveFile("s1", XLSX.replace(".xlsx", ".summary.json")), DriveFile("s2", "BK_recon_summary_x.summary.json")]
    drive.files = {f.id: (f.name, b"", "folder") for f in keep + old}
    assert report_summary.trash_old(drive, keep + old) == [f.name for f in old]
    assert sorted(drive.files) == ["b", "x"]                                      # the report and the bank file stay
    drive.refuse_trash = True
    assert report_summary.trash_old(drive, [DriveFile("b2", "IM_x.summary.json")]) == []   # refused: left, no error


class TwoPeople(Drive):
    """One Drive seen by two accounts: each one changes and bins only the files it made."""
    def __init__(self):
        super().__init__()
        self.me, self.owner = "a@irembo.com", {}

    def account_email(self):
        return self.me

    def list_children(self, parent, name=None, folders_only=False):
        return [DriveFile(f.id, f.name, mine=self.owner.get(f.id) == self.me) for f in super().list_children(parent, name)]

    def upload(self, parent, name, data, mime, replace_id=None, http=None):
        if replace_id and self.owner.get(replace_id) != self.me:
            raise DriveError(f"{name}: not available to this app")
        stored = super().upload(parent, name, data, mime, replace_id, http)
        self.owner.setdefault(stored.id, self.me)
        return stored

    def trash(self, fid, what=""):
        if self.owner.get(fid) != self.me:
            raise DriveError("not allowed")
        self.files.pop(fid)


def test_only_the_keeper_writes_a_banks_file_the_other_bins_its_own_copies(monkeypatch):
    # 18/11/2026: BK kept by a@; b@'s app writes nothing there and bins its own copy; a@ merges into one
    monkeypatch.setattr(report_summary, "KEEPERS", {"IM": "a@irembo.com"})
    drive = TwoPeople()
    drive.me = "b@irembo.com"
    monkeypatch.setattr(report_summary, "KEEPERS", {})                   # before: b@ made a copy (no keepers yet)
    report_summary.update_bank(drive, "folder", "IM", {"b|1": facts()})
    monkeypatch.setattr(report_summary, "KEEPERS", {"IM": "a@irembo.com"})
    drive.me = "a@irembo.com"
    report_summary.update_bank(drive, "folder", "IM", {"a|1": facts()})  # a@: b@'s copy cannot be replaced
    assert len(bank_files(drive)) == 2
    keeper = [d for fid, d in bank_files(drive) if drive.owner[fid] == "a@irembo.com"][0]
    assert set(report_summary.bank_from_bytes(keeper)) == {"a|1", "b|1"}  # every copy's figures in the keeper's
    drive.me = "b@irembo.com"
    assert report_summary.update_bank(drive, "folder", "IM", {"c|1": facts()}) is None   # b@ writes nothing
    (fid, data), = bank_files(drive)                                       # and its own copy went to the bin
    assert drive.owner[fid] == "a@irembo.com" and "c|1" not in report_summary.bank_from_bytes(data)
    drive.me = "a@irembo.com"
    report_summary.update_bank(drive, "folder", "IM", {"c|1": facts()})   # the keeper adds it (dashboard, Excel)
    (fid, data), = bank_files(drive)
    assert set(report_summary.bank_from_bytes(data)) == {"a|1", "b|1", "c|1"}


def test_an_unknown_account_writes_no_kept_bank_file(monkeypatch):
    monkeypatch.setattr(report_summary, "KEEPERS", {"IM": "a@irembo.com"})
    drive = Drive()                                                       # no account_email: not the keeper
    assert report_summary.update_bank(drive, "folder", "IM", {"a|1": facts()}) is None and not bank_files(drive)


def test_old_per_report_summaries_of_another_account_are_left():
    drive = Drive()
    drive.files = {"s1": ("IM_x.summary.json", b"", "folder")}
    assert report_summary.trash_old(drive, [DriveFile("s1", "IM_x.summary.json", mine=False)]) == []
    assert "s1" in drive.files
