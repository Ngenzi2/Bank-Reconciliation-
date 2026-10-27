"""The dashboard's figures kept on this computer between sessions (services/facts_cache.py), the reports downloaded
together (report_history.download_many) and the summary read without the large Payment details sheet."""
import pandas as pd

from services import dashboard, facts_cache, report_history


class Rec:
    def __init__(self, fid, modified="2026-10-01T10:00:00Z", name="IM_recon_summary_01-09-2026_to_30-09-2026.xlsx",
                mime=""):
        self.file_id, self.modified, self.name, self.mime = fid, modified, name, mime


def fact():
    return dashboard.Fact("IM", "month", pd.Timestamp("2026-09-01"), pd.Timestamp("2026-09-30"),
                          {"processed": {"RWF": 10}}, {"processed": {"RWF": 1000.0}}, record="drive record")


def test_kept_by_file_and_date_and_dropped_when_gone(tmp_path):
    path = tmp_path / "output" / ".dashboard_facts.pkl"
    a, b = Rec("A"), Rec("B")
    facts_cache.save(path, {facts_cache.key(a): [fact()], facts_cache.key(b): [fact()]})
    stored = facts_cache.load(path)
    assert set(stored) == {facts_cache.key(a), facts_cache.key(b)}
    assert stored[facts_cache.key(a)][0].record is None                      # stored without its Drive record
    assert facts_cache.attach(stored[facts_cache.key(a)], a)[0].record is a
    assert facts_cache.key(Rec("A", "2026-10-02T09:00:00Z")) not in stored   # a replaced report is read again
    facts_cache.save(path, {}, keep=[facts_cache.key(a)])                    # B no longer on Drive
    assert set(facts_cache.load(path)) == {facts_cache.key(a)}


def test_another_version_or_a_broken_file_is_ignored(tmp_path, monkeypatch):
    path = tmp_path / "f.pkl"
    facts_cache.save(path, {"k": [fact()]})
    monkeypatch.setattr(facts_cache, "VERSION", facts_cache.VERSION + 1)
    assert facts_cache.load(path) == {}
    path.write_bytes(b"not a pickle")
    assert facts_cache.load(path) == {}


def test_download_many_without_parallel_connections():
    class Client:
        def thread_http(self):
            return None

        def download(self, fid, mime="", http=None, as_xlsx=False):
            if fid == "bad":
                raise report_history.DriveError("not found")
            return fid.encode()
    done = []
    out = report_history.download_many(Client(), [Rec("x"), Rec("bad")], done=done.append)
    assert out["x"] == b"x" and isinstance(out["bad"], report_history.DriveError) and done == [1, 2]


def test_download_many_with_an_older_client_goes_one_by_one():
    class OldClient:                                   # download() without http, as in a running app's old client
        def thread_http(self):
            return object()

        def download(self, fid, mime="", as_xlsx=False):
            return fid.encode()
    out = report_history.download_many(OldClient(), [Rec("x"), Rec("y")])
    assert out == {"x": b"x", "y": b"y"}
