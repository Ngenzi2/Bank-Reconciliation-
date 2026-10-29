"""Follow-up (22/10/2026): the pending and paid-twice items of the saved reports, decisions shared on Drive (here a
fake Drive), the figures and the Slack reminder."""
import json

import pandas as pd

from services import followup


def item(key, kind="Pending", amount=1000.0, cur="RWF", reported="2026-10-12", paid="2026-10-08", bank="IM", inv="886"):
    return {"key": key, "bank": bank, "type": kind, "invoice": inv, "reference": "", "amount": amount, "currency": cur,
            "paid_on": pd.Timestamp(paid), "merchant": "", "provider": "", "report": "r.xlsx",
            "reported": pd.Timestamp(reported)}


def test_summary_overdue_and_rate_after_follow_up():
    items = [item("a"), item("b", "Paid twice", 500), item("c", reported="2026-10-20"), item("d", cur="USD", amount=20)]
    dec = {"a": {"decision": "Received later"}, "b": {"decision": "Refunded"}}
    s = followup.summary(items, dec, today="2026-10-22", days=7)
    assert (s["total"], s["resolved"], s["open"]) == (4, 2, 2)
    assert s["overdue"] == 1 and s["oldest"] == 10                       # d: 10 days; c: 2 days
    assert s["open_amount"] == {"RWF": 1000.0, "USD": 20.0} and s["received_later"] == 1
    assert s["by_type"] == {"Pending": (2, 3), "Paid twice": (0, 1)}


def test_reminder_lists_the_overdue_items():
    items = [item("a", amount=40000, inv="886111"), item("c", reported="2026-10-20")]
    text = followup.reminder_text(items, {}, days=7, today="2026-10-22")
    assert text.startswith("⚠️ 1 item open for more than 7 days: RWF 40,000")
    assert "invoice 886111" in text and "open 10 days" in text
    assert followup.reminder_text(items, {"a": {"decision": "Refunded"}}, 7, "2026-10-22").startswith("✅")


class FakeDrive:
    def __init__(self):
        self.files = {}

    def list_children(self, parent, name=None, folders_only=False):
        class F:
            pass
        out = []
        for fid, (n, _) in self.files.items():
            f = F(); f.id, f.name = fid, n
            out.append(f)
        return out

    def download(self, fid, mime="", http=None):
        return self.files[fid][1]

    def upload(self, parent, name, data, mime, replace_id=None, http=None):
        fid = replace_id or f"id{len(self.files)}"
        self.files[fid] = (name, data)


def test_decisions_are_merged_with_the_latest_on_drive():
    drive = FakeDrive()
    followup.save_decisions(drive, "root", {"a": {"decision": "Refunded", "comment": "sent back"}}, "ann@irembo.com")
    followup.save_decisions(drive, "root", {"b": {"decision": "Received later", "comment": ""}}, "bob@irembo.com")
    dec, fid = followup.load_decisions(drive, "root")
    assert set(dec) == {"a", "b"} and dec["a"]["by"] == "ann@irembo.com" and dec["a"]["comment"] == "sent back"
    assert len(drive.files) == 1 and json.loads(drive.files[fid][1])["version"] == 1    # one shared file, replaced
    followup.save_decisions(drive, "root", {"a": {"decision": ""}}, "ann@irembo.com")    # back to open
    assert set(followup.load_decisions(drive, "root")[0]) == {"b"}


def test_latest_version_of_each_report_since_the_start():
    class R:
        def __init__(self, stem, v, start, ext="xlsx"):
            self.stem, self.version, self.start, self.ext, self.modified = stem, v, pd.Timestamp(start), ext, None
    recs = [R("IM_w1", 1, "2026-10-05"), R("IM_w1", 2, "2026-10-05"), R("IM_sep", 1, "2026-09-01"),
            R("IM_w1", 3, "2026-10-05", "pdf")]
    got = followup.latest_reports(recs, "2026-10-01")
    assert [(r.stem, r.version) for r in got] == [("IM_w1", 2)]
    # 28/10/2026: no start date (the period is chosen on the Follow-up page): every report, latest version
    every = followup.latest_reports(recs, None)
    assert sorted((r.stem, r.version) for r in every) == [("IM_sep", 1), ("IM_w1", 2)]


def test_items_in_the_period_of_some_banks():
    items = [item("a", paid="2026-10-08"), item("b", paid="2026-11-02"), item("c", bank="BOA")]
    assert [i["key"] for i in followup.in_period(items, "2026-10-01", "2026-10-31", ["IM"])] == ["a"]


def test_week_names_and_file_names():
    from pathlib import Path
    import engine
    from services import report_history
    p = engine.make_period("week", "2026-09-28")                 # 3 days in September, 4 in October
    name = engine.report_path(Path("."), p, "IM").name
    assert name == "IM_recon_summary_28-09-2026_to_04-10-2026 (week 1).xlsx"
    assert engine.week_name(p) == "Week 1 · October 2026"
    assert engine.week_number(pd.Timestamp("2026-09-28"), month="2026-09-01")[1] == 5   # in September's report
    assert report_history.parse_name(name)["label"] == p.label                         # still read by its dates
    assert report_history.parse_name("IM_recon_summary_28-09-2026_to_04-10-2026.xlsx")["label"] == p.label
    assert report_history.parse_name(name.replace(".xlsx", "_v2.pdf"))["version"] == 2
