"""Send to Slack (Results): the team message matching the notebooks, the automatic-delivery block format agreed
on 2026-10-15, and the webhook kept as a secret. Nothing is sent to Slack: the web call is faked."""
import io
import json
import os
import stat
import urllib.error

import numpy as np
import pandas as pd
import pytest

import engine as e
from integrations import slack
from services import slack_message as M

TITLES = [M.P, M.R, M.S, M.PE, M.DU]


def table(rwf, usd, amount=1000.0):
    return pd.DataFrame({"Currency": ["RWF", "USD", np.nan], "Number of Transactions": [rwf, usd, rwf + usd],
                         "Total Amount": [rwf * amount, usd * 10.0, np.nan]})


def block(proc=1824, rec=1824, sla=0, pend=0, dup=19):
    return {M.P: table(proc, 0), M.R: table(rec, 0), M.S: table(sla, 0), M.PE: table(pend, 0), M.DU: table(dup, 0)}


MERCH = pd.DataFrame({"Merchant name": ["RSSB", "GOR", "All"], "Currency": ["RWF"] * 3,
                      "Number of Transactions": [1000, 824, 1824], "Total Amount": [0.0] * 3, "%": [0.5, 0.5, 1.0]})
WEEK = e.make_period("week", "2026-08-17")
HEAD = "Week 1: 17/08/2026 - 23/08/2026"


def msg(previous=None, link=None, pending_statuses=None, **kw):
    return M.build(e.get_bank("ECO"), WEEK, [("Summary-Weekly", HEAD, block(**kw))], {HEAD: MERCH},
                   pending_statuses, previous, link)


def test_message_with_comparison_follows_the_notebook_format():
    prev = {M.P: {"Total": 1376}, M.R: {"Total": 1376}, M.S: {"Total": 0}, M.PE: {"Total": 0}, M.DU: {"Total": 0}}
    text = msg(prev, link="https://drive.google.com/file/d/ID/view")
    assert text.startswith("Hello Team, here is the weekly report on "
                           "*Ecobank payments* processed in `Aug-Week3` vs `Aug-Week2`")
    assert "• *Merchants who used this channel:* [GOR, RSSB]" in text
    assert "Total number of Transactions: `1,824` \U0001F4C8 +32.56% vs Prev-Week" in text
    assert "`100.00% Success rate` \U0001F534 with a failure rate of `0.00%`" in text
    assert "• *Pending notifications:* ⚠️" in text
    assert "Total number of transactions: `0` `0.00% OF OVERALL` N/A vs Prev-Week" in text
    assert "• *Paid with other channels, which results in double payments:* ⚠️" in text
    assert "Total number of transactions: `19` `1.04% OF OVERALL` N/A vs Prev-Week" in text
    assert "Total Amount: 19,000 RWF" in text and "USD" not in text.split("other channels")[1]
    assert "which is below `10%`" in text
    assert "Full details in the Google Sheets report https://drive.google.com/file/d/ID/view" in text
    assert text.endswith("CC: ")


def test_message_without_comparison_shows_na():
    text = msg(None)
    assert "`Aug-Week3`" in text
    assert "`N/A` \U0001F4C8 in the processed transactions vs Prev-Week" in text
    assert "Full details in the Google Sheets report (link)" in text


def test_sla_above_10_percent_is_red():
    text = msg(None, sla=300)
    assert "• *SLA Violated:* \U0001F534" in text and "which is above `10%`" in text


def test_pending_by_status_is_listed_when_counted():
    text = msg(None, pend=2, pending_statuses={"NEW": 1, "EXPIRED": 1})
    assert "We have 2 pending transactions ⚠️" in text
    assert "Pending by status: NEW: 1, EXPIRED: 1" in text


@pytest.mark.skipif(os.name == "nt", reason="Windows has no POSIX permission bits to check")
def test_webhook_is_kept_as_a_secret(tmp_path, monkeypatch):
    path = tmp_path / "slack-webhook.txt"
    monkeypatch.setenv("BANK_RECON_SLACK_WEBHOOK_FILE", str(path))
    with pytest.raises(slack.SlackError):
        slack.save_webhook("https://example.com/not-slack")
    url = "https://hooks.slack.com/services/TAAAAAAAA/BBBBBBBBBBB/abcdefghijklmnopqrstuvwx"
    slack.save_webhook(url)
    assert slack.load_webhook() == url and stat.S_IMODE(os.stat(path).st_mode) == 0o600
    assert slack.masked(url) == "hooks.slack.com/services/TAA…/…vwx" and "abcdefgh" not in slack.masked(url)


def test_send_posts_the_text_and_reports_refusals(monkeypatch):
    posted = []

    class R(io.BytesIO):
        status = 200
        def __enter__(self): return self
        def __exit__(self, *a): return False

    def ok(req, timeout=None):
        posted.append(json.loads(req.data))
        return R(b"ok")
    monkeypatch.setattr(slack.urllib.request, "urlopen", ok)
    slack.send("hello", url="https://hooks.slack.com/services/T/B/x")
    assert posted == [{"text": "hello"}]

    def gone(req, timeout=None):
        raise urllib.error.HTTPError(req.full_url, 404, "no_service", {}, io.BytesIO(b"no_service"))
    monkeypatch.setattr(slack.urllib.request, "urlopen", gone)
    with pytest.raises(slack.SlackError, match="paste the new one"):
        slack.send("hello", url="https://hooks.slack.com/services/T/B/x")


# ---- the automatic-delivery message (Slack blocks, approved on 2026-10-15) ------------------------------------------
MONTH = e.make_period("month", "2026-08-01")
MHEAD = "Month: August 2026 (01/08/2026 - 31/08/2026)"


def blocks_msg(native=True, link={"xlsx": "https://drive.google.com/file/d/X/view",
                                   "pdf": "https://drive.google.com/file/d/P/view"}, **kw):
    return M.build_blocks(e.get_bank("IM"), MONTH, [("Summary-Weekly", HEAD, block()), ("Summary-Monthly", MHEAD, block(**kw))],
                          {MHEAD: MERCH}, link, pd.Timestamp("2026-10-12 10:42"), native)


def test_step6_message_reads_the_report_and_links_it():
    p = blocks_msg(proc=5419, rec=5409, sla=19, pend=1, dup=9)
    heads = [b["text"]["text"] for b in p["blocks"] if b["type"] == "header"]
    assert heads == ["📊 I&M Bank — Monthly Reconciliation | August 2026", "Overall performance",
                     "Reconciliation summary", "🏢 Top merchants by RWF notifications"]
    text = json.dumps(p, ensure_ascii=False)
    assert "Period: 1–31 August 2026 · Report date: 12 October 2026" in text
    assert "On target — follow-up required" in text and "Received: `99.82%` · Target ≥ 99%" in text
    assert "On time: `99.65%` · Target ≥ 90%" in text                     # the rates in Slack's code style
    table_rows = next(b for b in p["blocks"] if b["type"] == "table")["rows"]
    assert [c["text"] for c in table_rows[0]] == ["Category", "Transactions", "% of processed", "RWF"]   # no USD: 0 payments
    assert [c["text"] for c in table_rows[1]] == ["Payments processed", "5,419", "100%", "5,419,000"]
    assert [c["text"] for c in table_rows[3]] == ["Late (SLA violated)", "19", "0.35%", "19,000"]
    assert "Exception value (pending + other channel): RWF 10,000" in text
    assert "*1.* RSSB — 1,000 · 50.0%" in text and "All" not in text.split("Top merchants")[1]
    assert ("Explore the full report:* <https://drive.google.com/file/d/X/view|📊 Excel report> · "
            "<https://drive.google.com/file/d/P/view|📄 PDF report>") in text
    assert p["text"].startswith("I&M Bank · August 2026 · 5,419 payments")


def test_step6_message_status_and_text_table():
    assert "Off target" in json.dumps(blocks_msg(sla=300), ensure_ascii=False)                # on time below 90 %
    clean = json.dumps(blocks_msg(dup=0, pend=0), ensure_ascii=False)
    assert "*On target*" in clean and "follow-up" not in clean
    p = blocks_msg(native=False)
    assert not any(b["type"] == "table" for b in p["blocks"])
    assert "```Category" in json.dumps(p, ensure_ascii=False)
    assert "Payments processed" in M.preview_text(p)


def test_step6_message_is_sent_as_blocks(monkeypatch):
    sent = {}

    class Answer(io.BytesIO):
        status = 200
        def __enter__(self): return self
        def __exit__(self, *a): return False

    def fake(req, timeout=15):
        sent.update(json.loads(req.data.decode()))
        return Answer(b"ok")

    monkeypatch.setattr(slack.urllib.request, "urlopen", fake)
    slack.send_payload(blocks_msg(), "https://hooks.slack.com/services/T/B/x")
    assert sent["blocks"][0]["type"] == "header" and "5,419" not in sent["text"]
