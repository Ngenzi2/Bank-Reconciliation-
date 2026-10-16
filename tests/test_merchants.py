"""Payment per merchant (approved 2026-10-12): the 6th table of every summary block, for every bank - the
notifications received by merchant and currency (amounts never added across currencies), with "VS prev" - and
everything that shows a report (PDF, custom report, dashboard). Real samples are used when they are on this
computer (skipped otherwise)."""
import glob
import os
from pathlib import Path

import pandas as pd
import pytest
from reportlab import rl_config

import engine as e
import pdf_report
from services import dashboard

D = Path(os.path.expanduser("~/Downloads"))
IM_AUG = (sorted(glob.glob(str(D / "I&M Bank/2026-08/statements/*.xls"))),
          D / "_SELECT_pi_invoice_number_pi_transaction_id_pi_amount_pi_currenc_202610021501.csv")


def run_im_august(tmp_path, **kw):
    stm, pay = IM_AUG
    if not stm or not pay.exists():
        pytest.skip("real samples not on this computer")
    files = [e.read_bank_file(Path(p).name, Path(p).read_bytes(), "IM") for p in stm] + \
            [e.read_payment_file(pay.name, pay.read_bytes(), "IM")]
    period = e.make_period("month", "2026-08-01")
    acc = []
    for f in files:
        if e.accept_file(period, f, acc, "IM") is None:
            acc.append(f)
    return e.run(tmp_path / "base", period, acc, require_review=False, bank="IM", report_dir=tmp_path / "r", **kw)


def test_every_block_has_a_merchant_table_whose_all_rows_are_the_notifications_received(tmp_path):
    res = run_im_august(tmp_path)
    assert set(res.merchants) == {heading for _, heading, _ in res.blocks}
    for _, heading, tables in res.blocks:
        m = res.merchants[heading]
        received = tables["Overall notifications received by Irembo"]
        expected = {c: n for c, n in zip(received["Currency"], received["Number of Transactions"]) if pd.notna(c) and n}
        alls = m[m["Merchant name"] == "All"].set_index("Currency")["Number of Transactions"].to_dict()
        assert alls == expected                                         # one "All" per currency, never added up
        for cur, part in m[m["Merchant name"] != "All"].groupby("Currency"):
            assert part["Number of Transactions"].sum() == expected[cur]
            assert part["Number of Transactions"].is_monotonic_decreasing


def test_the_report_reads_back_its_merchant_tables_and_older_reports_still_open(tmp_path):
    res = run_im_august(tmp_path)
    back = e.read_report(res.report_file)
    assert len(back["blocks"]) == len(res.blocks) and set(back["merchants"]) == set(res.merchants)
    for heading, table in res.merchants.items():
        cols = ["Merchant name", "Currency", "Number of Transactions"]
        assert back["merchants"][heading][cols].reset_index(drop=True).equals(table[cols].reset_index(drop=True))
    for name, (_, _, tables) in zip(back["blocks"], res.blocks):         # the 5 tables are unchanged
        assert list(name[2]) == list(tables)
    old = D / "ECO_recon_summary_05-10-2026_to_11-10-2026.xlsx"        # made before the merchant table
    if old.exists():
        r = e.read_report(old)
        assert r["blocks"] and r["merchants"] == {}


def test_vs_prev_compares_each_merchant_of_the_same_currency():
    now = pd.DataFrame({"Merchant name": ["GOR", "RSSB", "NEW", "All"], "Currency": ["RWF"] * 4,
                        "Number of Transactions": [2279, 1278, 3, 3560], "Total Amount": [0.0] * 4, "%": [0.0] * 4})
    prev = {("GOR", "RWF"): 1939, ("RSSB", "RWF"): 2753, ("All", "RWF"): 4692}
    t = e.add_vs_merchants(now, prev, "Month")
    assert t["VS prev Month (%)"].tolist() == ["17.53%", "-53.58%", "N/A", "-24.13%"]
    assert e.add_vs_merchants(now, None, "Week")["VS prev Week (%)"].isna().all()   # week 1: nothing to compare
    assert (e.add_vs_merchants(now, {}, "Month")["VS prev Month (%)"] == "N/A").all()   # no table before


def test_report_pdf_shows_the_merchant_table(tmp_path, monkeypatch):
    monkeypatch.setattr(rl_config, "pageCompression", 0)
    res = run_im_august(tmp_path)
    text = pdf_report.report_pdf(res.report_file).decode("latin-1")
    assert "(Payment per merchant) Tj" in text and "(RSSB) Tj" in text


def test_dashboard_merchants_add_transactions_and_keep_currencies_apart():
    f = lambda m: dashboard.Fact("IM", "month", pd.Timestamp("2026-08-01"), pd.Timestamp("2026-08-31"), {}, {}, m)
    facts = [f({("GOR", "RWF"): (10, 100.0), ("RDB", "USD"): (5, 50.0)}), f({("GOR", "RWF"): (4, 40.0)}), f(None)]
    rwf, missing = dashboard.merchants(facts, [], "RWF")
    usd, _ = dashboard.merchants(facts, [], "USD")
    assert missing == 1
    assert rwf.to_dict("records")[0]["Merchant"] == "GOR" and rwf["Transactions"].iloc[0] == 14
    assert rwf["Amount"].iloc[0] == 140.0 and rwf["Share %"].iloc[0] == 100.0
    assert list(usd["Merchant"]) == ["RDB"]                      # RWF and USD are never mixed


def test_custom_report_has_a_merchant_section():
    assert "merchants" in e.CUSTOM_SECTIONS and e.CUSTOM_SECTIONS["merchants"][2] == [e.TABLE, e.BAR]
    assert "merchants" in e.REPORT_TYPES["Full Report"][1]
