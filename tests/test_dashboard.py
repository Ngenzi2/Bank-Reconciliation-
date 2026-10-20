"""Reports Dashboard (manager view): every figure comes from the saved reports' tables.
Facts are built from small report dicts shaped like engine.read_report() and checked by hand."""
import io
import os
from pathlib import Path
from types import SimpleNamespace

import pandas as pd
import pytest
from openpyxl import Workbook

import config
import engine as e
from boa_fixtures import payments
from services import auto_review, dashboard as d, folder_files


def table(counts, amounts):
    """A summary table: one row per currency (RWF, USD) + the Total row."""
    return pd.DataFrame({"Currency": ["RWF", "USD", None], "Number of Transactions": [*counts, sum(counts)],
                         "Total Amount": [*amounts, None]})


def tables(proc=(1000, 10), amt=(5_000_000.0, 900.0), received=(990, 10), sla=(5, 0), pending=(2, 0), dup=(1, 0)):
    a = lambda n: (n[0] * 1000.0, n[1] * 10.0)
    return {d.TITLES["processed"]: table(proc, amt), d.TITLES["received"]: table(received, a(received)),
            d.TITLES["sla"]: table(sla, a(sla)), d.TITLES["pending"]: table(pending, a(pending)),
            d.TITLES["duplicate"]: table(dup, a(dup))}


def record(bank="IM", kind="month", start="2026-08-01", version=1):
    return SimpleNamespace(bank=bank, kind=kind, start=pd.Timestamp(start), version=version, name=f"{bank}-{start}",
                           ext="xlsx", modified=pd.Timestamp("2026-10-01", tz="UTC"))


def month_report(bank="IM", month="2026-08", weeks=((1, 2), (3, 9)), **kw):
    """A monthly report: the month block + week blocks (day ranges of that month)."""
    first = pd.Timestamp(f"{month}-01")
    last = first + pd.offsets.MonthEnd(0)
    blocks = [("Summary-Weekly", f"Week {i}: {first.replace(day=a):%d/%m/%Y} - {first.replace(day=b):%d/%m/%Y}",
               tables(**kw)) for i, (a, b) in enumerate(weeks, 1)]
    blocks.append(("Summary-Monthly", f"Month: {first:%B %Y} ({first:%d/%m/%Y} - {last:%d/%m/%Y})", tables(**kw)))
    return {"blocks": blocks, "merchants": {}}, record(bank, "month", f"{month}-01")


def test_records_keep_the_first_version_of_monthly_and_weekly_reports_only():
    recs = [record("IM", "month", "2026-08-01", 2), record("IM", "month", "2026-08-01", 1),
            record("IM", "week", "2026-08-03"), record("IM", "day", "2026-08-03")]
    kept = d.report_records(recs)
    assert [(r.kind, r.version) for r in kept] == [("week", 1), ("month", 1)]


def test_a_monthly_report_gives_its_month_and_its_weeks():
    rep, rec = month_report()
    facts = d.facts_from_report(rep, rec, weekend={"2026-08-03": {"RWF": 2}})
    assert [(f.kind, f"{f.start:%d/%m}", f"{f.end:%d/%m}") for f in facts] == [
        ("week", "01/08", "02/08"), ("week", "03/08", "09/08"), ("month", "01/08", "31/08")]
    assert facts[-1].counts["processed"] == {"RWF": 1000, "USD": 10}
    assert facts[1].weekend == {"RWF": 2} and facts[0].weekend == {}           # 03/08 is in the second week


def test_a_weekly_report_replaces_the_months_blocks_of_its_week():
    rep, rec = month_report(weeks=((1, 2), (3, 9)))
    month_facts = d.facts_from_report(rep, rec)
    week = d.facts_from_report({"blocks": [("Summary-Weekly", "Week 1: 03/08/2026 - 09/08/2026",
                                            tables(proc=(70, 0)))], "merchants": {}}, record("IM", "week", "2026-08-03"))
    used = d.combine(month_facts + week)
    weeks = [f for f in used if f.kind == "week"]
    assert [(f"{f.start:%d/%m}", f.source) for f in weeks] == [("01/08", "month"), ("03/08", "week")]
    assert d.totals(d.select(used, "month", pd.Timestamp("2026-08-01")))["processed"] == 1010   # the month as saved


def test_a_week_cut_by_a_month_end_adds_its_two_parts():
    aug, r1 = month_report("IM", "2026-08", weeks=((31, 31),), proc=(5, 0))
    sep, r2 = month_report("IM", "2026-09", weeks=((1, 6),), proc=(7, 0))
    facts = d.combine(d.facts_from_report(aug, r1) + d.facts_from_report(sep, r2))
    assert d.totals(d.select(facts, "week", pd.Timestamp("2026-08-31")))["processed"] == 12


def test_totals_rates_and_currency_filter():
    rep, rec = month_report()
    facts = d.select(d.facts_from_report(rep, rec), "month", pd.Timestamp("2026-08-01"))
    t, rwf = d.totals(facts), d.totals(facts, "RWF")
    assert t["processed"] == 1010 and rwf["processed"] == 1000
    assert t["amounts"]["RWF"]["processed"] == 5_000_000.0 and t["amounts"]["USD"]["processed"] == 900.0
    r = d.rates(t)
    assert r["receipt"] == pytest.approx(1000 / 1010) and r["ontime"] == pytest.approx(1005 / 1010)


def test_status_against_the_targets():
    good = d.totals(d.facts_from_report(*month_report(pending=(0, 0), dup=(0, 0)))[-1:])
    warn = d.totals(d.facts_from_report(*month_report())[-1:])
    late = d.totals(d.facts_from_report(*month_report(sla=(500, 0)))[-1:])
    assert (d.status(good), d.status(warn), d.status(late), d.status(None)) == ("good", "warn", "bad", "none")


def test_no_transactions_has_no_rate_not_zero():
    t = d.totals(d.facts_from_report(*month_report(proc=(0, 0), received=(0, 0), sla=(0, 0), pending=(0, 0),
                                                   dup=(0, 0)))[-1:])
    assert d.rates(t)["receipt"] is None and d.tiles(t, None, {"IM": t}, "")[3][1] == "N/A"


def test_attention_puts_a_missed_target_first_and_names_missing_reports():
    late = d.totals(d.facts_from_report(*month_report("BOA", sla=(500, 0)))[-1:])
    warn = d.totals(d.facts_from_report(*month_report("IM"))[-1:])
    items = d.attention({"IM": warn, "ECO": None, "BOA": late}, {}, "August 2026")
    assert items[0]["bank"] == "BOA" and items[0]["kind"] == "bad"
    assert "50.5% on time" in items[0]["title"]                                # (1010 - 500) / 1010
    assert any(i["bank"] == "ECO" and i["kind"] == "info" for i in items)


def test_previous_period_and_change():
    assert d.previous_period("month", pd.Timestamp("2026-09-01"), None, "previous") == (
        pd.Timestamp("2026-08-01"), pd.Timestamp("2026-08-01"))
    assert d.previous_period("week", pd.Timestamp("2026-09-07"), pd.Timestamp("2026-09-14"), "previous") == (
        pd.Timestamp("2026-08-24"), pd.Timestamp("2026-08-31"))
    assert d.previous_period("quarter", pd.Timestamp("2026-07-01"), None, "year")[0] == pd.Timestamp("2025-07-01")
    assert d.change(80, 100) == pytest.approx(-0.2) and d.change(5, 0) is None


def test_waterfall_goes_from_before_to_now_bank_by_bank():
    now = {"IM": {"processed": 80, "amounts": {}}, "ECO": None, "BOA": {"processed": 5, "amounts": {}}}
    before = {"IM": {"processed": 100, "amounts": {}}, "ECO": None, "BOA": {"processed": 10, "amounts": {}}}
    steps = d.waterfall(now, before)
    assert steps[0] == ("Before", 0, 110, "total") and steps[-1] == ("Now", 0, 85, "total")
    assert [s[3] for s in steps[1:-1]] == ["down", "none", "down"]


def test_what_if_weekend_moves_late_payments_to_on_time():
    t = {"processed": 66, "sla": 45, "weekend": 45}
    w = d.what_if_weekend(t, {"processed": 4393, "sla": 59}, 1.0)
    assert w["moved"] == 45 and w["bank_after"] == 1.0 and round(w["all_after"], 4) == round(4379 / 4393, 4)


def test_weekend_late_counts_fri_to_sun_payments_booked_monday():
    details = pd.DataFrame({"sla_status": ["SLA violated", "SLA violated", "SLA violated", "On time"],
                            "Value Date": ["2026-09-07", "2026-09-07", "2026-09-08", "2026-09-07"],
                            "paid_at": ["2026-09-04 22:03", "2026-09-05 11:00", "2026-09-05 11:00", "2026-09-05"],
                            "Currency": ["RWF"] * 4})
    assert d.weekend_late(details, "BOA") == {"2026-09-07": {"RWF": 2}}       # Tuesday's row is not a weekend


def test_merchants_compare_only_banks_with_merchant_data():
    a = d.Fact("IM", "month", pd.Timestamp("2026-08-01"), pd.Timestamp("2026-08-31"), {}, {},
               {("RSSB", "RWF"): (10, 100.0), ("RDB", "USD"): (3, 30.0)})
    b = d.Fact("IM", "month", pd.Timestamp("2026-09-01"), pd.Timestamp("2026-09-30"), {}, {}, None)
    empty, missing = d.merchants([b], [a], "RWF")
    assert len(empty) == 0 and missing == 1                                     # no data now: nothing "dropped"
    table, _ = d.merchants([a], [], "RWF")
    assert table.to_dict("records")[0]["Merchant"] == "RSSB" and table["Share %"].iloc[0] == 100.0


# ---- the hierarchy: year > quarter > month > week > day, all from the monthly reports -----------------------------
def _month_with_days():
    rep, rec = month_report(proc=(3, 0), received=(2, 0), sla=(1, 0), pending=(1, 0), dup=(0, 0), weeks=((1, 2), (3, 9)))
    month = [f for f in d.facts_from_report(rep, rec) if f.kind == "month"][0]
    details = pd.DataFrame({"Value Date": ["2026-08-01", "2026-08-03", "2026-08-03"], "Currency": ["RWF"] * 3,
                            "Credit": [100, 200, 300], "notification_received": [True, True, False],
                            "sla_status": ["On time", "SLA violated", "On time"]})
    pending = details.iloc[[2]]
    return month, details, pending, details.iloc[0:0]


def test_days_come_from_the_rows_and_must_add_up_to_the_month():
    month, details, pending, dup = _month_with_days()
    days = d.daily_facts(month, details, pending, dup, "IM")
    assert [(f"{f.start:%d/%m}", f.counts["processed"]["RWF"], f.counts["sla"]["RWF"]) for f in days] == [
        ("01/08", 1, 0), ("03/08", 2, 1)]
    assert days[1].amounts["processed"]["RWF"] == 500.0
    assert d.daily_facts(month, details.iloc[:2], pending, dup, "IM") == []      # rows missing: no days, never wrong


def test_each_level_shows_the_level_below_and_partial_periods_count_months():
    month, details, pending, dup = _month_with_days()
    facts = d.combine([month] + [f for f in d.facts_from_report(*month_report(proc=(3, 0), received=(2, 0), sla=(1, 0),
                                                                              pending=(1, 0), dup=(0, 0),
                                                                              weeks=((1, 2), (3, 9)))) if f.kind == "week"]
                      + d.daily_facts(month, details, pending, dup, "IM"))
    q = d.breakdown(facts, "quarter", pd.Timestamp("2026-07-01"), None, "All")
    assert list(q["Period"]) == ["August 2026"]
    w = d.breakdown(facts, "month", pd.Timestamp("2026-08-01"), None, "All")
    assert list(w["Period"]) == ["Week of 27/07/2026 (part in this period)", "Week of 03/08/2026"]
    assert d.CHILD["week"] == "day" and d.CHILD["day"] is None
    assert d.partial(facts, "quarter", pd.Timestamp("2026-07-01")) == [("IM", pd.Timestamp("2026-07-01"), 1, 3, "months")]
    assert d.totals(d.select(facts, "day", pd.Timestamp("2026-08-03")))["processed"] == 2


def test_complete_periods_and_the_reports_to_make():
    aug = d.facts_from_report(*month_report("IM", "2026-08", weeks=((1, 2), (3, 9))))
    sep = d.facts_from_report(*month_report("IM", "2026-09", weeks=((1, 6),)))
    facts = d.combine(aug + sep)
    assert d.complete_options(facts, "month") == [pd.Timestamp("2026-09-01"), pd.Timestamp("2026-08-01")]
    assert d.complete_options(facts, "quarter") == []                       # Q3 has no July
    assert pd.Timestamp("2026-07-27") not in d.complete_options(facts, "week")   # 27/07 - 02/08: July not saved
    todo = d.reports_to_make(facts, today="2026-10-14")
    assert ("ECO", pd.Timestamp("2026-08-01")) in todo and ("IM", pd.Timestamp("2026-09-01")) not in todo
    assert d.last_ended_month("2026-10-14") == pd.Timestamp("2026-09-01")
    assert "Aug 2026 – Sep 2026 saved" in d.coverage_line(facts, today="2026-10-14")


# ---- dashboard totals vs. the report's own tables (real engine.run(), not hand-built dicts) ------------------------
ECO_COLUMNS = ["Source Channel", "Biller Code", "Bill Ref No", "Ccy", "Amount", "Request Date", "Cba Rsp Code",
               "Cba Rsp Msg", "Biller Rsp Code", "Biller Rsp Msg", "Cba Ref No", "Customer Ref No"]


def ecobank_statement(rows):
    """rows: [(Bill Ref No, Ccy, Amount, Request Date UTC 'YYYY-MM-DD HH:MM:SS', Cba Ref No)], all SUCCESS."""
    wb = Workbook()
    ws = wb.active
    ws.append(ECO_COLUMNS)
    for ref, ccy, amount, request_date, cba_ref in rows:
        ws.append(["AGENCYBANKING", "B1", ref, ccy, amount, request_date, "000", "SUCCESS", "000", "SUCCESS",
                  cba_ref, ""])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def test_dashboard_totals_match_the_report_for_one_ecobank_week(tmp_path):
    period = e.make_period("week", "2026-09-07")
    rows = [
        ("880000000701", "RWF", 1000, "2026-09-08 08:00:00", "C701"),    # -> Kigali 10:00, paid 10:00:30: on time
        ("880000000702", "RWF", 2000, "2026-09-09 09:00:00", "C702"),    # -> Kigali 11:00, paid 11:00:10: on time
    ]
    pays = [
        dict(invoice="880000000701", amount=1000, paid_at="2026-09-08 10:00:30", ref="C701", provider="ECOBANK"),
        dict(invoice="880000000702", amount=2000, paid_at="2026-09-09 11:00:10", ref="C702", provider="ECOBANK"),
    ]
    files = [e.read_bank_file("eco.xlsx", ecobank_statement(rows), "ECO"),
            e.read_payment_file("p.csv", payments(pays), "ECO")]
    res = e.run(tmp_path / "base", period, files, require_review=False, bank="ECO", report_dir=tmp_path / "r")
    report_totals = e.counts_of(next(t for s, _, t in res.blocks if s == "Summary-Weekly"))

    rep = e.read_report(res.report_file)
    facts = d.facts_from_report(rep, record("ECO", "week", f"{period.start:%Y-%m-%d}"))
    totals = d.totals(facts)

    assert totals["processed"] == report_totals["Overall processed payments"]["Total"] == 2
    assert totals["received"] == report_totals["Overall notifications received by Irembo"]["Total"]
    assert totals["sla"] == report_totals["SLA Violated"]["Total"]
    assert totals["pending"] == report_totals["Pending notifications"]["Total"]
    assert totals["duplicate"] == report_totals["Duplicate Payment"]["Total"]


def test_dashboard_totals_match_the_known_bpr_august_report(tmp_path):
    """The same agreed numbers as tests/test_bpr.py's real-sample test: 537/512/13/25/0."""
    samples = os.environ.get("BANK_RECON_SAMPLES")
    if not samples:
        pytest.skip("BANK_RECON_SAMPLES not set")
    data = Path(samples)
    folders = {"statements": config.FolderSpec("local", str(data / "Bpr-irembo")),
              "exports": config.FolderSpec("local", str(data / "Bpr"))}
    if not (data / "Bpr-irembo").is_dir() or not (data / "Bpr").is_dir():
        pytest.skip("real samples not found under BANK_RECON_SAMPLES")
    period = e.make_period("month", "2026-08-01")
    found = folder_files.find_files(period, "BPR", {"BPR": folders})
    acc = folder_files.all_files(found)
    first = e.run(tmp_path / "base", period, acc, require_review=True, bank="BPR", report_dir=tmp_path / "r1")
    answers = {inv: "Open, not paid yet" for inv in
              ["876239124919", "886233705544", "886236247937", "886237258713", "886237993535"]}
    filled, _ = auto_review.classify(first.review, answers)
    e.save_decisions(tmp_path / "base", period, filled, "BPR")
    res = e.run(tmp_path / "base", period, acc, require_review=True, bank="BPR", report_dir=tmp_path / "r2")
    assert not res.stopped_for_review

    rep = e.read_report(res.report_file)
    facts = d.facts_from_report(rep, record("BPR", "month", "2026-08-01"))
    totals = d.totals([f for f in facts if f.kind == "month"])
    assert (totals["processed"], totals["received"], totals["sla"], totals["pending"], totals["duplicate"]) == \
        (537, 512, 13, 25, 0)
