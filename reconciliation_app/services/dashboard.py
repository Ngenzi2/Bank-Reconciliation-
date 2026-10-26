"""
REPORTS DASHBOARD (manager view): EVERY FIGURE FROM THE SAVED REPORTS
=====================================================================
The dashboard only reads the numbers already saved in the reports on Google Drive; no reconciliation rule is run
again and nothing is typed by hand: every figure, sentence, colour and ranking below is computed from them.

    Facts       one per summary block of a saved report, for one bank:
                  a monthly report  -> the month + each of its weeks (Summary-Weekly blocks)
                  a weekly report   -> its week; it replaces the week blocks of a monthly report for that week
                Of a report saved more than once (v1, v2, ...) the FIRST version is used.
                A fact holds, per currency, the count and the amount of the 5 tables, the Payment per merchant
                table (None for reports made before it existed) and the late payments made over the weekend and
                booked on Monday (read from the report's Payment details, for the what-if panel).
    Periods     week (Monday to Sunday: the two parts of a week cut by a month end are added together),
                month, quarter, year (months added), or a custom range of months.
    Totals      counts added up; amounts added per currency only (RWF and USD are never mixed); a rate is the
                added counts divided again: received / processed, on time = (processed - late) / processed.
    Status      against the targets (the same as the Slack message): received >= 99 %, on time >= 90 %
                -> "bad" when a target is missed; pending or paid with another channel -> "warn"; else "good".

    report_records() / facts_from_report() / weekend_late()   reading
    period_options() / select() / previous_period()            choosing what is shown
    totals() / rates() / status() / by_bank()                   figures
    tiles() / attention() / scorecard() / trend() / waterfall() / merchants() / what_if() / coverage()
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field

import pandas as pd

import engine

TARGET_RECEIPT, TARGET_ONTIME = 0.99, 0.90          # the defaults; the app sets the targets saved in Settings


def pct_target(v: float) -> str:
    """A target as people write it: 0.99 -> '99%', 0.995 -> '99.5%'."""
    return f"{v * 100:g}%"
TITLES = {"processed": "Overall processed payments", "received": "Overall notifications received by Irembo",
          "sla": "SLA Violated", "pending": "Pending notifications", "duplicate": "Duplicate Payment"}
METRICS = list(TITLES)
KINDS = ("year", "quarter", "month", "week", "day")    # the hierarchy: a year holds quarters, ..., a week holds days
CHILD = {"year": "quarter", "quarter": "month", "month": "week", "week": "day", "day": None}
PARENT = {v: k for k, v in CHILD.items() if v}


@dataclass
class Fact:
    bank: str                                         # bank code ('IM')
    kind: str                                         # "week" or "month"
    start: pd.Timestamp                               # first day
    end: pd.Timestamp                                 # last day
    counts: dict                                      # {metric: {currency: n}}
    amounts: dict                                     # {metric: {currency: amount}}
    merchants: dict | None = None                     # {(merchant, currency): (n, amount)}; None = no table
    weekend: dict = field(default_factory=dict)       # {currency: late payments made Fri-Sun, booked Monday}
    weekend_known: bool = False                       # False: the report's details could not be read
    record: object = None                             # the report file it comes from (report_history.Record)
    source: str = "month"                             # the kind of report it was read from


# =====================================================================================================================
# READING
# =====================================================================================================================
def report_records(records: list) -> list:
    """The Excel reports the dashboard uses: monthly and weekly reports, one per bank and period - the first version
    (at the same version, the first saved). Daily reports are left out (their days are in the weeks)."""
    best = {}
    for r in records:
        if r.ext != "xlsx" or r.kind not in ("month", "week"):
            continue
        key = (r.bank, r.kind, pd.Timestamp(r.start).normalize())
        saved = r.modified if r.modified is not None and pd.notna(r.modified) else pd.Timestamp.max.tz_localize("UTC")
        rank = (r.version, saved)
        if key not in best or rank < best[key][0]:
            best[key] = (rank, r)
    return sorted((r for _, r in best.values()), key=lambda r: (r.start, r.bank), reverse=True)


_DATES = re.compile(r"(\d{2}/\d{2}/\d{4})\s*-\s*(\d{2}/\d{2}/\d{4})")


def _block_dates(heading: str):
    """('Week 2: 07/09/2026 - 13/09/2026' or 'Month: September 2026 (01/09/2026 - 30/09/2026)') -> (first, last)."""
    m = _DATES.search(heading or "")
    if not m:
        return None
    return tuple(pd.to_datetime(x, format="%d/%m/%Y") for x in m.groups())


def _tables(tables: dict):
    """({metric: {currency: n}}, {metric: {currency: amount}}) of the 5 summary tables (the Total row left out)."""
    counts, amounts = {}, {}
    for key, title in TITLES.items():
        t = tables[title]
        counts[key], amounts[key] = {}, {}
        for c, n, a in zip(t["Currency"], t["Number of Transactions"], t["Total Amount"]):
            if pd.isna(c):
                continue
            counts[key][str(c)] = int(n)
            amounts[key][str(c)] = 0.0 if pd.isna(a) else float(a)
    return counts, amounts


def _merchants(table):
    if table is None:
        return None
    part = table[table["Merchant name"].astype(str).ne("All")]
    return {(str(m), str(c)): (int(n), 0.0 if pd.isna(a) else float(a))
            for m, c, n, a in zip(part["Merchant name"], part["Currency"], part["Number of Transactions"],
                                  part["Total Amount"])}


def facts_from_report(rep: dict, record, weekend: dict | None = None) -> list:
    """The facts of one report read by engine.read_report: its month block and its week blocks (a monthly report),
    or its week (a weekly report). weekend = weekend_late() of the same file: {day: {currency: n}}."""
    out = []
    for sheet, heading, tables in rep["blocks"]:
        kind = {"Summary-Monthly": "month", "Summary-Weekly": "week"}.get(sheet)
        dates = _block_dates(heading)
        if kind is None or dates is None:
            continue
        counts, amounts = _tables(tables)
        f = Fact(record.bank, kind, dates[0], dates[1], counts, amounts,
                 _merchants((rep.get("merchants") or {}).get(heading)), record=record, source=record.kind)
        if weekend is not None:
            f.weekend_known = True
            for day, per in weekend.items():
                if f.start <= pd.Timestamp(day) <= f.end:
                    for c, n in per.items():
                        f.weekend[c] = f.weekend.get(c, 0) + n
        out.append(f)
    return out


def no_transaction_fact(path, record) -> Fact | None:
    """The fact of a No-Transaction Report (step 1, "No-Transaction Report": one table, every count 0, the status
    "No Transactions Recorded"), or None when the file is a standard report. The counts are read from its table."""
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb[wb.sheetnames[0]]
        if not any(row and row[0] == engine.NO_TX_STATUS for row in ws.iter_rows(max_col=1, values_only=True)):
            return None
    finally:
        wb.close()
    bank = engine.get_bank(record.bank)
    kind = "week" if record.kind == "week" else "month"
    start = pd.Timestamp(record.start).normalize()
    end = period_end(start, kind) if kind == "month" else start + pd.Timedelta(days=6)
    counts = {m: {c: 0 for c in bank.CURRENCIES} for m in METRICS}
    amounts = {m: {c: 0.0 for c in bank.CURRENCIES} for m in METRICS}
    return Fact(record.bank, kind, start, end, counts, amounts, None, {}, True, record, record.kind)


def daily_facts(month: Fact, details: pd.DataFrame, pending: pd.DataFrame, duplicate: pd.DataFrame, bank,
                weekend: dict | None = None) -> list:
    """The days of a monthly or weekly report, from its detail sheets: processed = its Payment details rows (each
    invoice once; Ecobank: SUCCESS requests), received = those
    with notification_received, late = SLA violated, pending / paid with another channel = the rows of those sheets;
    by the bank date, per currency. They are kept only when they add up exactly to the month's own tables (a bank
    whose tables follow other rules, or an older report, has no days rather than wrong ones)."""
    bank = engine.get_bank(bank)
    day_col, amt_col = bank.COLUMNS["payment_date"], bank.COLUMNS["amount"]
    if details is None or not {day_col, "Currency"} <= set(details.columns):
        return []
    truthy = lambda col: details[col].astype(str).str.lower().isin(["true", "1", "1.0"]) if col in details else None
    # processed (26/10/2026): every credit, each invoice once (not the extra credits); Ecobank: SUCCESS requests only
    keep = pd.Series(True, index=details.index)
    status = details["review_status"].fillna("").astype(str) if "review_status" in details else pd.Series("", index=details.index)
    keep &= ~status.str.startswith("Excluded")              # 26/10/2026: an excluded row stays, in no table
    extra = truthy("is_extra_credit") if "is_extra_credit" in details else pd.Series(False, index=details.index)
    if bank.CODE == "ECO" and "Cba Rsp Code" in details:
        keep &= details["Cba Rsp Code"].map(bank._code).eq("000")
    done = details[keep & ~extra]
    # an extra credit counts in Received when the review put it there (26/10/2026); Pending / Duplicate: their sheets
    received_rows = (details[(keep & ~extra & truthy("notification_received")) | (keep & extra & status.eq("Notification received"))]
                     if "notification_received" in details else None)
    if received_rows is not None and "invoice_number" in received_rows:          # each invoice once
        received_rows = received_rows[~(received_rows["invoice_number"].notna()
                                        & received_rows["invoice_number"].duplicated())]
    parts = {"processed": done, "received": received_rows,
             "sla": done[done["sla_status"].astype(str).eq("SLA violated")] if "sla_status" in details else None,
             "pending": pending, "duplicate": duplicate}
    if any(v is None or not {day_col, "Currency"} <= set(v.columns) for v in parts.values()):
        return []
    per_day = {}
    for metric, t in parts.items():
        days = pd.to_datetime(t[day_col], errors="coerce").dt.normalize()
        amounts = pd.to_numeric(t[amt_col], errors="coerce").fillna(0.0) if amt_col in t else 0.0
        frame = pd.DataFrame({"day": days, "cur": t["Currency"].astype(str), "amt": amounts}).dropna(subset=["day"])
        for (d, c), g in frame.groupby(["day", "cur"]):
            f = per_day.setdefault(d, {"counts": {m: {} for m in METRICS}, "amounts": {m: {} for m in METRICS}})
            f["counts"][metric][c] = int(len(g))
            f["amounts"][metric][c] = float(g["amt"].sum())
    # the days must add up to the month, table by table and currency by currency
    for m in METRICS:
        for c, n in month.counts[m].items():
            if sum(f["counts"][m].get(c, 0) for f in per_day.values()) != n:
                return []
    out = []
    for d, f in sorted(per_day.items()):
        if not month.start <= d <= month.end:
            continue
        counts = {m: {c: f["counts"][m].get(c, 0) for c in month.counts[m]} for m in METRICS}
        amounts = {m: {c: f["amounts"][m].get(c, 0.0) for c in month.amounts[m]} for m in METRICS}
        out.append(Fact(month.bank, "day", d, d, counts, amounts, None,
                        dict((weekend or {}).get(f"{d:%Y-%m-%d}", {})), weekend is not None, month.record, "month"))
    return out


def weekend_late(details: pd.DataFrame, bank) -> dict:
    """{bank day: {currency: n}} of the late payments (SLA violated) made on a Friday, Saturday or Sunday and given
    the following Monday as bank date - read from a report's Payment details sheet (the what-if panel)."""
    bank = engine.get_bank(bank)
    day_col = bank.COLUMNS["payment_date"]
    need = {"sla_status", "paid_at", day_col, "Currency"}
    if details is None or not need <= set(details.columns):
        return {}
    d = details[details["sla_status"].astype(str).eq("SLA violated")]
    bank_day = pd.to_datetime(d[day_col], errors="coerce").dt.normalize()
    paid_day = pd.to_datetime(d["paid_at"], errors="coerce").dt.normalize()
    gap = (bank_day - paid_day).dt.days
    hit = bank_day.dt.dayofweek.eq(0) & gap.between(1, 3) & paid_day.dt.dayofweek.isin([4, 5, 6])
    out = {}
    for day, cur in zip(bank_day[hit], d.loc[hit, "Currency"].astype(str)):
        key = f"{day:%Y-%m-%d}"
        out.setdefault(key, {})
        out[key][cur] = out[key].get(cur, 0) + 1
    return out


def combine(facts: list) -> list:
    """Every fact to use, nothing counted twice. Months: the monthly reports. Weeks: a weekly report when there is
    one (it covers the whole week, Monday to Sunday); otherwise the week blocks of the monthly reports (a week cut
    by a month end then has its two parts, added together in the week view)."""
    monday = lambda f: f.start - pd.Timedelta(days=f.start.dayofweek)
    weekly = {(f.bank, monday(f)) for f in facts if f.kind == "week" and f.source == "week"}
    return [f for f in facts if f.kind in ("month", "day")
            or (f.source == "week")
            or (f.bank, monday(f)) not in weekly]


# =====================================================================================================================
# PERIODS
# =====================================================================================================================
def period_key(day: pd.Timestamp, kind: str) -> pd.Timestamp:
    """The first day of the week (Monday) / month / quarter / year that contains `day`."""
    day = pd.Timestamp(day).normalize()
    if kind == "day":
        return day
    if kind == "week":
        return day - pd.Timedelta(days=day.dayofweek)
    if kind == "month":
        return day.replace(day=1)
    if kind == "quarter":
        return day.replace(month=(day.month - 1) // 3 * 3 + 1, day=1)
    return day.replace(month=1, day=1)


def period_end(key: pd.Timestamp, kind: str) -> pd.Timestamp:
    """The last day of the period that starts on `key`."""
    if kind == "day":
        return key
    if kind == "week":
        return key + pd.Timedelta(days=6)
    months = {"month": 1, "quarter": 3, "year": 12}[kind]
    return key + pd.DateOffset(months=months) - pd.Timedelta(days=1)


def period_label(key: pd.Timestamp, kind: str, short: bool = False) -> str:
    if kind == "day":
        return f"{key:%d/%m}" if short else f"{key:%A %d/%m/%Y}"
    if kind == "week":                                  # 'W1 Oct' / 'Week 1 · October 2026 (28/09)' (engine.week_number)
        month, n = engine.week_number(key)
        return f"W{n} {month:%b}" if short else f"Week {n} · {month:%B %Y} ({key:%d/%m})"
    if kind == "month":
        return f"{key:%b %Y}" if short else f"{key:%B %Y}"
    if kind == "quarter":
        return f"Q{(key.month - 1) // 3 + 1} {key.year}"
    return f"{key.year}"


def _facts_for(facts: list, kind: str) -> list:
    """Days come from day facts, weeks from week facts (both inside the monthly reports); months, quarters and years
    from month facts."""
    return [f for f in facts if f.kind == (kind if kind in ("day", "week") else "month")]


def period_options(facts: list, kind: str) -> list:
    """The periods that have at least one report, newest first."""
    return sorted({period_key(f.start, kind) for f in _facts_for(facts, kind)}, reverse=True)


def select(facts: list, kind: str, first: pd.Timestamp, last: pd.Timestamp | None = None, banks=None) -> list:
    """The facts of the periods from `first` to `last` (one period when last is None), of some banks (None = all)."""
    first = pd.Timestamp(first)
    last = pd.Timestamp(first if last is None else last)
    base = "month" if kind == "custom" else kind
    return [f for f in _facts_for(facts, base)
            if first <= period_key(f.start, base) <= last and (banks is None or f.bank in banks)]


def previous_period(kind: str, first: pd.Timestamp, last: pd.Timestamp | None, mode: str):
    """(first, last) of the period compared with: "previous" = the same length just before, "year" = the same
    period one year before; None for "none"."""
    if mode == "none":
        return None
    first = pd.Timestamp(first)
    last = pd.Timestamp(first if last is None else last)
    if mode == "year":
        return first - pd.DateOffset(years=1), last - pd.DateOffset(years=1)
    if kind == "day":
        n = (last - first).days + 1
        return first - pd.Timedelta(days=n), last - pd.Timedelta(days=n)
    if kind == "week":
        n = (last - first).days // 7 + 1
        return first - pd.Timedelta(weeks=n), last - pd.Timedelta(weeks=n)
    step = {"month": 1, "custom": 1, "quarter": 3, "year": 12}[kind]
    n = ((last.year - first.year) * 12 + last.month - first.month) // step + 1
    return first - pd.DateOffset(months=step * n), last - pd.DateOffset(months=step * n)


def describe(kind: str, first, last=None) -> str:
    """'September 2026', 'Week of 31/08/2026', 'Jul 2026 – Sep 2026'."""
    base = "month" if kind == "custom" else kind
    if last is None or pd.Timestamp(last) == pd.Timestamp(first):
        return period_label(pd.Timestamp(first), base)
    return f"{period_label(pd.Timestamp(first), base, short=True)} – {period_label(pd.Timestamp(last), base, short=True)}"


# =====================================================================================================================
# FIGURES
# =====================================================================================================================
def totals(facts: list, currency: str = "All") -> dict:
    """The figures of some facts added up. currency "All": counts of every currency; else that currency only.
    {metric: n, "amounts": {currency: {metric: amount}}, "weekend": n, "weekend_known": bool, "reports": n,
     "banks": set, "currencies": [found]}"""
    t = {m: 0 for m in METRICS}
    t.update(amounts={}, weekend=0, weekend_known=bool(facts) and all(f.weekend_known for f in facts),
             reports=len({id(f.record) for f in facts}), banks={f.bank for f in facts}, currencies=set())
    for f in facts:
        for m in METRICS:
            for c, n in f.counts[m].items():
                t["currencies"].add(c)
                if currency in ("All", c):
                    t[m] += n
            for c, a in f.amounts[m].items():
                if currency in ("All", c):
                    t["amounts"].setdefault(c, {}).setdefault(m, 0.0)
                    t["amounts"][c][m] += a
        t["weekend"] += sum(n for c, n in f.weekend.items() if currency in ("All", c))
    t["currencies"] = sorted(t["currencies"], key=lambda c: (c != "RWF", c))
    return t


def rates(t: dict) -> dict:
    """received / processed and on time = (processed - late) / processed; None when nothing was processed."""
    p = t["processed"]
    return {"receipt": t["received"] / p if p else None, "ontime": (p - t["sla"]) / p if p else None,
            "not_received": p - t["received"] - t["pending"] - t["duplicate"] if p else 0}


def status(t: dict | None) -> str:
    """'bad' (a target missed), 'warn' (pending or paid with another channel), 'good', or 'none' (no report)."""
    if not t or not t["reports"]:
        return "none"
    r = rates(t)
    if r["receipt"] is not None and (r["receipt"] < TARGET_RECEIPT or r["ontime"] < TARGET_ONTIME):
        return "bad"
    if t["pending"] or t["duplicate"]:
        return "warn"
    return "good"


STATUS_TEXT = {"bad": "Off target", "warn": "Watch", "good": "On target", "none": "No report"}


def by_bank(facts: list, currency: str = "All", banks=None) -> dict:
    """{bank code: totals of its facts}, for every bank (a bank with no report: None), in the app's bank order."""
    codes = list(banks) if banks is not None else list(engine.BANKS)
    return {code: (totals([f for f in facts if f.bank == code], currency)
                   if any(f.bank == code for f in facts) else None) for code in codes}


def change(now, before):
    """Relative change (0.25 = +25 %), or None when there is nothing before to compare with."""
    if now is None or before in (None, 0):
        return None
    return (now - before) / before


def _money_at_risk(t: dict) -> dict:
    """{currency: amount} still at risk: pending notifications + paid with another channel."""
    return {c: a.get("pending", 0.0) + a.get("duplicate", 0.0) for c, a in t["amounts"].items()
            if a.get("pending", 0.0) + a.get("duplicate", 0.0)}


def _fmt_amount(v: float) -> str:
    for size, unit in ((1e9, "B"), (1e6, "M"), (1e3, "K")):
        if abs(v) >= size:
            return f"{v / size:,.1f}{unit}"
    return f"{v:,.0f}"


def attention(now_by_bank: dict, before_by_bank: dict, label: str) -> list:
    """The problems of the period, worst first: [{"kind", "bank", "title", "detail", "score", "metric"}].
    A missed target first (by how far it is missed), then pending, paid with another channel, missing reports."""
    items = []
    for code, t in now_by_bank.items():
        name = engine.BANKS[code].FULL_NAME if code in engine.BANKS else code
        if t is None:
            items.append({"kind": "info", "bank": code, "metric": "report", "score": 0.0,
                          "title": f"{name} · no report saved for {label}",
                          "detail": "The dashboard cannot include this bank until its report is made and saved."})
            continue
        r, b = rates(t), before_by_bank.get(code)
        rb = rates(b) if b else {}
        if r["ontime"] is not None and r["ontime"] < TARGET_ONTIME:
            weekend = (f" · {t['weekend']:,} of them made Fri–Sun and booked Monday" if t["weekend_known"]
                       and t["weekend"] else "")
            items.append({"kind": "bad", "bank": code, "metric": "ontime", "score": 3 + (TARGET_ONTIME - r["ontime"]),
                          "title": f"{name} · SLA {r['ontime']:.1%} on time (target {pct_target(TARGET_ONTIME)})",
                          "detail": f"{t['sla']:,} of {t['processed']:,} payments late{weekend}"
                                    + (f" · before: {rb['ontime']:.1%}" if rb.get("ontime") is not None else "")})
        if r["receipt"] is not None and r["receipt"] < TARGET_RECEIPT:
            items.append({"kind": "bad", "bank": code, "metric": "receipt", "score": 3 + (TARGET_RECEIPT - r["receipt"]),
                          "title": f"{name} · receipt rate {r['receipt']:.2%} (target {pct_target(TARGET_RECEIPT)})",
                          "detail": f"{t['processed'] - t['received']:,} of {t['processed']:,} notifications not "
                                    "received by Irembo"})
        for metric, what in (("pending", "notifications still pending"), ("duplicate", "paid with another channel")):
            if t[metric]:
                money = " · ".join(f"{c} {a.get(metric, 0.0):,.0f}" for c, a in t["amounts"].items() if a.get(metric))
                was = f" · before: {b[metric]:,}" if b else ""
                items.append({"kind": "warn", "bank": code, "metric": metric, "score": 1 + t[metric] / 1e4,
                              "title": f"{name} · {t[metric]:,} {what}",
                              "detail": (money or "no amount") + was})
    return sorted(items, key=lambda x: -x["score"])


def tiles(now: dict, before: dict | None, now_by_bank: dict, compared: str) -> list:
    """The 7 key figures: [(label, value, sub, kind)] - kind = "" / good / warn / bad (the colour of the tile)."""
    r = rates(now)
    rb = rates(before) if before else {}
    vs = f"vs {compared}" if before else "nothing before to compare"

    def ch(a, b, good_up=True):
        x = change(a, b)
        if x is None:
            return vs
        return f"{'▲' if x > 0 else '▼' if x < 0 else '='} {abs(x):.1%} {vs}"

    out = [("PAYMENTS", f"{now['processed']:,}", ch(now["processed"], before["processed"] if before else None), "")]
    for c in now["currencies"] or ["RWF"]:
        a = now["amounts"].get(c, {}).get("processed", 0.0)
        a0 = (before or {}).get("amounts", {}).get(c, {}).get("processed") if before else None
        out.append((f"AMOUNT · {c}", _fmt_amount(a), ch(a, a0), ""))
    if r["receipt"] is None:
        out.append(("RECEIPT RATE", "N/A", "no transactions", ""))
        out.append(("ON TIME (SLA)", "N/A", "no transactions", ""))
    else:
        out.append(("RECEIPT RATE", f"{r['receipt']:.2%}",
                    f"target {pct_target(TARGET_RECEIPT)} · {now['processed'] - now['received']:,} not received",
                    "good" if r["receipt"] >= TARGET_RECEIPT else "bad"))
        out.append(("ON TIME (SLA)", f"{r['ontime']:.1%}",
                    f"target {pct_target(TARGET_ONTIME)} · {now['sla']:,} late"
                    + (f" · before {rb['ontime']:.1%}" if rb.get("ontime") is not None else ""),
                    "good" if r["ontime"] >= TARGET_ONTIME else "bad"))
    risk = _money_at_risk(now)
    out.append(("MONEY AT RISK", " · ".join(f"{c} {_fmt_amount(a)}" for c, a in risk.items()) or "None",
                f"{now['pending']:,} pending · {now['duplicate']:,} paid with another channel",
                "warn" if risk else "good"))
    worst = next((a for a in attention(now_by_bank, {}, "") if a["kind"] in ("bad", "warn")), None)
    if worst is None:
        out.append(("NEEDS ATTENTION", "Nothing", "every bank with a report is on target", "good"))
    else:
        name = engine.BANKS[worst["bank"]].FULL_NAME if worst["bank"] in engine.BANKS else worst["bank"]
        out.append(("NEEDS ATTENTION", name, worst["title"].split(" · ", 1)[1], worst["kind"]))
    return out


def scorecard(now_by_bank: dict, before_by_bank: dict, currencies: list) -> pd.DataFrame:
    """One row per bank: status, payments and their change, amounts per currency, rates, pending, other channel,
    money at risk. A bank without a report is a row with its status "No report" and empty figures."""
    rows = []
    for code, t in now_by_bank.items():
        name = engine.BANKS[code].FULL_NAME if code in engine.BANKS else code
        b = before_by_bank.get(code)
        row = {"Bank": name, "Status": STATUS_TEXT[status(t)], "code": code}
        if t is None:
            rows.append(row)
            continue
        r = rates(t)
        row.update({"Payments": t["processed"],
                    "Change %": None if change(t["processed"], b["processed"] if b else None) is None
                    else round(change(t["processed"], b["processed"]) * 100, 1)})
        for c in currencies:
            row[f"Amount {c}"] = t["amounts"].get(c, {}).get("processed")
        risk = _money_at_risk(t)
        row.update({"Received %": None if r["receipt"] is None else round(r["receipt"] * 100, 2),
                    "On time %": None if r["ontime"] is None else round(r["ontime"] * 100, 1),
                    "Late": t["sla"], "Pending": t["pending"], "Other channel": t["duplicate"],
                    "Money at risk": " · ".join(f"{c} {a:,.0f}" for c, a in risk.items()) or "0"})
        rows.append(row)
    return pd.DataFrame(rows)


def trend(facts: list, kind: str, last: pd.Timestamp, n: int, currency: str, banks=None) -> pd.DataFrame:
    """Period by period, the n periods up to `last`: one row per period and bank with every measure.
    Columns: period (first day), label, bank, processed, received, sla, pending, duplicate, amount, receipt %,
    on time %."""
    base = "month" if kind == "custom" else kind
    keys = []
    k = period_key(last, base)
    for _ in range(n):
        keys.append(k)
        k = period_key(k - pd.Timedelta(days=1), base)
    keys = sorted(keys)
    rows = []
    for key in keys:
        part = [f for f in _facts_for(facts, base) if period_key(f.start, base) == key
                and (banks is None or f.bank in banks)]
        for code in sorted({f.bank for f in part}, key=lambda c: list(engine.BANKS).index(c) if c in engine.BANKS else 99):
            mine = [f for f in part if f.bank == code]
            t = totals(mine, currency)
            days = {d for f in mine for d in pd.date_range(f.start, f.end)}
            full = len(pd.date_range(key, period_end(key, base)))
            r = rates(t)
            amount_cur = currency if currency != "All" else "RWF"
            rows.append({"period": key, "label": period_label(key, base, short=True),
                         "bank": engine.BANKS[code].FULL_NAME if code in engine.BANKS else code, "code": code,
                         **{m: t[m] for m in METRICS}, "on_time": t["processed"] - t["sla"],
                         "amount": t["amounts"].get(amount_cur, {}).get("processed", 0.0),
                         "receipt_pct": None if r["receipt"] is None else r["receipt"] * 100,
                         "ontime_pct": None if r["ontime"] is None else r["ontime"] * 100,
                         "partial": len(days) < full})
    return pd.DataFrame(rows)


def waterfall(now_by_bank: dict, before_by_bank: dict, measure: str = "processed", currency: str = "RWF") -> list:
    """[(label, start, end, kind)] from the total before to the total now, one step per bank (its change):
    kind = total / up / down / none (a bank with no report on one side)."""
    def value(t):
        if t is None:
            return None
        return t["processed"] if measure == "processed" else t["amounts"].get(currency, {}).get("processed", 0.0)
    before_total = sum(value(t) or 0 for t in before_by_bank.values())
    steps, level = [("Before", 0, before_total, "total")], before_total
    for code in now_by_bank:
        a, b = value(now_by_bank.get(code)), value(before_by_bank.get(code))
        name = engine.BANKS[code].NAME if code in engine.BANKS else code
        if a is None and b is None:
            steps.append((name, level, level, "none"))
            continue
        d = (a or 0) - (b or 0)
        steps.append((name, level, level + d, "up" if d > 0 else "down" if d < 0 else "none"))
        level += d
    steps.append(("Now", 0, level, "total"))
    return steps


def merchants(now_facts: list, before_facts: list, currency: str) -> tuple:
    """(table, reports without merchant data): one row per merchant of that currency, notifications received now and
    before, the change, the share of the currency's total, and "new" / "dropped"."""
    def add(facts):
        acc, missing = {}, 0
        for f in facts:
            if f.merchants is None:
                missing += 1
                continue
            for (m, c), (n, a) in f.merchants.items():
                if c == currency:
                    n0, a0 = acc.get(m, (0, 0.0))
                    acc[m] = (n0 + n, a0 + a)
        return acc, missing
    # a report without merchant data says nothing about its merchants: only banks with data now are compared
    banks = {f.bank for f in now_facts if f.merchants is not None}
    now, missing = add(now_facts)
    before, _ = add([f for f in before_facts if f.bank in banks])
    total = sum(n for n, _ in now.values())
    rows = []
    for m in set(now) | set(before):
        n, a = now.get(m, (0, 0.0))
        n0 = before.get(m, (0, 0.0))[0] if before else None
        rows.append({"Merchant": m, "Transactions": n, "Amount": a, "Share %": n / total * 100 if total else 0.0,
                     "Before": n0, "Change %": None if not n0 else (n - n0) / n0 * 100,
                     "Note": "new" if before and m not in before else ("dropped" if m not in now else "")})
    table = pd.DataFrame(rows, columns=["Merchant", "Transactions", "Amount", "Share %", "Before", "Change %", "Note"])
    return table.sort_values(["Transactions", "Merchant"], ascending=[False, True]).reset_index(drop=True), missing


def what_if_weekend(t_bank: dict, t_all: dict, share: float) -> dict:
    """What the on-time rate would be if late payments made Friday-Sunday and booked Monday counted as on time
    (share = the part of them that would count: 0..1). Nothing is changed in the reports."""
    moved = round(t_bank["weekend"] * share)
    p, pa = t_bank["processed"], t_all["processed"]
    return {"moved": moved, "weekend": t_bank["weekend"],
            "bank_now": (p - t_bank["sla"]) / p if p else None, "bank_after": (p - t_bank["sla"] + moved) / p if p else None,
            "all_now": (pa - t_all["sla"]) / pa if pa else None, "all_after": (pa - t_all["sla"] + moved) / pa if pa else None}


def what_if_target(now_by_bank: dict, target: float) -> list:
    """[(bank, on-time rate, passes)] with another on-time target."""
    out = []
    for code, t in now_by_bank.items():
        if t is None:
            continue
        r = rates(t)["ontime"]
        out.append((engine.BANKS[code].FULL_NAME if code in engine.BANKS else code, r, r is not None and r >= target))
    return out


def coverage(all_facts: list, kind: str, first, last, banks) -> pd.DataFrame:
    """What the view is built from: for each bank and period, the report used and whether it has merchant data."""
    base = "month" if kind == "custom" else kind
    keys, k = [], period_key(last if last is not None else first, base)
    while k >= pd.Timestamp(first):
        keys.append(k)
        k = period_key(k - pd.Timedelta(days=1), base)
    rows = []
    for key in keys:
        for code in banks:
            part = [f for f in _facts_for(all_facts, base) if f.bank == code and period_key(f.start, base) == key]
            reports = sorted({f"{f.record.name}" for f in part if f.record is not None})
            rows.append({"Period": period_label(key, base), "Bank": engine.BANKS[code].FULL_NAME
                         if code in engine.BANKS else code,
                         "Report": ", ".join(sorted({("monthly" if f.source == "month" else "weekly")
                                                     + (f" v{f.record.version}" if f.record is not None else "")
                                                     for f in part})) or "not saved",
                         "Merchant data": ("yes" if part and all(f.merchants is not None for f in part)
                                           else "missing" if part else "–"),
                         "Files": len(reports)})
    return pd.DataFrame(rows)


def findings(kind: str, trend_rows: pd.DataFrame, now_by_bank: dict, before_by_bank: dict, label_before: str) -> dict:
    """One sentence per chart, written from the numbers: {"trend", "banks", "status", "waterfall"}."""
    out = {}
    if len(trend_rows):
        per = trend_rows.groupby(["period", "label"], as_index=False)[["processed", "sla"]].sum()
        peak = per.loc[per["processed"].idxmax()]
        avg = per["processed"].mean()
        out["trend"] = (f"Peak: {peak['label']} with {int(peak['processed']):,} payments; "
                        f"average {avg:,.0f} per {kind} over {len(per)} {kind}s.")
        late = per.sort_values("sla", ascending=False).head(2)
        late = late[late["sla"] > 0]
        if len(late):
            out["status"] = ("Most late payments: " + " and ".join(f"{int(r.sla):,} in {r.label}" for r in late.itertuples())
                             + ".")
        else:
            out["status"] = "No late payment in these periods."
    shown = [(engine.BANKS[c].FULL_NAME if c in engine.BANKS else c, rates(t)["ontime"])
             for c, t in now_by_bank.items() if t is not None and rates(t)["ontime"] is not None]
    if shown:
        under = [f"{n} {r:.1%}" for n, r in shown if r < TARGET_ONTIME]
        over = [f"{n} {r:.1%}" for n, r in shown if r >= TARGET_ONTIME]
        out["banks"] = (((f"Under the {pct_target(TARGET_ONTIME)} target: " + ", ".join(under) + ". ") if under else "")
                        + (("On target: " + ", ".join(over) + ".") if over else "")).strip()
    deltas = []
    for code, t in now_by_bank.items():
        b = before_by_bank.get(code)
        if t is not None and b is not None:
            deltas.append((engine.BANKS[code].NAME if code in engine.BANKS else code, t["processed"] - b["processed"]))
    if deltas:
        total = sum(d for _, d in deltas)
        main = max(deltas, key=lambda x: abs(x[1]))
        out["waterfall"] = (f"Payments {'rose' if total > 0 else 'fell' if total < 0 else 'did not change'}"
                            + (f" by {abs(total):,}" if total else "") + f" vs {label_before}: "
                            + ", ".join(f"{n} {d:+,}" for n, d in deltas)
                            + (f" (most from {main[0]})." if total and len(deltas) > 1 else "."))
    return out


def partial(facts: list, kind: str, first, last=None) -> list:
    """[(bank code, period key, saved, in the period, unit)] of the periods of the view that are not complete:
    a year or a quarter counts its months saved ('2 of 3 months'), a week or a month its days ('3 of 7 days').
    Their figures are lower than the real ones."""
    base = "month" if kind == "custom" else kind
    out = []
    for key in sorted({period_key(f.start, base) for f in select(facts, kind, first, last)}):
        for code in sorted({f.bank for f in _facts_for(facts, base) if period_key(f.start, base) == key}):
            mine = [f for f in _facts_for(facts, base) if f.bank == code and period_key(f.start, base) == key]
            if base in ("year", "quarter"):
                saved, full, unit = len({f.start for f in mine}), {"year": 12, "quarter": 3}[base], "months"
            else:
                saved = len({d for f in mine for d in pd.date_range(f.start, f.end)})
                full, unit = len(pd.date_range(key, period_end(key, base))), "days"
            if saved < full:
                out.append((code, key, saved, full, unit))
    return out


def breakdown(facts: list, kind: str, first, last, currency: str, banks=None) -> pd.DataFrame:
    """The level below the period chosen (a year -> its quarters, a quarter -> its months, a month -> its weeks,
    a week -> its days), one row each with its figures; empty for a day."""
    base_kind = "month" if kind == "custom" else kind
    child = CHILD.get(base_kind)
    if child is None:
        return pd.DataFrame()
    first = pd.Timestamp(first)
    end = period_end(period_key(pd.Timestamp(last if last is not None else first), "month" if kind == "custom" else kind),
                     "month" if kind == "custom" else kind)
    rows = []
    for key in sorted({period_key(f.start, child) for f in _facts_for(facts, child)
                       if first <= f.start <= end and (banks is None or f.bank in banks)}):
        part = [f for f in _facts_for(facts, child) if period_key(f.start, child) == key and first <= f.start <= end
                and (banks is None or f.bank in banks)]
        t = totals(part, currency)
        r = rates(t)
        cut = key < first or period_end(key, child) > end      # e.g. a week across the month end: its part only
        name = (f"Week {engine.week_number(key, month=first)[1]}" if child == "week" and base_kind == "month"
                else period_label(key, child))          # inside a month: its own weeks (Week 1 … Week 5)
        rows.append({"key": key, "Period": name + (" (part in this period)" if cut else ""),
                     "Payments": t["processed"],
                     "Received %": None if r["receipt"] is None else round(r["receipt"] * 100, 2),
                     "On time %": None if r["ontime"] is None else round(r["ontime"] * 100, 1),
                     "Late": t["sla"], "Pending": t["pending"], "Other channel": t["duplicate"],
                     "Status": STATUS_TEXT[status(t)]})
    return pd.DataFrame(rows)


def children(facts: list, kind: str, first, last=None, currency: str = "All", banks=None) -> list:
    """The periods one level below the one chosen, for the charts (a year -> its quarters, a quarter -> its months,
    a month -> its weeks, a week -> its days), in order. Each: {"key", "label", "partial" (a week cut by the month
    end: only its part in this period), the 5 counts, "amounts": {currency: amount received}}. Empty for a day."""
    base = "month" if kind == "custom" else kind
    child = CHILD.get(base)
    if child is None:
        return []
    first = pd.Timestamp(first)
    end = period_end(period_key(pd.Timestamp(last if last is not None else first), base), base)
    pool = [f for f in _facts_for(facts, child) if first <= f.start <= end and (banks is None or f.bank in banks)]
    out = []
    for key in sorted({period_key(f.start, child) for f in pool}):
        t = totals([f for f in pool if period_key(f.start, child) == key], currency)
        label = (f"{key:%a %d/%m}" if child == "day"
                 else f"Week {engine.week_number(key, month=first)[1]}" if child == "week"   # the month's own weeks
                 else period_label(key, child, short=True))
        out.append({"key": key, "label": label, "partial": key < first or period_end(key, child) > end,
                    **{m: t[m] for m in METRICS},
                    "amounts": {c: a.get("received", 0.0) for c, a in t["amounts"].items()}})
    return out


# =====================================================================================================================
# COMPLETENESS: periods whose every day (or month) is saved; the reports still to make
# =====================================================================================================================
def is_complete(facts: list, kind: str, key, banks=None) -> bool:
    """True when every bank that has a report in the period has all of it saved: every month of a year or a quarter,
    every day of a month or a week (a day is always complete)."""
    return not [g for g in partial([f for f in facts if banks is None or f.bank in banks], kind, key)
                if g[1] == period_key(pd.Timestamp(key), "month" if kind == "custom" else kind)]


def complete_options(facts: list, kind: str, banks=None) -> list:
    """The periods that have reports and are complete, newest first."""
    return [k for k in period_options([f for f in facts if banks is None or f.bank in banks], kind)
            if is_complete(facts, kind, k, banks)]


def last_ended_month(today=None) -> pd.Timestamp:
    """The first day of the last month that has ended (a month can be reconciled only once it has ended)."""
    today = pd.Timestamp.now() if today is None else pd.Timestamp(today)
    return (today.normalize().replace(day=1) - pd.Timedelta(days=1)).replace(day=1)


def reports_to_make(facts: list, today=None, banks=None) -> list:
    """[(bank code, month)] of the monthly reports missing: for every bank (or only `banks`), each month from the
    first month saved (by any bank) to the last month that has ended. Oldest first."""
    months = sorted({f.start.replace(day=1) for f in facts if f.kind == "month"})
    if not months:
        return []
    last = last_ended_month(today)
    span = list(pd.date_range(months[0], max(last, months[-1]), freq="MS"))
    have = {(f.bank, f.start.replace(day=1)) for f in facts if f.kind == "month"}
    return [(code, m) for m in span for code in engine.BANKS
            if (banks is None or code in banks) and (code, m) not in have]


def coverage_line(facts: list, today=None, banks=None) -> str:
    """'Aug 2026 – Sep 2026 saved · 2 monthly reports to make' (the months of the saved monthly reports), for every
    bank or only `banks` (the bank chosen on the dashboard)."""
    months = sorted({f.start.replace(day=1) for f in facts if f.kind == "month" and (banks is None or f.bank in banks)})
    if not months:
        return "No monthly report saved yet"
    todo = reports_to_make(facts, today, banks)
    span = (f"{months[0]:%b %Y}" if len(months) == 1 else f"{months[0]:%b %Y} – {months[-1]:%b %Y}")
    if not todo:
        return f"{span} saved · every monthly report is saved"
    named = ", ".join(f"{engine.BANKS[c].NAME} {m:%B %Y}" for c, m in todo[:4])
    return f"{span} saved · not saved yet: {named}" + (f" and {len(todo) - 4} more" if len(todo) > 4 else "")
