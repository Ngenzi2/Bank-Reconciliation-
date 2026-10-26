"""
THE SLACK MESSAGE OF A REPORT (step 6, "Send to Slack")
======================================================
Builds the team message from a reconciliation run, in the format agreed on 2026-10-12. Every number is read from
the report (the summary tables of the main block, its Payment per merchant table); nothing is reconciled again.

    With "Compare with previous period": "<period> vs <previous period>", and each figure with its change
    ("+32.56% vs Prev-Week"). Without it: the period's own figures only.

    Colours: 🟢 / 🔴 - success rate >= SUCCESS_OK, SLA (within SLA) >= SLA_OK (the targets saved in the app's
    Settings; 99% and 90% by default), pending = 0, paid with other channels = 0.
    Trend: 📈 a rise, 📉 a fall. Amounts per currency (never added together); "N/A" when there is none.
"""
from __future__ import annotations

import re

import pandas as pd

import engine

GREEN, RED = ":large_green_circle:", ":red_circle:"
UP, DOWN = ":chart_with_upwards_trend:", ":chart_with_downwards_trend:"
SUCCESS_OK, SLA_OK = 0.99, 0.90          # the defaults; the app sets the targets saved in Settings
P, R, S = "Overall processed payments", "Overall notifications received by Irembo", "SLA Violated"
PE, DU = "Pending notifications", "Duplicate Payment"


def _count(tables: dict, title: str, cur=None) -> int:
    t = tables[title]
    row = t[t["Currency"].isna()] if cur is None else t[t["Currency"] == cur]
    return int(row["Number of Transactions"].iloc[0]) if len(row) else 0


def _amounts(tables: dict, title: str) -> str:
    t = tables[title]
    parts = [f"{a:,.0f} {c}" for c, n, a in zip(t["Currency"], t["Number of Transactions"], t["Total Amount"])
             if pd.notna(c) and n]
    return " · ".join(parts) or "N/A"


def _within_amounts(tables: dict, currencies: list) -> str:
    """Amount within SLA = processed amount - SLA violated amount, per currency."""
    out = []
    for cur in currencies:
        p, s = tables[P], tables[S]
        n = _count(tables, P, cur) - _count(tables, S, cur)
        if n > 0:
            a = float(p.loc[p["Currency"] == cur, "Total Amount"].sum()) - float(s.loc[s["Currency"] == cur, "Total Amount"].sum())
            out.append(f"{a:,.0f} {cur}")
    return " · ".join(out) or "N/A"


def _vs(now: int, before: int) -> str:
    if not before:
        return "0.00%" if not now else "+100.00%"
    return f"{(now - before) / before * 100:+.2f}%"


def _trend(now: int, before: int) -> str:
    return UP if now > before else DOWN if now < before else ""


def _dot(good: bool) -> str:
    return GREEN if good else RED


def period_names(period) -> tuple[str, str, str, str]:
    """(kind word, this period, the previous one, Prev-label): 'weekly', 'Aug-Week4', 'Aug-Week3', 'Prev-Week'."""
    def week_name(start):
        first = start.replace(day=1)
        n = sum(1 for d in pd.date_range(first, start) if d.dayofweek == 0) + (0 if first.dayofweek == 0 else 1)
        return f"{start:%b}-Week{n}"
    prev = period.previous()
    if period.report_type == "week":
        return "weekly", week_name(period.start), week_name(prev.start), "Prev-Week"
    if period.report_type == "day":
        return "daily", f"{period.start:%d/%m/%Y}", f"{prev.start:%d/%m/%Y}", "Prev-Day"
    return "monthly", f"{period.start:%B %Y}", f"{prev.start:%B %Y}", "Prev-Month"


def build(bank, period, blocks: list, merchants: dict, previous_counts: dict | None = None,
          link: str | None = None) -> str:
    """The message of the report's main block (the month, the week or the day)."""
    main = next((b for b in blocks if b[0] in ("Summary-Monthly", "Summary-Daily")), blocks[0])
    t = main[2]
    kind, name, prev_name, prev_label = period_names(period)
    compare = previous_counts is not None
    pc = previous_counts or {}

    def before(title):
        return int((pc.get(title) or {}).get("Total", 0))

    proc, rec, sla_bad = _count(t, P), _count(t, R), _count(t, S)
    pend, dup = _count(t, PE), _count(t, DU)
    within = proc - sla_bad
    proc0, within0 = before(P), before(P) - before(S)
    pend0, dup0 = before(PE), before(DU)
    rate = rec / proc if proc else 0.0
    sla_rate = within / proc if proc else 0.0
    share = lambda n: f"{n / proc if proc else 0:.2%} OF OVERALL"

    def change(now, prev_n, arrow=True):
        if not compare:
            return ""
        return f" {_trend(now, prev_n) + ' ' if arrow and _trend(now, prev_n) else ''}{_vs(now, prev_n)} vs {prev_label}"

    m = merchants.get(main[1])
    names = sorted(set(m.loc[m["Merchant name"] != "All", "Merchant name"].astype(str))) if m is not None and len(m) else []
    title = f"{name} vs {prev_name}" if compare else name
    lines = [
        f"Hello Team, here is the {kind} report on",
        f"{bank.NAME} payments processed in {title}",
        "",
        f"• Merchants who used this channel: [{', '.join(names) if names else 'none'}]",
        "",
        f"• Overall processed payments ({bank.NAME}):",
        f"   • Total number of Transactions: {proc:,}{change(proc, proc0)}",
        f"   • Total Amount: {_amounts(t, P)}",
        "",
        "• Overall notifications received by Irembo:",
        f"   • Total number of Transactions: {rec:,} {_dot(rate >= SUCCESS_OK)} {rate:.2%} Success rate "
        f"{_dot(rate >= SUCCESS_OK)} with a failure rate of {1 - rate:.2%}",
        f"   • Total Amount: {_amounts(t, R)}",
        "",
        f"• SLA Performance (within SLA): {_dot(sla_rate >= SLA_OK)}",
        f"   • Total number of transactions: {within:,} {share(within)} {change(within, within0)}".rstrip(),
        f"   • Total Amount: {_within_amounts(t, list(bank.CURRENCIES))}",
        "",
        f"• Pending notifications: {_dot(pend == 0)}",
        f"   • Total number of transactions: {pend:,} {share(pend)} {change(pend, pend0, arrow=False)}".rstrip(),
        f"   • Total Amount: {_amounts(t, PE)}",
        "",
        f"• Paid with other channels (potential double payments): {_dot(dup == 0)}",
        f"   • Total number of transactions: {dup:,} {share(dup)} {change(dup, dup0, arrow=False)}".rstrip(),
        f"   • Total Amount: {_amounts(t, DU)}",
        "",
        "• Findings and updates",
    ]
    if compare:
        arrow = f" {_trend(proc, proc0)}" if _trend(proc, proc0) else ""
        lines.append(f"   • Overall, there was a {_vs(proc, proc0)}{arrow} in the successfully processed "
                     f"transactions vs previous {kind.replace('ly', '').replace('dai', 'day')}, with a success rate of "
                     f"{rate:.2%} {_dot(rate >= SUCCESS_OK)},")
    else:
        lines.append(f"   • Overall, {proc:,} transactions were processed, with a success rate of {rate:.2%} "
                     f"{_dot(rate >= SUCCESS_OK)},")
    lines += [f"     and the SLA performance over the last {kind.replace('ly', '').replace('dai', 'day')} was "
              f"{sla_rate:.2%} {_dot(sla_rate >= SLA_OK)} which is {'above' if sla_rate >= SLA_OK else 'below'} "
              f"{SLA_OK * 100:g}%",
              f"   • We have {pend:,} pending transactions {_dot(pend == 0)}",
              f"   • Full details in the Google Sheets report <{link}|Here>" if link
              else "   • Full details in the report on Google Drive"]
    return "\n".join(lines)



# =====================================================================================================================
# THE STEP 6 MESSAGE (format approved on 2026-10-15): Slack blocks - the title, "Overall performance" (status against
# the targets), "Reconciliation summary" (a table of the 5 summary tables), "Top merchants by RWF notifications"
# and the link of the report on Google Drive. Every number is read from the report's main block.
# =====================================================================================================================
TABLE_ROWS = [("Payments processed", P), ("Received by Irembo", R), ("Late (SLA violated)", S),
              ("Pending notification", PE), ("Paid with another channel", DU)]


def _amount(tables: dict, title: str, cur: str) -> float:
    t = tables[title]
    row = t[t["Currency"] == cur]
    return float(row["Total Amount"].iloc[0]) if len(row) and pd.notna(row["Total Amount"].iloc[0]) else 0.0


def money(value) -> str:
    """An amount for the message: in full below a million (25,000), short from a million (39.6M)."""
    v = float(value) if pd.notna(value) else 0.0
    return f"{v / 1e6:,.1f}M" if abs(v) >= 1e6 else f"{v:,.0f}"


def period_line(period) -> str:
    """'1–31 August 2026', '5–11 October 2026', '28 September – 4 October 2026', '15 October 2026'."""
    a, b = period.start, period.last_day
    if getattr(period, "report_type", "") == "week":    # 'Week 1 · 28 September – 4 October 2026'
        month, n = engine.week_number(a)
        days = (f"{a.day}–{b.day} {b:%B %Y}" if (a.month, a.year) == (b.month, b.year)
                else f"{a.day} {a:%B}{'' if a.year == b.year else f' {a.year}'} – {b.day} {b:%B %Y}")
        return f"Week {n} · {days}"
    if a == b:
        return f"{a.day} {a:%B %Y}"
    if (a.month, a.year) == (b.month, b.year):
        return f"{a.day}–{b.day} {b:%B %Y}"
    return f"{a.day} {a:%B}{'' if a.year == b.year else f' {a.year}'} – {b.day} {b:%B %Y}"


def status(rate: float, sla_rate: float, follow: int) -> tuple:
    """(emoji, words) of "Overall performance", from the targets set in Settings."""
    if rate >= SUCCESS_OK and sla_rate >= SLA_OK:
        return (":white_check_mark:", "On target — follow-up required") if follow else (":white_check_mark:", "On target")
    return ":red_circle:", "Off target"


def build_blocks(bank, period, blocks: list, merchants: dict, links: dict | None, made_at,
                 native_table: bool = True, notes: list | None = None) -> dict:
    """The Slack payload {"text", "blocks"} of a report. links: {"xlsx": url, "pdf": url} of the files saved on
    Google Drive. native_table=False: the summary as a text table in a code box (for a Slack that does not accept
    the table block)."""
    main = next((b for b in blocks if b[0] in ("Summary-Monthly", "Summary-Daily")), blocks[0])
    t = main[2]
    proc, rec, late = _count(t, P), _count(t, R), _count(t, S)
    pend, dup = _count(t, PE), _count(t, DU)
    rate = rec / proc if proc else 0.0
    sla_rate = (proc - late) / proc if proc else 0.0
    kind = {"month": "Monthly", "week": "Weekly", "day": "Daily"}[period.report_type]
    name = period.name[0].upper() + period.name[1:]
    made = pd.Timestamp(made_at)
    currencies = [c for c in t[P]["Currency"] if pd.notna(c) and _count(t, P, c)]
    icon, words = status(rate, sla_rate, pend + dup)
    md = lambda x: {"type": "mrkdwn", "text": x}
    head = lambda x: {"type": "header", "text": {"type": "plain_text", "text": x[:150], "emoji": True}}

    head_row = ["Category", "Transactions", "% of processed"] + currencies
    rows = [[label, f"{_count(t, title):,}", "100%" if title == P else f"{_count(t, title) / proc if proc else 0:.2%}"]
            + [f"{_amount(t, title, c):,.0f}" for c in currencies] for label, title in TABLE_ROWS]
    if native_table:
        raw = lambda x: {"type": "raw_text", "text": x}
        table = {"type": "table", "column_settings": [{"is_wrapped": True}] + [{"align": "right"}] * (len(head_row) - 1),
                 "rows": [[raw(x) for x in head_row]] + [[raw(x) for x in r] for r in rows]}
    else:
        w = [max(len(r[i]) for r in [head_row] + rows) for i in range(len(head_row))]
        line = lambda r: "  ".join(x.ljust(w[0]) if i == 0 else x.rjust(w[i]) for i, x in enumerate(r))
        table = {"type": "section", "text": md("```" + "\n".join([line(head_row), "─" * len(line(head_row))]
                                                                 + [line(r) for r in rows]) + "```")}
    risk = " · ".join(f"{c} {v:,.0f}" for c in currencies
                      if (v := _amount(t, PE, c) + _amount(t, DU, c))) or "none"
    out = [head(f"📊 {bank.FULL_NAME} — {kind} Reconciliation | {name}"),
           {"type": "context", "elements": [md(f"Period: {period_line(period)} · Report date: {made.day} {made:%B %Y}")]}]
    for note in notes or []:                      # days with no statement (did not block): someone should check them
        label, days = (note.split(": ", 1) + [""])[:2]
        out.append({"type": "section", "text": md(f":warning: *No statement for {days}* ({label}): no transaction "
                                                  "those days, or the statement is missing - please check.")})
    out += [{"type": "divider"},
           head("Overall performance"),
           {"type": "section", "text": md(f"{icon} *{words}*\nReceived: `{rate:.2%}` · Target ≥ {SUCCESS_OK * 100:g}%\n"
                                          f"On time: `{sla_rate:.2%}` · Target ≥ {SLA_OK * 100:g}%")},   # rates highlighted
           {"type": "divider"},
           head("Reconciliation summary"),
           table,
           {"type": "context", "elements": [md(f"Exception value (pending + other channel): {risk}")]},
           {"type": "divider"}]
    m = merchants.get(main[1]) if merchants else None
    if m is not None and len(m):
        m = m[(m["Currency"] == "RWF") & (m["Merchant name"] != "All")].head(3)
    if m is not None and len(m):
        out += [head("🏢 Top merchants by RWF notifications"),
                {"type": "section", "text": md("\n".join(
                    f"*{i}.* {r['Merchant name']} — {int(r['Number of Transactions']):,} · {float(r['%']):.1%} · "
                    f"RWF {money(r['Total Amount'])}" for i, (_, r) in enumerate(m.iterrows(), 1)))},
                {"type": "divider"}]
    names = {"xlsx": "📊 Excel report", "pdf": "📄 PDF report"}
    shown = [f"<{links[ext]}|{names[ext]}>" for ext in ("xlsx", "pdf") if (links or {}).get(ext)]
    out.append({"type": "section", "text": md("📂 *Explore the full report:* " + " · ".join(shown) if shown
                                              else "📂 The report is not saved on Google Drive yet")})
    return {"text": f"{bank.FULL_NAME} · {name} · {proc:,} payments · {rate:.2%} received · {sla_rate:.2%} on time",
            "blocks": out}


def preview_text(payload: dict) -> str:
    """The message as plain text, for "See message" in the app."""
    lines = []
    for b in payload["blocks"]:
        if b["type"] == "header":
            lines += ["", b["text"]["text"].upper()]
        elif b["type"] == "context":
            lines += [e["text"] for e in b["elements"]]
        elif b["type"] == "section":
            lines.append(b["text"]["text"].replace("```", ""))
        elif b["type"] == "table":
            for r in b["rows"]:
                lines.append(" | ".join(c["text"] for c in r))
    text = "\n".join(lines).strip()
    for code, emoji in ((":white_check_mark:", "✅"), (":red_circle:", "🔴"), (":large_green_circle:", "🟢"),
                        (":warning:", "⚠️")):
        text = text.replace(code, emoji)
    text = re.sub(r"<([^|>]+)\|([^>]+)>", r"\2: \1", text)          # <link|words> -> words: link
    return text.replace("*", "").replace("`", "")

def build_no_transactions(bank, period, link: str | None = None, previous: dict | None = None) -> str:
    """The message of a No-Transaction Report (step 1, "No-Transaction Report"): the same opening as a normal
    report, the status, and each figure read from the report's own rows (engine.no_transaction_rows); the rates
    are "N/A" when nothing was processed (there is nothing to measure, so no 🟢 / 🔴). previous: the counts of the
    previous period's report ({table: {currency or "Total": n}}), said beside each figure when found."""
    from engine import NO_TX_NOTE, NO_TX_STATUS, no_transaction_rows
    kind, name, prev_name, _ = period_names(period)
    n = {}
    for title, _cur, count, *_ in no_transaction_rows(bank):
        n[title] = n.get(title, 0) + int(count)
    proc, rec = n.get(P, 0), n.get(R, 0)
    rate = f"{rec / proc:.2%}" if proc else "N/A — No transactions"
    sla = f"{(proc - n.get(S, 0)) / proc:.2%}" if proc else "N/A — No transactions"

    def was(title):
        """' (August 2026: 1,250)' - the previous period's count of this table, when its report was found."""
        if previous is None:
            return ""
        return f" ({prev_name}: {int((previous.get(title) or {}).get('Total', 0)):,})"

    return "\n".join([
        f"Hello Team, here is the {kind} report on",
        f"{bank.NAME} payments processed in {name}",
        "",
        f"• Reconciliation status: :information_source: {NO_TX_STATUS}",
        f"   • {NO_TX_NOTE}",
        "",
        f"• Overall processed payments ({bank.NAME}): {proc:,}{was(P)}",
        f"• Overall notifications received by Irembo: {rec:,}{was(R)} · Success rate: {rate}",
        f"• SLA Performance (within SLA): {sla}",
        f"• Pending notifications: {n.get(PE, 0):,}{was(PE)}",
        f"• Paid with other channels (potential double payments): {n.get(DU, 0):,}{was(DU)}",
        "",
        f"• Full details in the Google Sheets report <{link}|Here>" if link
        else "• Full details in the report on Google Drive",
    ])
