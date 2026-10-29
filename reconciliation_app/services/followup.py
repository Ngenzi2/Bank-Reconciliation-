"""
FOLLOW-UP: WHAT HAPPENED TO THE PENDING AND PAID-TWICE ITEMS AFTER THE REPORT (decided 22/10/2026)
=================================================================================================
A saved report is a snapshot: its Pending notification and Duplicate payment sheets list the items still to follow up
on the day it was made. Afterwards the team records what finally happened to each item; the report itself never
changes. The decisions are kept on Google Drive, next to the reports ("Follow-up decisions.json" at the top of
Reconciliation Reports), so everyone sees the same; anyone using the app can record them.

    items_from_report(path, record)   the open items of one saved Excel report
    collect(...)                      the items of every report since the start date (Settings), each item once
    load_decisions / save_decisions   the shared decisions on Drive
    summary(...)                      open / resolved / amounts / oldest, for the page and the dashboard
    reminder_text(...)                the Slack reminder: the items open for more than N days

Decisions: a pending item (paid at the bank, IremboPay still NEW) -> Received later · Refunded · Not an Irembo
payment; a paid-twice item -> Refunded · Kept / used for another invoice · Bank error reversed. "Received later"
counts in the dashboard's "rate after follow-up" (the report's own rate is never changed).
"""
from __future__ import annotations

import json
import pickle
from datetime import datetime
from pathlib import Path

import pandas as pd

import engine

STORE_NAME = "Follow-up decisions.json"
DECISIONS = {"Pending": ["Received later", "Refunded", "Not an Irembo payment"],
             "Paid twice": ["Refunded", "Kept / used for another invoice", "Bank error reversed"]}
SHEETS = {"Pending notification": "Pending", "Duplicate payment other channel": "Paid twice"}
CACHE_VERSION = 1


def _text(v) -> str:
    return "" if v is None or (isinstance(v, float) and pd.isna(v)) or str(v).strip().lower() in ("nan", "none", "nat") \
        else str(v).strip()


def items_from_report(path: Path, record, bank_code: str) -> list:
    """[{key, bank, type, invoice, reference, amount, currency, paid_on, merchant, provider, report, reported}] of
    one saved Excel report: the rows of its Pending notification and Duplicate payment sheets."""
    bank = engine.get_bank(bank_code)
    c = bank.COLUMNS
    try:
        sheets = pd.read_excel(path, sheet_name=list(SHEETS), dtype={"invoice_number": str},
                               engine=engine.EXCEL_ENGINE)
    except Exception:                                   # an older report without these sheets: nothing to follow up
        return []
    reported = pd.Timestamp(record.modified) if record.modified is not None else pd.Timestamp.now()
    if reported.tzinfo is not None:
        reported = reported.tz_localize(None)
    out = []
    for sheet, kind in SHEETS.items():
        df = sheets.get(sheet)
        if df is None or df.empty:
            continue
        for _, r in df.iterrows():
            inv, ref = _text(r.get("invoice_number")), _text(r.get(c["reference"])) if c.get("reference") else ""
            amount = pd.to_numeric(r.get(c["amount"]), errors="coerce")
            amount = 0.0 if pd.isna(amount) else float(amount)
            paid_on = pd.to_datetime(r.get(c["payment_date"]), errors="coerce")
            out.append({"key": f"{bank.CODE}|{kind}|{inv}|{ref}|{amount:.2f}", "bank": bank.CODE, "type": kind,
                        "invoice": inv, "reference": ref, "amount": amount,
                        "currency": _text(r.get("Currency")) or bank.CURRENCIES[0],
                        "paid_on": None if pd.isna(paid_on) else pd.Timestamp(paid_on).normalize(),
                        "merchant": _text(r.get("name")), "provider": _text(r.get("payment_provider")),
                        "report": record.name, "reported": reported.normalize()})
    return out


def latest_reports(records: list, start) -> list:
    """The saved Excel reports whose period starts on or after `start` (None: every report), the latest version of
    each."""
    latest = {}
    for r in records:
        if r.ext != "xlsx" or (start is not None and pd.Timestamp(r.start) < pd.Timestamp(start)):
            continue
        if r.stem not in latest or r.version > latest[r.stem].version:
            latest[r.stem] = r
    return list(latest.values())


def collect(records: list, start: pd.Timestamp, read) -> list:
    """The items of every saved Excel report whose period starts on or after `start` (Settings), the latest version
    of each report; an item found in several reports (a week and its month) is listed once, from the latest report.
    read(record) -> its items (downloaded and cached by the caller)."""
    items = {}
    for r in sorted(latest_reports(records, start), key=lambda r: pd.Timestamp(r.modified) if r.modified is not None else pd.Timestamp(0)):
        for it in read(r):
            items[it["key"]] = it                       # a later report replaces the same item from an earlier one
    return sorted(items.values(), key=lambda i: (i["reported"], i["bank"], i["type"], i["invoice"]))


# ---- the items of each report, kept on this computer (like the dashboard's figures) --------------------------------
def cache_load(path: Path) -> dict:
    try:
        with open(path, "rb") as fh:
            saved = pickle.load(fh)
        return saved["items"] if saved.get("version") == CACHE_VERSION else {}
    except Exception:
        return {}


def cache_save(path: Path, entries: dict, keep=None) -> None:
    data = cache_load(path)
    data.update(entries)
    if keep is not None:
        data = {k: v for k, v in data.items() if k in set(keep)}
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        tmp = path.with_suffix(".tmp")
        with open(tmp, "wb") as fh:
            pickle.dump({"version": CACHE_VERSION, "items": data}, fh)
        tmp.replace(path)
    except OSError:
        pass


# ---- the decisions, shared on Google Drive -------------------------------------------------------------------------
def load_decisions(client, root_id: str) -> tuple:
    """({item key: {decision, comment, by, at}}, the file's Drive ID or None when there is none yet)."""
    hits = [f for f in client.list_children(root_id, name=STORE_NAME) if f.name == STORE_NAME]
    if not hits:
        return {}, None
    f = hits[0]
    try:
        data = json.loads(client.download(f.id).decode("utf-8"))
    except (ValueError, UnicodeDecodeError):
        data = {}
    return dict(data.get("decisions") or {}), f.id


def save_decisions(client, root_id: str, changes: dict, who: str) -> dict:
    """Write the changes ({key: {"decision", "comment"}}) on top of the latest decisions on Drive (read again just
    before, so a decision saved meanwhile by someone else is kept). An empty decision removes the item's decision.
    Returns every decision."""
    current, file_id = load_decisions(client, root_id)
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    for key, ch in changes.items():
        if not ch.get("decision"):
            current.pop(key, None)
        else:
            current[key] = {"decision": ch["decision"], "comment": ch.get("comment", ""), "by": who or "unknown",
                            "at": now}
    body = json.dumps({"version": 1, "decisions": current}, indent=1, ensure_ascii=False).encode("utf-8")
    client.upload(root_id, STORE_NAME, body, "application/json", replace_id=file_id)
    return current


# ---- figures --------------------------------------------------------------------------------------------------------
def days_open(item: dict, today=None) -> int:
    """Days since the report that listed the item was made."""
    today = pd.Timestamp(today if today is not None else pd.Timestamp.now()).normalize()
    return max(0, int((today - pd.Timestamp(item["reported"]).normalize()).days))


def summary(items: list, decisions: dict, today=None, days: int = 7) -> dict:
    """{"total", "resolved", "open", "overdue", "open_amount": {currency: amount}, "overdue_amount", "oldest",
    "received_later", "by_type": {type: (open, total)}}."""
    open_items = [i for i in items if not decisions.get(i["key"], {}).get("decision")]
    overdue = [i for i in open_items if days_open(i, today) > days]
    amount = lambda rows: {c: sum(i["amount"] for i in rows if i["currency"] == c)
                           for c in sorted({i["currency"] for i in rows}, key=lambda c: (c != "RWF", c))}
    return {"total": len(items), "resolved": len(items) - len(open_items), "open": len(open_items),
            "overdue": len(overdue), "open_amount": amount(open_items), "overdue_amount": amount(overdue),
            "oldest": max((days_open(i, today) for i in open_items), default=0),
            "received_later": sum(1 for i in items if i["type"] == "Pending"
                                  and decisions.get(i["key"], {}).get("decision") == "Received later"),
            "by_type": {t: (sum(1 for i in open_items if i["type"] == t), sum(1 for i in items if i["type"] == t))
                        for t in DECISIONS}}


def in_period(items: list, start, end, banks=None) -> list:
    """The items paid between start and end (both included), of some banks (None = all)."""
    start, end = pd.Timestamp(start), pd.Timestamp(end)
    return [i for i in items if i["paid_on"] is not None and start <= i["paid_on"] <= end
            and (banks is None or i["bank"] in banks)]


def money(amounts: dict) -> str:
    return " · ".join(f"{c} {v:,.0f}" for c, v in amounts.items() if v) or "–"


def reminder_text(items: list, decisions: dict, days: int = 7, today=None) -> str:
    """The Slack reminder: the items still open for more than `days` days (each listed), or that none is."""
    s = summary(items, decisions, today, days)
    over = [i for i in items if not decisions.get(i["key"], {}).get("decision") and days_open(i, today) > days]
    if not over:
        return (f"✅ Follow-up: no item open for more than {days} days "
                f"({s['open']} open, {s['resolved']} of {s['total']} resolved).")
    lines = [f"⚠️ {len(over)} item{'s' if len(over) != 1 else ''} open for more than {days} days: "
             f"{money(s['overdue_amount'])}", ""]
    for i in sorted(over, key=lambda i: -days_open(i, today))[:15]:
        name = engine.bank_name(i["bank"])
        lines.append(f"• {name} · {i['type']} · invoice {i['invoice'] or 'none'} · {i['currency']} {i['amount']:,.0f}"
                     f" · open {days_open(i, today)} days")
    if len(over) > 15:
        lines.append(f"… and {len(over) - 15} more (Follow-up page)")
    lines += ["", f"Follow-up: {s['resolved']} of {s['total']} resolved · {s['open']} still open."]
    return "\n".join(lines)
