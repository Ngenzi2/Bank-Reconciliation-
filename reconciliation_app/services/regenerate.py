"""
REDO A SAVED REPORT WITH THE CURRENT RULES (26/10/2026) - FROM THE REPORT ITSELF
================================================================================
A saved report keeps every credit of its period in its "Payment details" sheet: the statement columns, what IremboPay
said (found, provider, status, paid_at), the SLA, and the decisions of Exception Resolution (review_status). Its five
summary tables can therefore be computed again with the current rules without the original bank and IremboPay files:

    Processed   every credit of the period (Ecobank: every request answered SUCCESS, Cba Rsp Code 000), each invoice once
    Received    in IremboPay, paid (paid_at filled, whatever the status) through this bank
    Duplicate   paid (paid_at), through another provider
    Pending     in IremboPay without paid_at (BPR: also a credit with no invoice number)
    SLA         the processed ones marked "SLA violated" (the SLA rules did not change)

The decisions of Exception Resolution are those recorded in the report (they win, as in the app). Limits: what needs
the original files is not redone - the IremboPay checkout is not run again, and a second credit that was not in the
report (Ecobank kept one request per invoice) cannot be found. Nothing is saved on Google Drive: the result is a file
on this computer.

    redo(path, bank_code, out_path)  -> (old totals, new totals)
"""
from __future__ import annotations

from pathlib import Path

import pandas as pd

import engine

_TABLE_OF = {"Notification received": "Notification received", "Paid with other channel": "Paid with other channel"}


def _review_table(status) -> str:
    """The Report table chosen in Exception Resolution, read back from the report's review_status."""
    s = "" if pd.isna(status) else str(status)
    if s.startswith("Pending notification"):
        return "Pending notification"
    if s.startswith("Excluded"):
        return "Exclude"
    return _TABLE_OF.get(s, "")


def _bool(series) -> pd.Series:
    return series.astype(str).str.strip().str.lower().isin(["true", "1", "yes"])


def _details(path: Path, bank) -> pd.DataFrame:
    c = bank.COLUMNS
    d = pd.read_excel(path, sheet_name="Payment details", dtype={"invoice_number": str}, engine=engine.EXCEL_ENGINE)
    for col in (c["payment_date"], c.get("posted_date"), "paid_at", "created_at", "settled_at"):
        if col and col in d:
            d[col] = pd.to_datetime(d[col], errors="coerce")
    for col in ("found_in_irembopay", "is_extra_credit"):
        d[col] = _bool(d[col]) if col in d else False
    d["invoice_number"] = d["invoice_number"].where(d["invoice_number"].astype(str).str.strip().ne("")
                                                    & d["invoice_number"].notna(), None)
    d["review_table"] = d["review_status"].map(_review_table) if "review_status" in d else ""
    for col in ("payment_provider", "payment_status", "sla_status", "name"):
        if col not in d:
            d[col] = None
    if "Cba Rsp Code" in d:
        d["Cba Rsp Code"] = d["Cba Rsp Code"].map(engine.get_bank("ECO")._code)     # 0 / 0.0 / '000' -> '000'
    d[c["amount"]] = pd.to_numeric(d[c["amount"]], errors="coerce")
    return d


def _masks(rep, bank, a, b):
    paid_on = bank.COLUMNS["payment_date"]
    window = (rep[paid_on] >= a) & (rep[paid_on] < b)
    processed = window & ~rep["is_extra_credit"]
    if bank.CODE == "ECO":
        processed &= rep["Cba Rsp Code"].eq("000")
    received = processed & bank.received(rep)
    duplicate = processed & bank.duplicate(rep) & ~received
    pending = processed & bank.pending(rep) & ~received & ~duplicate
    return engine.with_extra_decisions(rep, {"processed": processed, "received": received,
                                             "sla": processed & rep["sla_status"].eq("SLA violated"),
                                             "pending": pending, "duplicate": duplicate}, window)


def _summary5(rep, bank, m):
    amount = bank.COLUMNS["amount"]
    df = rep[m["processed"]]
    processed = df.groupby("Currency").size()

    def table(part, with_pct=True):
        rows = []
        for cur in bank.CURRENCIES:
            sub = part[part["Currency"] == cur]
            base = processed.get(cur, 0)
            rows.append({"Currency": cur, "Number of Transactions": len(sub), "Total Amount": sub[amount].sum(),
                         "%": len(sub) / base if base else 0})
        rows.append({"Currency": None, "Number of Transactions": len(part), "Total Amount": None,
                     "%": len(part) / len(df) if len(df) else 0})
        out = pd.DataFrame(rows)
        return out if with_pct else out.drop(columns="%")

    return {"Overall processed payments": table(df, with_pct=False),
            "Overall notifications received by Irembo": table(rep[m["received"]]),
            "SLA Violated": table(rep[m["sla"]]),
            "Pending notifications": table(rep[m["pending"]]),
            "Duplicate Payment": table(rep[m["duplicate"]])}


def _merchants(rep, bank, m):
    amount = bank.COLUMNS["amount"]
    part = rep[m["received"]]
    name = part["name"].astype(object).where(part["name"].notna() & part["name"].astype(str).str.strip().ne(""),
                                             "(no merchant)").astype(str).str.strip()
    rows = []
    for cur in list(bank.CURRENCIES) + sorted(set(part["Currency"].dropna()) - set(bank.CURRENCIES)):
        sub = part[part["Currency"] == cur]
        if sub.empty:
            continue
        g = (pd.DataFrame({"m": name[sub.index], "amt": sub[amount]}).groupby("m")
             .agg(n=("m", "size"), amt=("amt", "sum")).reset_index().sort_values(["n", "m"], ascending=[False, True]))
        total = len(sub)
        rows += [{"Merchant name": mm, "Currency": cur, "Number of Transactions": int(n), "Total Amount": amt,
                  "%": n / total} for mm, n, amt in zip(g["m"], g["n"], g["amt"])]
        rows.append({"Merchant name": "All", "Currency": cur, "Number of Transactions": total,
                     "Total Amount": sub[amount].sum(), "%": 1.0})
    if not rows:
        rows = [{"Merchant name": "All", "Currency": bank.CURRENCIES[0], "Number of Transactions": 0,
                 "Total Amount": 0.0, "%": 0.0}]
    return pd.DataFrame(rows)


def redo(path: Path, bank_code: str, kind: str, start: pd.Timestamp, out_path: Path) -> tuple:
    """Write the report again with the current rules (out_path, on this computer) -> (old totals, new totals),
    each {table title: {currency: n, "Total": n}} of the main block."""
    bank = engine.get_bank(bank_code)
    period = engine.make_period(kind, start)
    old = engine.read_report(path, count_reviewed=False)
    old_main = next((b for b in old["blocks"] if b[0] in ("Summary-Monthly", "Summary-Daily")), old["blocks"][-1])
    rep = _details(path, bank)
    aliases = engine.option(bank, "PROVIDER_ALIASES", [])
    if aliases and "payment_channel" in rep:              # BPR: a bank transfer the old rule sent to review and the
        bt = (rep["payment_provider"].eq("OTHER") & rep["payment_channel"].eq("BANK_TRANSFER")   # reviewer put in
              & rep["review_table"].eq("Paid with other channel"))                              # "other channel":
        rep.loc[bt, "payment_provider"] = aliases[0]      # its provider back ('BPR'), and BPR rules apply
        rep.loc[bt, "review_table"] = ""
        rep.loc[bt, "review_status"] = "Received (BPR bank transfer, rule of 27/10/2026)"
    if aliases:                                            # BPR (27/10/2026): a 'BPR' bank transfer is BPR too - its
        from services import app_settings                 # SLA is checked now (it was "Not processed")
        redo_sla = rep["payment_provider"].isin(aliases) & rep["sla_status"].eq("Not processed")
        if redo_sla.any():
            part = rep[redo_sla].assign(sla_override="")
            for col in getattr(bank, "DETAIL_COLUMNS", []):
                if col not in part:
                    part[col] = ""
            diff, status = bank.sla(part, app_settings.sla_days(app_settings.load(), bank))
            rep.loc[redo_sla, "sla_status"] = status
            col = engine.option(bank, "SLA_DIFF_COLUMN", "sla_days_difference")
            if col in rep:
                rep.loc[redo_sla, col] = diff
    rep["notification_received"] = bank.received(rep)
    one_day = pd.Timedelta(days=1)
    s, e = period.start, period.end
    blocks, merchants = [], {}
    if kind == "day":
        m = _masks(rep, bank, s, e)
        blocks.append(("Summary-Daily", f"Day: {s:%d/%m/%Y}", _summary5(rep, bank, m)))
        merchants[blocks[-1][1]] = _merchants(rep, bank, m)
    else:
        if kind == "week":
            weeks = [(s, e)]
        else:
            cuts = [s] + [d for d in pd.date_range(s + one_day, e - one_day) if d.dayofweek == 0]
            weeks = list(zip(cuts, cuts[1:] + [e]))
        for i, (a, b) in enumerate(weeks, 1):
            name = engine.week_name(period) if kind == "week" else f"Week {i}"
            m = _masks(rep, bank, a, b)
            blocks.append(("Summary-Weekly", f"{name}: {a:%d/%m/%Y} - {b - one_day:%d/%m/%Y}", _summary5(rep, bank, m)))
            merchants[blocks[-1][1]] = _merchants(rep, bank, m)
        if kind == "month":
            m = _masks(rep, bank, s, e)
            blocks.append(("Summary-Monthly", f"Month: {s:%B %Y} ({s:%d/%m/%Y} - {period.last_day:%d/%m/%Y})",
                           _summary5(rep, bank, m)))
            merchants[blocks[-1][1]] = _merchants(rep, bank, m)
    whole = _masks(rep, bank, s, e)
    paid_on = bank.COLUMNS["payment_date"]
    rep["Report note"] = engine.table_notes(rep, whole, bank)     # an invoice credited more than once: why, where
    detail = rep.drop(columns=["review_table"])
    try:
        statement = pd.read_excel(path, sheet_name=bank.STATEMENT_SHEET, engine=engine.EXCEL_ENGINE)
    except Exception:
        statement = None
    sheets = {"Pending notification": detail[whole["pending"]].sort_values(paid_on),
              "Duplicate payment other channel": detail[whole["duplicate"]].sort_values(["payment_provider", paid_on])}
    if statement is not None:
        sheets[bank.STATEMENT_SHEET] = statement
    sheets["Payment details"] = detail
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.unlink(missing_ok=True)
    engine._write_report(out_path, blocks, sheets, bank, pd.Timestamp.now(), merchants)
    return engine.counts_of(old_main[2]), engine.counts_of(blocks[-1][2] if kind != "week" else blocks[0][2])
