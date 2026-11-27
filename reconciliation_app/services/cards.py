# """
# CARD PAYMENTS (25/11/2026): IREMBOPAY'S CARD PAYMENTS AGAINST THE CARD GATEWAYS' OWN STATEMENTS
# ================================================================================================
# Card payments are reconciled with the gateways' transaction exports, not the bank statement (team decision,
# 25/11/2026): CyberSource (BK: bokc00052 / bokc00053; also irembo1, rw_irembo_3ds) and MPGS (BK: BOK000181, BOK000182).

#     read_irembopay(df)        the IremboPay card export (the usual _SELECT_pi_invoice_number... columns)
#     read_gateway(df)          a gateway export, recognised by its column titles -> one row per attempt
#     reconcile(pay, gw, a, b)  the invoices paid from day a to day b (Kigali time): successful, failed (with the
#                               error in plain words), charged but not paid in IremboPay, paid but not charged, amounts
#     excel_bytes(res)          the report (.xlsx)    ·    slack_text(res)  the summary to paste in Slack

# Link: invoice number = CyberSource "Merchant Reference Number" (or "No") = MPGS "Order ID" before the "-".
# Times: CyberSource Transaction Search is in SAST (= Kigali), Secure Acceptance in its own offset (+08:00), MPGS and
# IremboPay in Kigali time; everything is turned into Kigali time.
# """
# from __future__ import annotations

# import io
# import re

# import pandas as pd

# KIGALI = "Africa/Kigali"

# # CyberSource reason codes, in plain words (CyberSource's list)
# CS_REASONS = {
#     "100": "Success", "101": "Missing payment details", "102": "Invalid payment details",
#     "150": "CyberSource system error", "151": "Timeout at CyberSource", "152": "Timeout at CyberSource",
#     "200": "Address check failed", "201": "The bank asks for a voice call", "202": "Expired card",
#     "203": "Declined by the card's bank", "204": "Insufficient funds", "205": "Lost or stolen card",
#     "207": "The card's bank could not be reached", "208": "Card not allowed for online payments",
#     "210": "Credit limit reached", "211": "Wrong CVV", "220": "Declined by the processor",
#     "221": "Customer on a blocked list", "222": "Account frozen", "230": "Wrong CVV",
#     "231": "Invalid card number", "232": "Card type not accepted", "233": "Declined by the processor",
#     "234": "Merchant set-up problem", "236": "Processor failure", "240": "Card type does not match the number",
#     "250": "Timeout at the processor", "251": "Card usage limit reached", "254": "Refund not allowed",
#     "475": "3-D Secure not completed by the customer", "476": "3-D Secure check failed",
#     "480": "Held for review by the fraud rules", "481": "Rejected by the fraud rules",
#     "520": "Declined by the fraud rules"}
# # MPGS: the acquirer's (ISO) response codes
# MPGS_CODES = {"0": "Approved", "1": "Refer to the card's bank", "3": "Invalid merchant", "4": "Card to be kept",
#               "5": "Declined by the card's bank", "12": "Invalid transaction", "13": "Invalid amount",
#               "14": "Invalid card number", "51": "Insufficient funds", "54": "Expired card", "55": "Wrong PIN",
#               "57": "Not allowed for this card", "58": "Not allowed for this terminal", "59": "Suspected fraud",
#               "61": "Card usage limit reached", "62": "Restricted card", "63": "Security violation",
#               "65": "Card usage limit reached", "82": "Wrong CVV", "91": "The card's bank could not be reached"}
# GATEWAYS = {"cs_search": "CyberSource", "cs_secure": "CyberSource", "mpgs": "MPGS"}
# NO_RECORD = "No record in the IremboPay file"
# PAID_STATUSES = {"PAID", "PAYOUT_INITIATED", "SETTLED"}
# # 25/11/2026 (Yvan): an invoice paid while its card tries failed was paid another way, not through this gateway
# AFTERWARDS = ["Paid another way", "Not paid", NO_RECORD]


# def afterwards(payment_status, transaction_status) -> str:
#     """What became of an invoice whose card payment failed on the gateway."""
#     if transaction_status == NO_RECORD:
#         return NO_RECORD
#     return "Paid another way" if payment_status in PAID_STATUSES else "Not paid"
# ATTEMPT_COLS = ["gateway", "account", "time", "invoice", "amount", "currency", "reference", "card", "status",
#                 "code", "error"]


# def _num(s: pd.Series) -> pd.Series:
#     return pd.to_numeric(s.astype(str).str.replace(",", "").str.strip(), errors="coerce")


# def read_irembopay(df: pd.DataFrame) -> pd.DataFrame:
#     """The IremboPay card export: one row per payment attempt, its route from provider_reference."""
#     d = df.copy()
#     d.columns = [c.strip() for c in d.columns]
#     d["invoice_number"] = d["invoice_number"].astype(str).str.strip()
#     d["amount"] = _num(d["amount"])
#     d["paid_at"] = pd.to_datetime(d["paid_at"], errors="coerce")
#     ref = d["provider_reference"].fillna("").astype(str)
#     d["route"] = ref.map(lambda r: "MPGS" if r.startswith("SESSION") else "CyberSource" if re.fullmatch(r"\d{22}", r)
#                          else "Other")
#     return d


# def kind_of(df: pd.DataFrame) -> str | None:
#     cols = set(c.strip() for c in df.columns)
#     if {"CyberSource Merchant ID", "Merchant Reference Number", "Applications", "Reason Code"} <= cols:
#         return "cs_search"
#     if {"Merchant Reference No", "Decision", "Reason Code", "Request ID"} <= cols:
#         return "cs_secure"
#     if {"Order ID", "Transaction Gateway Response Code", "Transaction Type"} <= cols:
#         return "mpgs"
#     return None


# def read_gateway(df: pd.DataFrame) -> pd.DataFrame:
#     """A gateway export -> one row per attempt: status Charged / Reversed / Failed / Step (a 3-D Secure step)."""
#     df = df.rename(columns=lambda c: c.strip())
#     kind = kind_of(df)
#     if kind is None:
#         raise ValueError("Not a CyberSource or MPGS export (its column titles are not recognised)")
#     out = pd.DataFrame(index=df.index)
#     if kind == "cs_search":
#         apps = df["Applications"].fillna("")
#         out["account"] = df["CyberSource Merchant ID"]
#         out["time"] = pd.to_datetime(df["Date and Time"].str.replace(r"\s+[A-Z]{3,4}$", "", regex=True),
#                                      format="%b %d %Y %I:%M:%S %p", errors="coerce")       # SAST = Kigali time
#         out["invoice"], out["reference"] = df["Merchant Reference Number"], df["Request ID"]
#         out["card"] = df["Payment Method"].fillna("") + " " + df["Account Suffix"].fillna("").astype(str)
#         code = df["Reason Code"].fillna("").astype(str)
#         only_reversal = ~apps.str.replace(r"Card Payments - Full Authorization Reversal\(\w+\)", "", regex=True) \
#             .str.strip(", ").astype(bool)
#         status = pd.Series("Failed", index=df.index)
#         status[code.isin(["100", "475"])] = "Step"     # a 3-D Secure or fraud-check step, the payment not run yet
#         status[apps.str.contains(r"Authorization\(Success\)")] = "Charged"
#         status[only_reversal & apps.str.contains(r"Reversal\(Success\)")] = "Reversed"
#         status[only_reversal & ~apps.str.contains(r"Reversal\(Success\)")] = "Ignore"     # a reversal that failed
#         out["status"], out["code"] = status, code
#         out["error"] = code.map(lambda c: CS_REASONS.get(c, f"Reason code {c}"))
#     elif kind == "cs_secure":
#         out["account"] = df["Merchant ID"]
#         t = pd.to_datetime(df["Date and Time"], format="%m-%d-%y %H:%M:%S %z", errors="coerce", utc=True)
#         out["time"] = t.dt.tz_convert(KIGALI).dt.tz_localize(None)
#         out["invoice"], out["reference"] = df["Merchant Reference No"], df["Request ID"]
#         out["card"] = df["Account Suffix"].fillna("").astype(str)
#         code = df["Reason Code"].fillna("").astype(str)
#         out["status"] = df["Decision"].str.upper().map({"ACCEPT": "Charged"}).fillna("Failed")
#         out["code"], out["error"] = code, code.map(lambda c: CS_REASONS.get(c, f"Reason code {c}"))
#     else:
#         pay = df["Transaction Type"].eq("Payment")
#         auth_only = df["Transaction Type"].eq("Authentication") & ~df["Order ID"].isin(df.loc[pay, "Order ID"])
#         df = df[pay | auth_only]                       # a 3-D Secure check counts only when no payment followed it
#         out = out.loc[df.index]
#         out["account"] = df["Merchant ID"]
#         out["time"] = pd.to_datetime(df["Transaction Date"], errors="coerce")
#         out["invoice"] = df["Order ID"].astype(str).str.split("-").str[0]
#         out["reference"] = df["Transaction ID"]
#         out["card"] = df["Payment Method"].fillna("") + " " + df["Account Identifier"].fillna("").astype(str).str[-4:]
#         resp = df["Transaction Gateway Response Code"].fillna("")
#         acq = _num(df["Acquirer Response Code"]).astype("Int64").astype(str).replace("<NA>", "")
#         approved = df["Transaction Type"].eq("Payment") & resp.eq("Approved")
#         out["status"] = approved.map({True: "Charged", False: "Failed"})
#         auth = df["Payer Authentication -  Authentication Status"].fillna("") if \
#             "Payer Authentication -  Authentication Status" in df else pd.Series("", index=df.index)
#         out["code"] = acq.where(acq != "", resp)
#         out["error"] = [("3-D Secure: " + a if t == "Authentication" else
#                          (MPGS_CODES.get(c, resp_) if c else resp_))
#                         for t, a, c, resp_ in zip(df["Transaction Type"], auth, acq, resp)]
#         out.loc[approved, "error"] = "Success"
#     amount_col = {"cs_search": "Amount", "cs_secure": "Amount", "mpgs": "Transaction Amount (amount only)"}[kind]
#     ccy_col = {"cs_search": "Currency", "cs_secure": "Currency", "mpgs": "Transaction Amount (currency only)"}[kind]
#     out["amount"], out["currency"] = _num(df[amount_col]), df[ccy_col]
#     out["gateway"] = GATEWAYS[kind]
#     out["invoice"] = out["invoice"].astype(str).str.strip()
#     return out[ATTEMPT_COLS].reset_index(drop=True)


# def _sum(df: pd.DataFrame, col="amount") -> dict:
#     return {c: float(v) for c, v in df.groupby("currency")[col].sum().items()}


# def reconcile(pay: pd.DataFrame, gw: pd.DataFrame, start, end, accounts=None) -> dict:
#     """The invoices with an attempt from day `start` to day `end` (inclusive, Kigali time), on the gateway accounts
#     given (all when None). The two sides are compared only where both files reach: a payment after the end of the
#     IremboPay file, or before the start of the gateway files, is not "missing"."""
#     a, b = pd.Timestamp(start).normalize(), pd.Timestamp(end).normalize() + pd.Timedelta(days=1)
#     gw = gw[gw.status != "Ignore"]
#     others = gw[~gw.account.isin(accounts)] if accounts is not None else gw.iloc[0:0]
#     gw = gw[gw.account.isin(accounts)] if accounts is not None else gw
#     routes = set(gw.gateway)
#     p = pay[pay.route.isin(routes)]
#     # IremboPay payments of the accounts not chosen (same gateway, e.g. bokc00053 when only bokc00052 is chosen)
#     p = p[~p.provider_reference.isin(others.reference) & ~p.invoice_number.isin(others.invoice[others.status == "Charged"])]
#     first, last = gw.time.min(), gw.time.max()
#     lo = max(a, first, p.paid_at.min() if len(p) else first)
#     hi = min(b, last + pd.Timedelta(seconds=1), p.paid_at.max() + pd.Timedelta(seconds=1) if len(p) else b)
#     att = gw[(gw.time >= a) & (gw.time < b)].copy()
#     p_win = p[(p.paid_at >= lo) & (p.paid_at < hi)]
#     # the gateway side, per invoice: charges less reversals
#     g = att.groupby("invoice")
#     inv = pd.DataFrame({"charges": g.status.apply(lambda s: (s == "Charged").sum()),
#                         "reversals": g.status.apply(lambda s: (s == "Reversed").sum()),
#                         "gateway": g.gateway.first(), "account": g.account.last(), "last_time": g.time.max()})
#     inv["net"] = inv.charges - inv.reversals
#     charged = att[att.status == "Charged"].sort_values("time").groupby("invoice").last()
#     inv = inv.join(charged[["amount", "currency", "reference", "time", "card"]])
#     # IremboPay, per invoice
#     ok = p[p.transaction_status.eq("SUCCESSFUL")].sort_values("paid_at").groupby("invoice_number").last()
#     info = p.sort_values("paid_at").groupby("invoice_number").last()

#     def ir(invoice, col):
#         return ok[col].get(invoice) if invoice in ok.index else None

#     rows_ok, rows_charged_not_paid, rows_amount, rows_reversed = [], [], [], []
#     for invoice, r in inv[inv.net > 0].iterrows():
#         row = {"Invoice": invoice, "Gateway": r.gateway, "Account": r.account, "Date and time": r.time,
#                "Amount": r.amount, "Currency": r.currency, "Card": r.card, "Gateway reference": r.reference,
#                "Merchant": info["name"].get(invoice, ""), "Times charged": int(r.net)}
#         if invoice not in ok.index:
#             if not lo <= r.time < hi:                 # outside the IremboPay file: charged, not compared
#                 row["Invoice status"] = "Not compared (outside the IremboPay file)"
#                 rows_ok.append(row)
#                 continue
#             row["transaction_status"] = info["transaction_status"].get(invoice, NO_RECORD)
#             rows_charged_not_paid.append(row)
#             continue
#         row.update({"IremboPay amount": ir(invoice, "amount"), "IremboPay currency": ir(invoice, "currency"),
#                     "IremboPay paid at": ir(invoice, "paid_at"), "Invoice status": ir(invoice, "payment_status")})
#         (rows_ok if (row["IremboPay amount"] == r.amount and row["IremboPay currency"] == r.currency)
#          else rows_amount).append(row)
#     approved = att[att.status == "Charged"].sort_values("time").groupby("invoice").last()
#     for invoice, r in inv[(inv.charges > 0) & (inv.net <= 0)].iterrows():
#         why = approved.error.get(invoice, "")             # e.g. the fraud rules rejected it after the bank approved
#         ts, ps = info["transaction_status"].get(invoice, NO_RECORD), info["payment_status"].get(invoice, "")
#         rows_reversed.append({"Invoice": invoice, "Gateway": r.gateway, "Account": r.account,
#                               "Approved at": approved.time.get(invoice), "Amount": approved.amount.get(invoice),
#                               "Currency": approved.currency.get(invoice), "Card": approved.card.get(invoice),
#                               "Why reversed": why if why and why != "Success" else "Reversed after approval",
#                               "Reason code": approved.code.get(invoice, ""), "Times approved": int(r.charges),
#                               "Merchant": info["name"].get(invoice, ""), "transaction_status": ts,
#                               "Invoice status": ps, "Invoice afterwards": afterwards(ps, ts)})
#     # failed: invoices never charged in the period, with their last error; and every failed attempt
#     amounts = p.groupby("invoice_number").agg(amount=("amount", "last"), currency=("currency", "last"))
#     fails = att[att.status.isin(["Failed", "Step"])].sort_values("time")
#     rows_failed = []
#     for invoice, r in inv[inv.net <= 0].iterrows():
#         f = fails[fails.invoice == invoice]
#         real = f[f.status == "Failed"]
#         if r.charges > 0:                              # approved by the bank, then reversed (e.g. the fraud rules)
#             final = approved.loc[invoice]
#             error = (final.error if final.error != "Success" else "Reversed") + " (approved, then reversed)"
#             code, tries_n = final.code, int(r.charges)
#         elif len(f):
#             final = (real if len(real) else f).iloc[-1]
#             error = final.error if len(real) else "Not finished by the customer" if final.code == "100" else final.error
#             code, tries_n = (final.code if len(real) else ""), int(len(real))
#         else:
#             continue
#         known = att[(att.invoice == invoice) & att.amount.notna()]      # 3-D Secure rows carry no amount
#         amt, ccy = (amounts.amount.get(invoice), amounts.currency.get(invoice)) if invoice in amounts.index else \
#             (known.amount.iloc[-1], known.currency.iloc[-1]) if len(known) else (None, None)
#         rows_failed.append({"Invoice": invoice, "Gateway": r.gateway, "Account": r.account, "Last try": final.time,
#                             "Amount": amt, "Currency": ccy, "Card": final.card, "Error": error,
#                             "Reason code": code, "Failed tries": tries_n,
#                             "Merchant": info["name"].get(invoice, ""),
#                             "transaction_status": info["transaction_status"].get(invoice, NO_RECORD),
#                             "Invoice status": info["payment_status"].get(invoice, ""),
#                             "Invoice afterwards": afterwards(info["payment_status"].get(invoice, ""),
#                                                              info["transaction_status"].get(invoice, NO_RECORD))})
#     tries = fails[fails.status == "Failed"].assign(
#         **{"Paid later": lambda x: x.invoice.isin(inv[inv.net > 0].index).map({True: "Yes", False: "No"})})
#     tries = tries.rename(columns={"invoice": "Invoice", "gateway": "Gateway", "account": "Account",
#                                   "time": "Date and time", "amount": "Amount", "currency": "Currency", "card": "Card",
#                                   "error": "Error", "code": "Reason code", "reference": "Gateway reference"})
#     tries = tries[["Invoice", "Gateway", "Account", "Date and time", "Amount", "Currency", "Card", "Error",
#                    "Reason code", "Paid later", "Gateway reference"]]
#     # IremboPay SUCCESSFUL in the period, not charged at the gateway
#     paid_not_charged = p_win[p_win.transaction_status.eq("SUCCESSFUL") & ~p_win.invoice_number.isin(inv[inv.net > 0].index)]
#     paid_not_charged = paid_not_charged.rename(columns={
#         "invoice_number": "Invoice", "route": "Gateway", "paid_at": "IremboPay paid at", "amount": "Amount",
#         "currency": "Currency", "provider_reference": "IremboPay provider reference", "name": "Merchant",
#         "payment_status": "Invoice status"})[["Invoice", "Gateway", "IremboPay paid at", "Amount", "Currency",
#                                               "IremboPay provider reference", "Merchant", "Invoice status"]]
#     frames = {"Successful": pd.DataFrame(rows_ok), "Failed": pd.DataFrame(rows_failed), "Failed tries": tries,
#               "Charged, not paid in IremboPay": pd.DataFrame(rows_charged_not_paid),
#               "Paid in IremboPay, not charged": paid_not_charged.reset_index(drop=True),
#               "Different amount": pd.DataFrame(rows_amount), "Charged then reversed": pd.DataFrame(rows_reversed)}
#     for k, v in frames.items():
#         if v.empty:
#             frames[k] = pd.DataFrame(columns=["Invoice", "Gateway", "Account", "Amount", "Currency"])
#     succ = frames["Successful"]
#     summary = []
#     for (gwy, acct), s in succ.groupby(["Gateway", "Account"]) if len(succ) else []:
#         for ccy, x in s.groupby("Currency"):
#             summary.append({"Gateway": gwy, "Account": acct, "Currency": ccy, "Successful": len(x),
#                             "Amount": float(x.Amount.sum())})
#     errors = frames["Failed"].groupby("Error").size().sort_values(ascending=False) if len(rows_failed) else pd.Series(dtype=int)
#     after = frames["Failed"]["Invoice afterwards"].value_counts() if len(rows_failed) else pd.Series(dtype=int)
#     return {"start": a, "end": b - pd.Timedelta(days=1), "covered": (first, last), "compared": (lo, hi),
#             "irembopay_covers": (p.paid_at.min(), p.paid_at.max()) if len(p) else (None, None), "frames": frames,
#             "summary": pd.DataFrame(summary), "errors": errors,
#             "afterwards": {k: int(after.get(k, 0)) for k in AFTERWARDS},
#             "successful": len(succ), "successful_amounts": _sum(succ.rename(columns={"Amount": "amount", "Currency": "currency"})) if len(succ) else {},
#             "failed": len(rows_failed), "failed_tries": len(tries),
#             "issues": len(rows_charged_not_paid) + len(paid_not_charged) + len(rows_amount),
#             "accounts": sorted(set(att.account.dropna()))}


# def _money(amounts: dict) -> str:
#     return " + ".join(f"{c} {v:,.2f}".replace(".00", "") for c, v in sorted(amounts.items())) or "0"


# def slack_text(res: dict) -> str:
#     f = lambda t: t.strftime("%d/%m/%Y")
#     lines = [f"*Card payments · {', '.join(res['accounts'])} · {f(res['start'])} – {f(res['end'])}*",
#              f"✅ Successful: *{res['successful']:,}* ({_money(res['successful_amounts'])})",
#              f"❌ Failed on the gateway: *{res['failed']:,}* invoices ("
#              + " · ".join(f"{k[0].lower() + k[1:]} {n:,}" for k, n in res["afterwards"].items() if n)
#              + f") · failed tries in all: {res['failed_tries']:,}"]
#     if len(res["errors"]):
#         lines.append("Errors: " + " · ".join(f"{e} {n}" for e, n in res["errors"].items()))
#     fr = res["frames"]
#     lines.append(f"⚠️ Charged, not paid in IremboPay: {len(fr['Charged, not paid in IremboPay'])} · "
#                  f"paid in IremboPay, not charged: {len(fr['Paid in IremboPay, not charged'])} · "
#                  f"different amount: {len(fr['Different amount'])}")
#     lines.append("_Paid another way: the card failed on this gateway, the invoice was paid on another channel._")
#     first, last = res["covered"]
#     if res["start"] < first.normalize() or res["end"] + pd.Timedelta(days=1) > last:
#         lines.append(f"_The gateway files cover {first:%d/%m %H:%M} – {last:%d/%m %H:%M} only._")
#     i_first, i_last = res["irembopay_covers"]
#     if i_last is not None and res["end"] + pd.Timedelta(days=1) > i_last + pd.Timedelta(hours=1):
#         lines.append(f"_The IremboPay file ends {i_last:%d/%m %H:%M}: later payments are not compared._")
#     return "\n".join(lines)


# CURRENCY_ORDER = ["RWF", "USD", "EUR"]
# NEED_CHECKING = ["Charged, not paid in IremboPay", "Paid in IremboPay, not charged", "Different amount"]


# def _currencies(*frames) -> list:
#     found = set().union(*(set(f["Currency"].dropna()) for f in frames if "Currency" in f))
#     return [c for c in CURRENCY_ORDER if c in found] + sorted(found - set(CURRENCY_ORDER))


# def summary_tables(res: dict) -> dict:
#     """The Summary sheet's tables, built like the banks' (one row per currency + a total row, count only):
#     {title: table}. % = share of the invoices tried in the same currency."""
#     ok, failed = res["frames"]["Successful"], res["frames"]["Failed"]
#     curs = _currencies(ok, failed)
#     tried = {c: (ok.Currency == c).sum() + (failed.Currency == c).sum() for c in curs}

#     def table(part, with_pct=True):
#         rows = [{"Currency": c, "Number of Transactions": int((part.Currency == c).sum()),
#                  "Total Amount": float(part.Amount[part.Currency == c].sum()),
#                  "%": (part.Currency == c).sum() / tried[c] if tried[c] else 0} for c in curs]
#         unknown = int(part.Currency.isna().sum()) if "Currency" in part else 0
#         if unknown:                                  # a try stopped before the card step: no amount recorded
#             rows.append({"Currency": "Not recorded", "Number of Transactions": unknown, "Total Amount": None, "%": None})
#         all_tried = len(ok) + len(failed)
#         rows.append({"Currency": None, "Number of Transactions": len(part), "Total Amount": None,
#                      "%": len(part) / all_tried if all_tried else 0})
#         out = pd.DataFrame(rows)
#         return out if with_pct else out.drop(columns="%")

#     errors = res["errors"]
#     err = pd.DataFrame({"Error": list(errors.index) + [None], "Number of Transactions": list(errors.values) + [int(errors.sum())],
#                         "%": [n / errors.sum() for n in errors.values] + [1.0]}) if len(errors) else \
#         pd.DataFrame({"Error": [None], "Number of Transactions": [0], "%": [0.0]})
#     checks = pd.DataFrame({"Check": NEED_CHECKING + [None],
#                            "Number of Transactions": [len(res["frames"][k]) for k in NEED_CHECKING]
#                            + [sum(len(res["frames"][k]) for k in NEED_CHECKING)]})
#     n_after = res["afterwards"]
#     total = sum(n_after.values())
#     after = pd.DataFrame({"Invoice afterwards": AFTERWARDS + [None],
#                           "Number of Transactions": [n_after[k] for k in AFTERWARDS] + [total],
#                           "%": [n_after[k] / total if total else 0 for k in AFTERWARDS] + [1.0 if total else 0]})
#     return {"Overall card payments tried": table(pd.concat([ok, failed]), with_pct=False),
#             "Successful payments": table(ok), "Failed payments": table(failed),
#             "Failed payments by error": err, "Failed payments – the invoice afterwards": after,
#             "Need checking": checks}


# def merchant_table(res: dict) -> pd.DataFrame:
#     """Successful payments per merchant, like the banks' Payment per merchant: one band per currency."""
#     ok = res["frames"]["Successful"]
#     rows = []
#     for cur in _currencies(ok):
#         sub = ok[ok.Currency == cur]
#         name = sub.Merchant.fillna("").astype(str).str.strip().replace("", "(no merchant)")
#         g = sub.assign(m=name).groupby("m").agg(n=("m", "size"), amt=("Amount", "sum")).reset_index() \
#             .sort_values(["n", "m"], ascending=[False, True])
#         rows += [{"Merchant name": m, "Currency": cur, "Number of Transactions": int(n), "Total Amount": amt,
#                   "%": n / len(sub)} for m, n, amt in zip(g.m, g.n, g.amt)]
#         rows.append({"Merchant name": "All", "Currency": cur, "Number of Transactions": len(sub),
#                      "Total Amount": float(sub.Amount.sum()), "%": 1.0})
#     return pd.DataFrame(rows or [{"Merchant name": "All", "Currency": "RWF", "Number of Transactions": 0,
#                                   "Total Amount": 0.0, "%": 0.0}])


# def account_table(res: dict) -> pd.DataFrame:
#     s = res["summary"]
#     if s.empty:
#         return pd.DataFrame({"Account": [None], "Currency": [None], "Number of Transactions": [0], "Total Amount": [0.0]})
#     return s.rename(columns={"Successful": "Number of Transactions", "Amount": "Total Amount"})[
#         ["Account", "Currency", "Number of Transactions", "Total Amount"]]


# def _style_detail(ws):
#     """A detail sheet like the banks': bold shaded boxed header, frozen, filters, widths, number and date formats."""
#     from openpyxl.utils import get_column_letter
#     import engine
#     for cell in ws[1]:
#         cell.font, cell.fill, cell.border = engine.BOLD, engine.HEADER_FILL, engine.BOX
#     ws.freeze_panes = "A2"
#     ws.auto_filter.ref = ws.dimensions
#     for col in ws.columns:
#         header = str(col[0].value or "")
#         width = max(len(str(c.value)) if c.value is not None else 0 for c in col[:200])
#         ws.column_dimensions[get_column_letter(col[0].column)].width = min(max(width + 2, 10), 45)
#         for cell in col[1:]:
#             if "Amount" in header or header == "IremboPay amount":
#                 cell.number_format = "#,##0.00"
#             elif any(w in header for w in ("Date", "time", "try", "paid at", "Charged at")):
#                 cell.number_format = "dd/mm/yyyy hh:mm:ss"


# SOURCE_SHEETS = {"cs_search": "CyberSource file", "cs_secure": "CyberSource file", "mpgs": "MPGS file"}


# def source_sheets(irembopay: list, gateways: list) -> dict:
#     """The files used, as uploaded (25/11/2026): {sheet name: rows} - the IremboPay file(s), then the gateway files of
#     each kind together (CyberSource / MPGS); the same file uploaded twice is kept once."""
#     out = {"IremboPay file": pd.concat(irembopay).drop_duplicates()} if irembopay else {}
#     for df in gateways:
#         name = SOURCE_SHEETS.get(kind_of(df))
#         if name:
#             out[name] = pd.concat([out[name], df]).drop_duplicates() if name in out else df
#     return out


# def excel_bytes(res: dict, made_on=None, sources=None) -> bytes:
#     """The report, laid out like the banks' reports: the Summary (title, period, reconciled on, the tables in two
#     columns, then per account and per merchant), one detail sheet per list, then the files used (sources:
#     {sheet name: rows}, see source_sheets)."""
#     from openpyxl.styles import Font
#     import engine
#     made_on = pd.Timestamp(made_on) if made_on is not None else pd.Timestamp.now()
#     buf = io.BytesIO()
#     with pd.ExcelWriter(buf, engine="openpyxl") as xw:
#         pd.DataFrame().to_excel(xw, sheet_name="Summary", index=False)
#         for name, df in list(res["frames"].items()) + list((sources or {}).items()):
#             if df.empty:                              # an empty list gets no sheet (its 0 is on the Summary)
#                 continue
#             df.to_excel(xw, sheet_name=name[:31], index=False)
#             _style_detail(xw.sheets[name[:31]])
#         ws = xw.sheets["Summary"]
#         for col, w in engine.GS_WIDTHS.items():
#             ws.column_dimensions[col].width = w
#         accounts = ", ".join(res["accounts"])
#         gateways = " / ".join(sorted(set(res["frames"]["Successful"].get("Gateway", pd.Series(dtype=str)).dropna())
#                                      | set(res["frames"]["Failed"].get("Gateway", pd.Series(dtype=str)).dropna())))
#         ws.cell(1, 1, f"Card payments – {gateways or 'Card gateways'} ({accounts})").font = Font(size=14, bold=True)
#         ws.cell(1, 6, f"{engine.RECONCILED_LABEL}{made_on:%A %d/%m/%Y}").font = Font(size=12, italic=True)
#         ws.cell(2, 1, f"Reporting period: {res['start']:%d/%m/%Y} to {res['end']:%d/%m/%Y}").font = Font(italic=True)
#         first, last = res["covered"]
#         ws.cell(2, 6, f"Gateway files cover: {first:%d/%m/%Y %H:%M} to {last:%d/%m/%Y %H:%M}").font = Font(italic=True)
#         i_first, i_last = res["irembopay_covers"]
#         if i_last is not None:
#             ws.cell(3, 6, f"IremboPay file covers: {i_first:%d/%m/%Y %H:%M} to {i_last:%d/%m/%Y %H:%M}").font = \
#                 Font(italic=True)
#         tables = summary_tables(res)
#         top = {1: 4, 6: 4}                            # the next free row of the left (A) and right (F) columns
#         for title, col in [("Overall card payments tried", 1), ("Failed payments by error", 6),
#                            ("Successful payments", 1), ("Failed payments", 1),
#                            ("Failed payments – the invoice afterwards", 6), ("Need checking", 6)]:
#             n = engine._write_table(ws, top[col], col, title, tables[title])
#             top[col] += n + 4
#         row = max(top.values())
#         engine._write_table(ws, row, 1, "Successful payments per account", account_table(res))
#         engine._write_table(ws, row, 6, engine.MERCHANT_TITLE, merchant_table(res), bands="Currency")
#     return buf.getvalue()
