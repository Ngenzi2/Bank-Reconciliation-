"""
BPR BANK RWANDA (BPR)
=====================
Everything that is specific to BPR (the common rules are in engine.py). The flow is the same as the other banks:
every credit of the statement is checked transaction by transaction in the IremboPay payment file; what cannot be
settled automatically goes to Exception Resolution. Decisions of the user, 16/10/2026 (docs/BPR bank configuration.md).

The BPR statement (an export of BNR.ACCOUNT_STATEMENT_BALANCE, .xlsx; one day or a few days per file)
    BOOKING_DATE · BANK_REFERENCE · DESCRIPTION · DOC ID · VALUE_DATE (date + time, no seconds) · DEBIT (negative)
    · CREDIT · CLOSING_BALANCE. The first row is "Opening balance". The header row is found, not assumed (the data
    sheet is usually "Export Worksheet", once "Sheet1"; the other sheets - lookup lists, the SQL - are ignored).
    The DOC ID title varies (DOC ID, DOCID, Doc ID, doc id, or blank) and some files have no DOC ID or no VALUE_DATE.
    There is no currency column.

Rules
    - Invoice = DOC ID (filled in by the team). A credit without DOC ID: found in IremboPay by its BANK_REFERENCE or
      DESCRIPTION (= provider_reference) or DESCRIPTION (= invoice number) - only when IremboPay has it; otherwise it
      goes straight to Pending notification (not to Exception Resolution).
    - Only the credits are reconciled; the debits (the daily transfer out) are only used for the checks below.
    - Date = VALUE_DATE: the period of a credit and the SLA. A row without VALUE_DATE goes to review; BOOKING_DATE
      only decides which period's review it appears in. When the same credit (BANK_REFERENCE + amount) is also in a
      file with VALUE_DATE, only that copy is kept.
    - Currency = the IremboPay currency of the invoice; a credit not in the IremboPay file (DOC ID unknown, or no DOC
      ID): from the file name ("USD" in it = USD), changeable in the review. Credits without DOC ID never stop the
      reconciliation. The IremboPay check of step 5 looks a DOC ID not in the file
      up on IremboPay's checkout: "Paid" -> Paid with other channel (confirmed by the reviewer).
    - IremboPay payment_provider BPR_RW is BPR, and so is BPR (bank transfer, 27/10/2026); a credit IremboPay shows
      under another provider goes to review.
    - SLA in whole minutes: paid_at's minute against VALUE_DATE's minute, 1 minute allowed (Bank Management).
    - An invoice credited more than once: the credit IremboPay recorded (DESCRIPTION = provider_reference) counts,
      the others go to review. A credit cancelled by a debit of the same description and amount goes to review.
    - Every file is checked row by row against its balances (a file that does not add up is refused), and the files
      against each other: money that came in with no row = a missing payment (shown, not blocking); money that went out with no
      row = a missing debit (shown, not blocking); a file whose balance column is only shifted, while its rows are
      complete, is accepted with a note; a file that starts by repeating the last row(s) of the file before is
      checked after those rows.

CONTENTS
    1. Identity            4. Reading a statement (checked against its own balances)
    2. Statement columns   5. Checks between files (step 3)
    3. Files needed        6. SLA + calculations + what goes to review
"""
import io
import re

import pandas as pd


# =====================================================================================================================
# 1. IDENTITY
# =====================================================================================================================
CODE = "BPR"
NAME = "BPR"
FULL_NAME = "BPR Bank Rwanda"
PROVIDER = "BPR_RW"                       # payment_provider of BPR in the IremboPay payment file
# 27/10/2026: 'BPR' (channel BANK_TRANSFER: a transfer into IremboPay's BPR account) is BPR too - received when paid,
# SLA checked like any BPR payment; the provider IremboPay recorded is kept as it is in Payment details
PROVIDER_ALIASES = ["BPR"]
PROVIDERS = {PROVIDER, *PROVIDER_ALIASES}
CURRENCIES = ["RWF", "USD"]               # taken from IremboPay (the statement has no currency)
FILE_TYPES = ["xlsx", "xls"]
STATEMENT_SHEET = "BPR statement"

COMPLETED_STATUSES = {"SETTLED", "PAID", "PAYOUT_INITIATED"}     # same as I&M
# pending = not paid yet: payment_status NEW (20/10/2026: "expired" is not a payment status - an expired
# invoice was never paid, it stays NEW; the expiry is a note in the report)
PENDING_STATUSES = {"NEW"}


# =====================================================================================================================
# 2. STATEMENT COLUMNS
# =====================================================================================================================
COLUMNS = {
    "posted_date": "BOOKING_DATE",       # the bank's business day
    "payment_date": "VALUE_DATE",        # date and time of the payment: the period and the SLA
    "description": "DESCRIPTION",
    "reference": "BANK_REFERENCE",
    "amount": "CREDIT",                  # credits only
}
INVOICE_COLUMN = "DOC ID"                # the invoice number, filled in by the team
FOUND_BY = "Invoice found by"            # how the invoice of a credit without DOC ID was found in IremboPay
NO_VALUE_DATE = "VALUE_DATE missing"     # "yes" on a row whose file has no VALUE_DATE
CANCELLED = "Cancelled by"               # the debit that cancels a credit (its bank reference)
DETAIL_COLUMNS = [INVOICE_COLUMN, FOUND_BY, NO_VALUE_DATE, CANCELLED]
DATETIME_COLUMNS = ["VALUE_DATE"]
_HEADER = "BOOKING_DATE"
_ALL = ["BOOKING_DATE", "BANK_REFERENCE", "DESCRIPTION", "DOC ID", "VALUE_DATE", "DEBIT", "CREDIT", "CLOSING_BALANCE"]


# =====================================================================================================================
# 3. FILES NEEDED (step 3)
# =====================================================================================================================
MONDAY_AFTER_WEEKEND = False             # BPR books on Saturdays too
DAY_BEFORE_CURRENCIES = []
DAY_AFTER_RECOMMENDED = True             # evening payments are booked the next morning: recommended, never blocking
ONE_FILE_PER_CURRENCY = False            # the currency of a file is not known before the matching
MISSING_DAYS_BLOCK = False               # days with no statement are shown, Run stays possible
REVERSE_CHECK = True                     # step 3: IremboPay BPR_RW payments that are on no statement
CHECKOUT_ALL = True                      # 26/10/2026: every invoice goes through IremboPay's checkout before the run
CURRENCY_FROM_IREMBOPAY = True           # the currency of a credit = the IremboPay currency of its invoice
ALL_PROVIDER_REFERENCES = True           # keep every IremboPay transaction's provider_reference (credited again)
NO_INVOICE_PENDING = True                # a credit with no invoice number (not found): Pending notification, not reviewed
ALWAYS_REVIEW_COLUMNS = ["Cancelled by"]    # ... unless it was cancelled by a debit: that one is still reviewed
SEPARATE_CREDITED_AGAIN = True           # step 5: the extra credits of an invoice in a table of their own

# SLA: minutes (a setting in Bank Management), 1 by default
SLA_UNIT = "minutes"
SLA_DEFAULT = 1
SLA_MAX = 60
SLA_DIFF_COLUMN = "sla_minutes_difference"


# =====================================================================================================================
# 4. READING A STATEMENT (checked against its own balances)
# =====================================================================================================================
def _num(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series.astype(str).str.replace(",", "", regex=False).str.strip(), errors="coerce")


def _clean_text(series: pd.Series) -> pd.Series:
    return series.fillna("").astype(str).str.replace("\xa0", "", regex=False).str.strip()


def _table(content, name, excel_engine):
    """The sheet that holds the statement (the one with a BOOKING_DATE header), as text, with its header row."""
    try:
        book = pd.read_excel(io.BytesIO(content) if isinstance(content, (bytes, bytearray)) else content,
                             sheet_name=None, header=None, dtype=str, engine=excel_engine)
    except Exception:
        raise ValueError(f"{name}: this file cannot be opened as an Excel bank statement.") from None
    for sheet in book.values():
        sheet = sheet.fillna("")
        hits = sheet.index[sheet.apply(lambda r: r.astype(str).str.strip().str.upper().eq(_HEADER).any(), axis=1)]
        if len(hits):
            return sheet, hits[0]
    raise ValueError(f"{name}: this is not a BPR bank statement (no 'BOOKING_DATE' header).")


def _columns(header: list) -> list:
    """The statement's column names, made the same in every file: 'Doc ID' / 'doc id' / 'DOCID' / a blank title
    between DESCRIPTION and VALUE_DATE -> 'DOC ID'."""
    out = []
    for i, h in enumerate(header):
        key = re.sub(r"[^A-Z_]", "", str(h).upper())
        if key == "DOCID" or (not key and out and out[-1] == "DESCRIPTION"):
            out.append(INVOICE_COLUMN)
        elif key in {c.replace(" ", "") for c in _ALL}:
            out.append(next(c for c in _ALL if c.replace(" ", "") == key))
        else:
            out.append(f"_col{i}")
    return out


def read_statement(content: bytes, name: str, excel_engine=None):
    """Return (rows, other_credits, info). rows = the credits (the Irembo payments), with the statement's own column
    names + Currency (blank: taken from IremboPay) + the checks' columns. info: currency (blank), first / last (the
    days covered, from the dates inside), and the balances used by the checks between files.
    Raise ValueError with a plain message when the file is not a BPR statement or does not add up."""
    raw, head = _table(content, name, excel_engine)
    data = raw.loc[head + 1:].copy()
    data.columns = _columns(list(raw.loc[head]))
    data = data[[c for c in data.columns if not c.startswith("_col")]]
    data = data[data.apply(lambda r: r.astype(str).str.strip().ne("").any(), axis=1)]
    missing = [c for c in ("BOOKING_DATE", "DESCRIPTION", "DEBIT", "CREDIT", "CLOSING_BALANCE") if c not in data]
    if missing:
        raise ValueError(f"{name}: the BPR statement is missing the column(s) {', '.join(missing)}.")
    for col in ("BANK_REFERENCE", "DESCRIPTION", INVOICE_COLUMN, "VALUE_DATE"):
        data[col] = _clean_text(data[col]) if col in data else ""
    credit, debit, balance = _num(data["CREDIT"]).fillna(0), _num(data["DEBIT"]).fillna(0), _num(data["CLOSING_BALANCE"])
    opening = data["DESCRIPTION"].str.lower().str.startswith("opening balance")

    # ---- the balances: each row's balance = the row before + its credit + its debit (debits are negative)
    if balance.isna().all():
        raise ValueError(f"{name}: no CLOSING_BALANCE in this file, so it cannot be checked.")
    moves = (credit + debit).where(~opening, 0)
    expected = balance.shift(1) + moves
    breaks = (expected - balance).abs().gt(0.5) & balance.notna() & balance.shift(1).notna()
    if breaks.any():
        at = data.loc[breaks.idxmax()]
        raise ValueError(f"{name}: the balances do not add up at {at['BOOKING_DATE']} {at['BANK_REFERENCE']} "
                         f"(expected {expected[breaks.idxmax()]:,.0f}, the file says {balance[breaks.idxmax()]:,.0f}). "
                         "A row may be missing: ask BPR for the complete export.")
    first_bal = balance[opening].iloc[0] if opening.any() else balance.iloc[0] - moves.iloc[0]

    # ---- a credit cancelled by a debit of the same description and amount (its money did not stay)
    is_credit = credit.gt(0) & ~opening
    is_debit = debit.lt(0) & ~opening
    cancel = pd.Series("", index=data.index)
    debits = data[is_debit & data["DESCRIPTION"].ne("")]
    for i in data.index[is_credit]:
        same = debits[(debits["DESCRIPTION"] == data.at[i, "DESCRIPTION"]) & (-debit[debits.index] == credit[i]).values]
        if len(same):
            cancel[i] = same["BANK_REFERENCE"].iloc[0] or "a debit"

    # ---- the credits kept; VALUE_DATE missing: BOOKING_DATE decides the period, the row goes to review
    rows = data[is_credit].copy()
    rows[NO_VALUE_DATE] = rows["VALUE_DATE"].eq("").map({True: "yes", False: ""})
    rows["VALUE_DATE"] = rows["VALUE_DATE"].where(rows["VALUE_DATE"].ne(""), rows["BOOKING_DATE"])
    rows[CANCELLED] = cancel[rows.index]
    rows["Currency"] = ""
    rows = rows[[c for c in _ALL if c in rows] + [NO_VALUE_DATE, CANCELLED, "Currency"]].drop(
        columns=["DEBIT", "CLOSING_BALANCE"], errors="ignore")
    for c in ("VALUE_DATE", INVOICE_COLUMN):
        if c not in rows:
            rows[c] = ""
    days = pd.to_datetime(data.loc[~opening, "BOOKING_DATE"], errors="coerce", format="ISO8601")
    values = pd.to_datetime(rows["VALUE_DATE"], errors="coerce", format="ISO8601")
    first = min(d for d in (days.min(), values.min()) if pd.notna(d)).normalize() if days.notna().any() else None
    last = max(d for d in (days.max(), values.max()) if pd.notna(d)).normalize() if days.notna().any() else None
    info = {"currency": "", "date_from": None, "date_to": None, "first": first, "last": last,
            "balances": {"name": name, "first": first, "last": last, "opening": float(first_bal),
                         "closing": float(balance.dropna().iloc[-1]), "net": float(moves[~opening].sum()),
                         "usd_name": "USD" in name.upper(),
                         "refs": frozenset(r for r in data.loc[~opening, "BANK_REFERENCE"] if r),
                         "rows": [(r, float(m), float(b)) for r, m, b in zip(data.loc[~opening, "BANK_REFERENCE"],
                                                                             moves[~opening], balance[~opening])]}}
    return rows.reset_index(drop=True), 0, info          # every credit is kept: no "other credits"


def clean_statement(statement: pd.DataFrame) -> pd.DataFrame:
    """Typed columns + the invoice number (DOC ID; blank -> none)."""
    statement = statement.copy()
    inv = _clean_text(statement[INVOICE_COLUMN]) if INVOICE_COLUMN in statement else pd.Series("", index=statement.index)
    statement["invoice_number"] = inv.where(inv.ne(""), None)
    statement["CREDIT"] = _num(statement["CREDIT"])
    statement["BOOKING_DATE"] = pd.to_datetime(statement["BOOKING_DATE"], errors="coerce", format="ISO8601")
    statement["VALUE_DATE"] = pd.to_datetime(statement["VALUE_DATE"], errors="coerce", format="ISO8601")
    for c in (NO_VALUE_DATE, CANCELLED):
        if c not in statement:
            statement[c] = ""
    # the same credit in a file with VALUE_DATE and in a file without it (e.g. "sept 22 to oct 13.xlsx" and the day
    # files of 12 and 13 Oct): only the copy with VALUE_DATE is kept (decided 19/10/2026)
    ref = _clean_text(statement["BANK_REFERENCE"]) if "BANK_REFERENCE" in statement else pd.Series("", index=statement.index)
    no_date = statement[NO_VALUE_DATE].fillna("").eq("yes")
    dated = set(zip(ref[~no_date & ref.ne("")], statement.loc[~no_date & ref.ne(""), "CREDIT"]))
    copy = no_date & ref.ne("") & pd.Series([k in dated for k in zip(ref, statement["CREDIT"])], index=statement.index)
    return statement[~copy].reset_index(drop=True)


def find_invoices(statement: pd.DataFrame, payments: pd.DataFrame) -> pd.DataFrame:
    """A credit without DOC ID: its invoice is looked for in the IremboPay payment file, in this order (decided
    16/10/2026), and set ONLY when IremboPay has it (an invoice number is never counted when it is not found):
        1. BANK_REFERENCE = an IremboPay provider_reference   (e.g. TT26226M0NPS, a branch payment)
        2. DESCRIPTION    = an IremboPay provider_reference   (agent SZ… / internet-banking numbers)
        3. DESCRIPTION    = an IremboPay invoice number       (some USD exports have it there)
    A reference shared by several invoices is not used. The way it was found is kept ("Invoice found by")."""
    statement = statement.copy()
    statement[FOUND_BY] = ""
    statement.loc[statement["invoice_number"].notna(), FOUND_BY] = "DOC ID"
    by_ref = {}
    for inv, refs, pref in zip(payments["invoice_number"].astype(str),
                               payments.get("_provider_refs", pd.Series("", index=payments.index)).fillna(""),
                               payments["provider_reference"].fillna("").astype(str)):
        for r in (set(str(refs).split("|")) | {pref.strip()}) - {""}:
            by_ref.setdefault(r, set()).add(inv)
    by_ref = {r: next(iter(i)) for r, i in by_ref.items() if len(i) == 1}       # only a reference of ONE invoice
    invoices = set(payments["invoice_number"].astype(str))
    for i in statement.index[statement["invoice_number"].isna()]:
        ref = str(statement.at[i, "BANK_REFERENCE"] or "").strip()
        desc = str(statement.at[i, "DESCRIPTION"] or "").strip()
        for found, how in ((by_ref.get(ref), "BANK_REFERENCE = provider_reference"),
                           (by_ref.get(desc), "DESCRIPTION = provider_reference"),
                           (desc if desc in invoices else None, "DESCRIPTION = invoice number")):
            if found:
                statement.at[i, "invoice_number"] = found
                statement.at[i, FOUND_BY] = how
                break
    return statement


def channel(rows: pd.DataFrame) -> pd.Series:
    """The BPR channel, read from the DESCRIPTION (shown in the review)."""
    d = rows["DESCRIPTION"].fillna("").astype(str)
    return pd.Series(["AGENT" if x.startswith("SZ") else "INTERNET BANKING" if x.isdigit()
                      else "BRANCH" if x.startswith("FTCM") else "OTHER" for x in d], index=rows.index)


# =====================================================================================================================
# 5. CHECKS BETWEEN FILES (step 3): nothing may be missing
# =====================================================================================================================
def check_files(statements: list) -> tuple:
    """(days covered, lines, blocking): the files of each account (RWF, or USD by its name) one after the other.
    Two files follow when the second opens where the first closed; the days between them are then covered (a day
    without any payment has no file). A file whose own balance column is shifted while its rows are complete (the
    file before + its rows = the file after) is accepted with a note; anything else is a gap: shown, never blocking
    (blocking stays empty: Run is always possible)."""
    covered, lines, blocking = set(), [], []
    for f in statements:
        covered |= set(pd.date_range(f.first, f.last))
    accounts = {}
    for f in statements:
        b = getattr(f, "balances", None)
        if b:
            accounts.setdefault("USD" if b["usd_name"] else "RWF", []).append(b)
    for account, files in accounts.items():
        files = sorted(files, key=lambda b: (b["first"], b["last"]))
        # a file whose rows are all in another file (e.g. one day exported again inside a longer export) is left
        # out of the check: its rows are counted once (repeated rows are dropped when reconciling)
        inside = [b for b in files if b["refs"] and any(o is not b and b["refs"] < o["refs"] for o in files)]
        for b in inside:
            lines.append((True, f"{account}: {b['name']} is already inside "
                                f"{next(o['name'] for o in files if o is not b and b['refs'] < o['refs'])}"))
        files = [b for b in files if b not in inside]
        i = 1
        while i < len(files):
            a, b = files[i - 1], files[i]
            # a file may start by repeating the last row(s) of the file before (e.g. the daily debit exported
            # again): the check starts after those rows (their balance = where the file really opens)
            lead = 0
            for ref, _, _ in b.get("rows", []):
                if not ref or ref not in a["refs"]:
                    break
                lead += 1
            if lead and not b.get("lead_done"):
                b["opening"] = b["rows"][lead - 1][2]
                b["net"] -= sum(m for _, m, _ in b["rows"][:lead])
                b["lead_done"] = True
                lines.append((True, f"{account}: {b['name']} starts with {lead} row(s) already in {a['name']} "
                                    f"({', '.join(r for r, _, _ in b['rows'][:lead])}); counted once"))
            if abs(a["closing"] - b["opening"]) < 0.5:
                covered |= set(pd.date_range(a["last"], b["first"]))
                i += 1
                continue
            c = files[i + 1] if i + 1 < len(files) else None
            if c is not None and abs(a["closing"] + b["net"] - c["opening"]) < 0.5:
                covered |= set(pd.date_range(a["last"], c["first"]))
                lines.append((True, f"{account}: {b['name']} has its balance column shifted by "
                                    f"{b['opening'] - a['closing']:+,.0f}; its rows are complete (the file before + "
                                    "its rows = the file after)"))
                b["closing"] = a["closing"] + b["net"]
                i += 1
                continue
            gap = b["opening"] - a["closing"]
            if gap < 0:                                     # less money than expected: a debit is missing
                lines.append((False, f"{account}: between {a['name']} ({a['last']:%d/%m}) and {b['name']} "
                                     f"({b['first']:%d/%m}) {-gap:,.0f} went out with no row in the files (a debit "
                                     "is missing; no payment is missing)"))
            else:                                           # more money than expected: a credit (a payment) is missing
                lines.append((False, f"{account}: between {a['name']} ({a['last']:%d/%m}) and {b['name']} "
                                     f"({b['first']:%d/%m}) {gap:,.0f} came in with no row in the files: a payment "
                                     "is missing → upload the statement of those days"))
            covered |= set(pd.date_range(a["last"], b["first"]))
            i += 1
        if len(files) > 1 and not any(not ok for ok, t in lines if t.startswith(account + ":")):
            lines.append((True, f"{account}: every file follows the one before ({len(files)} files)"))
    return covered, lines, blocking


# =====================================================================================================================
# 6. SLA + CALCULATIONS + WHAT GOES TO REVIEW
# =====================================================================================================================
def sla(rows: pd.DataFrame, sla_minutes: int):
    """(minutes between VALUE_DATE and paid_at, in whole minutes; SLA status). On time when paid_at's minute is
    VALUE_DATE's minute or up to sla_minutes after (or before). Only BPR payments are checked; a row without
    VALUE_DATE cannot be checked. A review decision (sla_override) wins."""
    diff = (rows["paid_at"].dt.floor("min") - rows["VALUE_DATE"]).dt.total_seconds() / 60
    status = pd.Series("Cannot check", index=rows.index, dtype=object)
    has = diff.notna() & rows[NO_VALUE_DATE].fillna("").ne("yes")
    within = diff.abs() <= sla_minutes
    status[has & within] = "On time"
    status[has & ~within] = "SLA violated"
    override = rows["sla_override"].ne("") & rows["payment_provider"].isin(PROVIDERS)
    status[override] = rows.loc[override, "sla_override"]
    status[~rows["payment_provider"].isin(PROVIDERS)] = "Not processed"
    return diff, status


def received(rows: pd.DataFrame) -> pd.Series:
    """Notification received: in IremboPay, paid (paid_at filled, whatever the status) through this bank."""
    import engine                                 # (here: engine imports the banks)
    return engine.paid_through(rows, PROVIDERS)


def pending(rows: pd.DataFrame) -> pd.Series:
    """Pending notification: in IremboPay but not paid (no paid_at)."""
    import engine                                 # (here: engine imports the banks)
    return engine.not_paid(rows, PENDING_STATUSES)


def duplicate(rows: pd.DataFrame) -> pd.Series:
    """Duplicate Payment: in IremboPay, paid (paid_at), but through another provider."""
    import engine                                 # (here: engine imports the banks)
    return engine.paid_elsewhere(rows, PROVIDERS)


def _text(rows, col) -> list:
    return [("" if pd.isna(v) else str(v)).strip() for v in (rows[col] if col in rows else [""] * len(rows))]


def review_reasons(rows: pd.DataFrame) -> pd.Series:
    """Why a credit FOUND in IremboPay still goes to Exception Resolution ("" = it does not). rows = the credits of
    the period, matched with IremboPay."""
    found = rows["found_in_irembopay"].fillna(False).astype(bool).tolist()
    provider, desc, pref = _text(rows, "payment_provider"), _text(rows, "DESCRIPTION"), _text(rows, "provider_reference")
    all_refs = [set(r.split("|")) - {""} for r in _text(rows, "_provider_refs")]   # every IremboPay transaction
    inv, cancel, no_date = _text(rows, "invoice_number"), _text(rows, CANCELLED), _text(rows, NO_VALUE_DATE)
    refs, when = _text(rows, "BANK_REFERENCE"), list(rows["VALUE_DATE"])
    out = [[] for _ in range(len(rows))]
    for i in range(len(rows)):
        if found[i] and provider[i] not in PROVIDERS:
            out[i].append(f"Paid through {provider[i] or 'no provider'} in IremboPay (not BPR)")
    # an invoice credited more than once: the credit IremboPay recorded counts, the others are checked
    groups = {}
    for i in range(len(rows)):
        if inv[i]:                    # 26/10/2026: also an invoice NOT found in IremboPay
            groups.setdefault(inv[i], []).append(i)
    for idx in groups.values():
        if len(idx) < 2:
            continue
        kept = [i for i in idx if desc[i] and desc[i] == pref[i]]          # the transaction the report counts
        seen = [i for i in idx if desc[i] and (desc[i] in all_refs[i] or desc[i] == pref[i])]
        keep = kept[0] if kept else (seen[0] if seen else min(idx, key=lambda i: when[i]))
        for i in idx:
            if i == keep:
                continue
            if desc[i] and desc[i] in all_refs[i]:
                out[i].append(f"Invoice credited again: IremboPay also recorded this credit (another transaction); "
                              f"the credit counted is {refs[keep]}")
            else:
                out[i].append(f"Invoice credited again: IremboPay has no record of this credit; the credit counted "
                              f"is {refs[keep]}")
    for i in range(len(rows)):
        if found[i] and cancel[i]:
            out[i].append(f"Cancelled by the debit {cancel[i]}")
        if found[i] and no_date[i] == "yes":
            out[i].append("No VALUE_DATE in the statement (BOOKING_DATE used for the period)")
    return pd.Series(["; ".join(r) for r in out], index=rows.index, dtype=object)


def currency_not_found(rows: pd.DataFrame) -> pd.Series:
    """The currency of a credit not in the IremboPay payment file (a DOC ID IremboPay does not have, or no DOC ID):
    from the statement's file name ('USD' in the name = USD, otherwise RWF); the reviewer can change it ("Choose
    currency"). Decided 16/10/2026: these credits never stop the reconciliation."""
    return pd.Series(["USD" if "USD" in str(f).upper() else "RWF" for f in rows["Source file"]], index=rows.index)


def not_found_reason(rows: pd.DataFrame) -> pd.Series:
    """Why a credit NOT found in IremboPay is in Exception Resolution."""
    inv, cancel, no_date = _text(rows, "invoice_number"), _text(rows, CANCELLED), _text(rows, NO_VALUE_DATE)
    out = []
    for i in range(len(rows)):
        parts = ["No DOC ID" if not inv[i] or inv[i] == "None" else "DOC ID not in the IremboPay payment file"]
        if cancel[i]:
            parts.append(f"cancelled by the debit {cancel[i]}")
        if no_date[i] == "yes":
            parts.append("no VALUE_DATE")
        out.append("; ".join(parts))
    return pd.Series(out, index=rows.index, dtype=object)
