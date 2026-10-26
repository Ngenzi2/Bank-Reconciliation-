"""
BANK OF AFRICA RWANDA (BOA)
===========================
Everything that is specific to Bank of Africa (the common rules are in engine.py). It works like I&M: only the
Irembo payment CREDITS of the statement are reconciled, the period is set by the bank Value Date, and the SLA compares
the IremboPay paid_at day with the Value Date. It is a file of its own: I&M (im.py) is not used or changed.

The BOA statement ("Account history" export of the e-banking, .xlsx)
    - a block of account details first (Account currency, Opening/Closing Bal, Total Credit/Debit, Transaction
      Count), then the table "# | Ledger Date | Value Date | Branch | Ref no | Description | Debit | Credit |
      Balance", newest first, then a "Totals (D / C)" row and a footer line. The header row is found, not assumed.
    - every file is checked against its own account details before use: the number of rows = Transaction Count,
      the credits = Total Credit, the debits = Total Debit, Opening + credits - debits = Closing, and each Balance
      follows from the row before. A file that does not add up is refused with the figures (nothing is "fixed").
      This is the only use of the debits.
    - the export is "Last N transactions": no "Date From / Date To". The days it covers come from the dates of its
      rows; the OLDEST day is left out of the coverage, because the "Last N" limit can cut that day in the middle.
    - one file = one currency (Account currency); only RWF accounts are supported.

The credits kept (evidence: the sample statement + IremboPay export of 08/10/2026, 429 rows / 366 payments)
    "ePay irembo <invoice>"          internet banking  (IremboPay payment_channel BOA_IB)
    "irembo Payment for <invoice>"   USSD              (BOA_USSD)
    "YOUR CASH DEPOSIT <invoice>"    branch deposit    (BOA_BRANCH)
    "irembo Invoice Permit"          no invoice in the description: found by its Ref no (see below)
    Not kept (only counted, as I&M does with its other credits): credits without Irembo wording or invoice (e.g.
    "REFUND/IBank monthly"), and a credit cancelled in the same file by a "Cancel" debit with the same Ref no and
    amount (that money did not stay in the account).

Matching with IremboPay (evidence from the samples)
    - the 12-digit number in the description is the IremboPay invoice_number (362 of 362 payments that have one)
    - the bank Ref no is the IremboPay provider_reference (366 of 366, permits included), so a payment with no
      invoice in its description is matched by Ref no = provider_reference, only when exactly one payment has it
    - "Ref check" (detail sheets) says, for every row, whether IremboPay recorded this very credit (same Ref no); an
      invoice credited twice is flagged like for I&M (times_on_statement / is_extra_credit)

Dates: the bank Value Date (= Ledger Date on BOA) sets the period; the SLA compares it with IremboPay paid_at.

CONTENTS
    1. Identity            4. Reading a statement (checked against its own totals)
    2. Statement columns   5. Cleaning + matching by Ref no
    3. Files needed        6. SLA            7. Calculations
"""
import io
import re

import pandas as pd


# =====================================================================================================================
# 1. IDENTITY
# =====================================================================================================================
CODE = "BOA"
NAME = "Bank of Africa"
FULL_NAME = "Bank of Africa Rwanda"
PROVIDER = "BANK_OF_AFRICA"              # payment_provider in the IremboPay payment file
CURRENCIES = ["RWF"]                     # only RWF accounts are known; another currency is refused (section 4)
FILE_TYPES = ["xlsx", "xls"]
STATEMENT_SHEET = "BOA statement"

COMPLETED_STATUSES = {"SETTLED", "PAID", "PAYOUT_INITIATED"}     # same as I&M
# pending = not paid yet: payment_status NEW (20/10/2026: "expired" is not a payment status - an expired
# invoice was never paid, it stays NEW; the expiry is a note in the report)
PENDING_STATUSES = {"NEW"}


# =====================================================================================================================
# 2. STATEMENT COLUMNS (as they are named on the BOA statement)
# =====================================================================================================================
COLUMNS = {
    "posted_date": "Ledger Date",
    "payment_date": "Value Date",        # sets the period and is compared with paid_at (SLA)
    "description": "Description",
    "reference": "Ref no",               # = IremboPay provider_reference
    "amount": "Credit",                  # credits only
}
_TABLE = ["#", "Ledger Date", "Value Date", "Branch", "Ref no", "Description", "Debit", "Credit", "Balance"]
_META = {"Account currency": "currency", "Opening Bal": "opening", "Closing Bal": "closing",
         "Total Credit": "total_credit", "Total Debit": "total_debit", "Transaction Count": "count"}
DETAIL_COLUMNS = ["Branch", "Ref check"]                   # shown in the report's detail sheets


# =====================================================================================================================
# 3. FILES NEEDED (used by the checks of step 3)
# =====================================================================================================================
MONDAY_AFTER_WEEKEND = False      # BOA gives weekend credits a Monday Value Date: they belong to Monday
DAY_BEFORE_CURRENCIES = []        # no payment of the period sits in the statement of the day before
SEPARATE_CREDITED_AGAIN = True     # step 5: the extra credits of an invoice in a table of their own (20/10/2026)
MISSING_DAYS_BLOCK = False         # a day with no statement may simply have had no transaction (the export has no
                                   # "from / to" dates): a warning, Run stays possible, the Slack message says it


# =====================================================================================================================
# 4. READING A STATEMENT
# =====================================================================================================================
_INVOICE = re.compile(r"(?<!\d)(\d{12})(?!\d)")


def _number(value):
    """'27,516,669' / '-10' / 1410.0 / '' -> float or NaN."""
    text = str(value).replace(",", "").replace(" ", "").strip()
    return pd.to_numeric(text, errors="coerce") if text else float("nan")


def _dates(values: pd.Series) -> pd.Series:
    """'05/10/2026' (day first) or an Excel date ('2026-10-05 00:00:00') -> Timestamp; anything else -> NaT."""
    text = values.astype(str).str.strip()
    out = pd.to_datetime(text, format="%d/%m/%Y", errors="coerce")
    iso = out.isna() & text.str.match(r"^\d{4}-\d{2}-\d{2}")
    out[iso] = pd.to_datetime(text[iso].str[:10], format="%Y-%m-%d", errors="coerce")
    return out.astype("datetime64[ns]")                  # one type, whatever the cells held


def _irembo_credits(desc: pd.Series, ref: pd.Series, debit: pd.Series, credit: pd.Series) -> pd.Series:
    """True for the Irembo payment credits: a credit with Irembo wording, or a cash deposit / an invoice number, that
    is not a refund and was not cancelled in the same file (a "Cancel" debit with the same Ref no and amount)."""
    is_credit = credit.fillna(0) > 0
    looks_irembo = (desc.str.contains(r"(?i)irembo|cash deposit") | desc.str.contains(r"(?<!\d)\d{12}(?!\d)")) \
        & ~desc.str.contains(r"(?i)^\s*refund")
    cancel = (debit.fillna(0) < 0) & desc.str.match(r"(?i)\s*cancel\b")
    cancelled = set(zip(ref[cancel].str.strip(), (-debit[cancel]).round(2)))
    reversed_ = pd.Series([(r.strip(), round(c, 2)) in cancelled if c == c else False
                           for r, c in zip(ref, credit)], index=ref.index)
    return is_credit & looks_irembo & ~reversed_


def read_statement(content: bytes, name: str, excel_engine=None):
    """Return (rows, other_credits, info): the Irembo payment credits (text columns plus Currency), the number of
    other credits, and info (currency, no date_from/date_to, first/last = the days covered). Raise ValueError with a
    plain message when the file is not a BOA account history, uses another currency, or does not add up."""
    try:
        raw = pd.read_excel(io.BytesIO(content) if isinstance(content, (bytes, bytearray)) else content,
                            header=None, dtype=str, engine=excel_engine).fillna("")
    except Exception:
        raise ValueError(f"{name}: this file cannot be opened as an Excel bank statement.") from None
    raw = raw.apply(lambda col: col.str.strip())

    # ---- the header row: the row that has the table's column names (wherever it is)
    is_header = raw.apply(lambda r: {"Ledger Date", "Value Date", "Ref no", "Credit"} <= set(r), axis=1)
    if not is_header.any():
        raise ValueError(f"{name}: this is not a Bank of Africa account history (no 'Ledger Date / Value Date / "
                         "Ref no / Credit' header).")
    header_idx = is_header.idxmax()
    positions = {v: c for c, v in raw.loc[header_idx].items() if v}
    missing = [c for c in _TABLE if c not in positions]
    if missing:
        raise ValueError(f"{name}: the bank file is missing the column(s) {', '.join(missing)}.")

    # ---- the account details above the table (label in one cell, value in the next filled one)
    meta = {}
    for i in raw.index[raw.index < header_idx]:
        if raw.iat[i, 0] in _META:
            meta[_META[raw.iat[i, 0]]] = next((v for v in raw.loc[i].iloc[1:] if v), "")
    currency = str(meta.get("currency", "")).upper()
    if not currency:
        raise ValueError(f"{name}: no 'Account currency' in the account details, so the currency is unknown.")
    if currency not in CURRENCIES:
        raise ValueError(f"{name}: this is a {currency} account; only {', '.join(CURRENCIES)} Bank of Africa "
                         "accounts are supported.")

    # ---- the table rows: a row number in '#', down to the 'Totals' row; repeated headers and blank lines skipped
    body = raw.loc[header_idx + 1:, [positions[c] for c in _TABLE]]
    body.columns = _TABLE
    totals_rows = body.index[body["#"].str.lower().str.startswith("totals")]
    if len(totals_rows):
        body = body.loc[:totals_rows[0] - 1]
    body = body[~body.eq("").all(axis=1) & body["Ledger Date"].ne("Ledger Date")]
    debit, credit = body["Debit"].map(_number), body["Credit"].map(_number)
    ledger, value = _dates(body["Ledger Date"]), _dates(body["Value Date"])
    bad = (~body["#"].str.fullmatch(r"\d+") | ledger.isna() | value.isna() | (debit.isna() & credit.isna())
           | (debit.fillna(0).ne(0) & credit.fillna(0).ne(0)))
    if bad.any():
        shown = ", ".join(str(i + 1) for i in body.index[bad][:5])
        raise ValueError(f"{name}: {int(bad.sum())} row(s) of the table cannot be read as transactions "
                         f"(Excel rows {shown}…): a row needs a number, two dates and either a debit or a credit.")

    # ---- the file must add up to its own account details (the only use of the debits; nothing is corrected)
    problems = []
    n_credit, n_debit = credit.fillna(0).sum(), (-debit.fillna(0)).sum()
    count, opening, closing = (_number(meta.get(k, "")) for k in ("count", "opening", "closing"))
    total_credit, total_debit = _number(meta.get("total_credit", "")), _number(meta.get("total_debit", ""))
    if pd.notna(count) and int(count) != len(body):
        problems.append(f"{len(body)} rows read but the statement says {int(count)} transactions")
    if pd.notna(total_credit) and abs(n_credit - total_credit) > 0.5:
        problems.append(f"credits add up to {n_credit:,.0f} but Total Credit is {total_credit:,.0f}")
    if pd.notna(total_debit) and abs(n_debit - abs(total_debit)) > 0.5:
        problems.append(f"debits add up to {n_debit:,.0f} but Total Debit is {abs(total_debit):,.0f}")
    if pd.notna(opening) and pd.notna(closing) and abs(opening + n_credit - n_debit - closing) > 0.5:
        problems.append(f"Opening {opening:,.0f} + credits - debits = {opening + n_credit - n_debit:,.0f}, not the "
                        f"Closing {closing:,.0f}")
    order = body["#"].astype(int).to_numpy().argsort()            # newest first: row 1 is the latest
    b, c, d = (s.to_numpy()[order] for s in (body["Balance"].map(_number), credit.fillna(0), -debit.fillna(0)))
    breaks = [i for i in range(len(b) - 1)
              if pd.notna(b[i]) and pd.notna(b[i + 1]) and abs(b[i] - (b[i + 1] + c[i] - d[i])) > 0.5]
    if breaks:
        problems.append(f"the Balance column does not follow from the rows at {len(breaks)} place(s)")
    if problems:
        raise ValueError(f"{name}: the statement does not add up to its own account details: "
                         + "; ".join(problems) + ". The file may be incomplete or edited; download it again.")

    # ---- the rows kept: the Irembo payment credits only (like I&M); the other credits are only counted
    keep = _irembo_credits(body["Description"], body["Ref no"], debit, credit)
    other_credits = int((credit.fillna(0) > 0).sum() - keep.sum())
    data = body[keep].drop(columns=["Debit"]).copy()
    data["Currency"] = currency
    days = value.dt.normalize()                                   # coverage: every row of the file
    info = {"currency": currency, "date_from": None, "date_to": None,
            "first": days.min() + pd.Timedelta(days=1) if days.nunique() > 1 else days.min(), "last": days.max()}
    return data.reset_index(drop=True), other_credits, info


# =====================================================================================================================
# 5. CLEANING + MATCHING BY REF NO
# =====================================================================================================================
def clean_statement(statement: pd.DataFrame) -> pd.DataFrame:
    """Typed columns + the invoice number (the 12-digit number of the description), like I&M."""
    s = statement.copy()
    s["Credit"], s["Balance"] = s["Credit"].map(_number), s["Balance"].map(_number)
    s["Ledger Date"], s["Value Date"] = _dates(s["Ledger Date"]), _dates(s["Value Date"])
    s["Ref no"] = s["Ref no"].astype(str).str.strip()
    s["invoice_number"] = s["Description"].astype(str).str.extract(_INVOICE)[0]
    s["Ref check"] = ""
    return s


def channel(rows: pd.DataFrame) -> pd.Series:
    """BOA channel of a row, from its description (shown in the review)."""
    desc = rows["Description"].astype(str).str.lower()
    out = pd.Series("Other", index=rows.index, dtype=object)
    out[desc.str.contains("epay irembo")] = "Internet banking"
    out[desc.str.contains("irembo payment for")] = "USSD"
    out[desc.str.contains("cash deposit")] = "Branch"
    out[desc.str.contains("invoice permit")] = "Permit"
    return out


def one_row_per_invoice(rows: pd.DataFrame, payments: pd.DataFrame) -> pd.DataFrame:
    """(engine hook, called with the IremboPay payments) Every credit is kept. A credit without an invoice in its
    description gets the invoice of the IremboPay payment whose provider_reference is its Ref no, only when exactly
    one payment has it. "Ref check" says whether IremboPay recorded this very credit."""
    rows = rows.copy()
    if "provider_reference" not in payments:
        return rows
    refs = payments[["invoice_number", "provider_reference"]].dropna(subset=["provider_reference"]).copy()
    refs["provider_reference"] = refs["provider_reference"].astype(str).str.strip()
    counts = refs["provider_reference"].value_counts()
    unique = refs[refs["provider_reference"].map(counts).eq(1)].set_index("provider_reference")["invoice_number"]
    by_invoice = refs.drop_duplicates("invoice_number").set_index("invoice_number")["provider_reference"]
    fill = rows["invoice_number"].isna() & rows["Ref no"].isin(unique.index)
    rows.loc[fill, "invoice_number"] = rows.loc[fill, "Ref no"].map(unique)
    rows.loc[fill, "Ref check"] = "matched by Ref no (no invoice in the description)"
    recorded = rows["invoice_number"].map(by_invoice)
    same = ~fill & recorded.notna() & recorded.eq(rows["Ref no"])
    other = ~fill & recorded.notna() & recorded.ne(rows["Ref no"])
    rows.loc[same, "Ref check"] = "Ref no = provider_reference"
    rows.loc[other, "Ref check"] = "IremboPay recorded another credit: Ref " + recorded[other].astype(str)
    none = rows["invoice_number"].isna()
    shared = none & rows["Ref no"].map(counts).gt(1)
    rows.loc[none, "Ref check"] = "no invoice in the description, Ref no not in IremboPay"
    rows.loc[shared, "Ref check"] = ("no invoice in the description, Ref no shared by "
                                     + rows.loc[shared, "Ref no"].map(counts).astype(int).astype(str)
                                     + " IremboPay payments: not matched (ambiguous)")
    return rows


# =====================================================================================================================
# 6. SLA: the paid_at day against the Value Date (as for I&M; 0 days allowed by default, "Settings" in the side panel)
# =====================================================================================================================
def sla(rows: pd.DataFrame, sla_days: int):
    """(days difference, SLA status) per matched row: the IremboPay paid_at day against the bank Value Date.
    Only BOA payments are checked; the others are "Not processed". A review decision (sla_override) wins."""
    diff = (rows["paid_at"].dt.normalize() - rows["Value Date"]).dt.days
    status = pd.Series("Cannot check", index=rows.index, dtype=object)
    has_date = diff.notna()
    within = diff.abs() <= sla_days
    status[has_date & within] = "On time"
    status[has_date & ~within] = "SLA violated"
    override = rows["sla_override"].ne("") & rows["payment_provider"].eq(PROVIDER)
    status[override] = rows.loc[override, "sla_override"]
    status[~rows["payment_provider"].eq(PROVIDER)] = "Not processed"
    return diff, status


# =====================================================================================================================
# 7. CALCULATIONS (one True/False per credit; processed = the invoices on both sides, each once - engine)
# =====================================================================================================================
def received(rows: pd.DataFrame) -> pd.Series:
    """Notification received: in IremboPay, paid (paid_at filled, whatever the status) through this bank."""
    import engine                                 # (here: engine imports the banks)
    return engine.paid_through(rows, PROVIDER)


def pending(rows: pd.DataFrame) -> pd.Series:
    """Pending notification: in IremboPay but not paid (no paid_at)."""
    import engine                                 # (here: engine imports the banks)
    return engine.not_paid(rows, PENDING_STATUSES)


def duplicate(rows: pd.DataFrame) -> pd.Series:
    """Duplicate Payment: in IremboPay, paid (paid_at), but through another provider."""
    import engine                                 # (here: engine imports the banks)
    return engine.paid_elsewhere(rows, PROVIDER)


def review_reasons(rows: pd.DataFrame) -> pd.Series:
    """Why a credit FOUND in IremboPay still goes to Exception Resolution ("" = it does not). Only one case
    (20/10/2026): an invoice credited more than once - the credit IremboPay recorded counts (its Ref no is the
    provider_reference, see "Ref check"), otherwise the earliest; every other credit of that invoice goes to the
    step-5 table "Invoice credited more than once" to be viewed."""
    out = pd.Series("", index=rows.index, dtype=object)
    # 26/10/2026: also an invoice NOT found in IremboPay - its second credit goes to "credited more than once" too
    found = rows[rows["invoice_number"].notna()]
    for _, group in found.groupby("invoice_number"):
        if len(group) < 2:
            continue
        check = group["Ref check"].fillna("").astype(str) if "Ref check" in group else pd.Series("", index=group.index)
        recorded = check.eq("Ref no = provider_reference") | check.str.startswith("matched by Ref no")
        group = group.assign(_r=(~recorded).astype(int)).sort_values(["_r", "Value Date", "Ref no"], kind="stable")
        kept = str(group["Ref no"].iloc[0]).strip()
        out[group.index[1:]] = f"Invoice credited again: the credit counted is {kept}"
    return out
