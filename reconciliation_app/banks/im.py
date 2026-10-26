"""
I&M BANK
========
Everything that is specific to I&M (the common rules are in engine.py). This file is also the TEMPLATE for a new
bank: copy it, change the values and functions, and add the new file to banks/__init__.py.

The I&M statement
    - one Excel file per currency (RWF or USD); the currency is read from the "Balance (RWF)" / "Balance (USD)" column
    - the days a file covers are in its header: "Date From" and "Date To"
    - only the credits whose Description contains "IREMBO PAYMENT" are used; the invoice number follows those words

Dates on an I&M statement
    - Transaction Date = the day the bank posted the payment
    - Value Date       = the day the payment took place; when the SLA is met it is the paid_at day in IremboPay
    So the period and the SLA use the Value Date. A payment is posted on its Value Date or later, never before: a
    payment made on a weekend or a holiday is posted on the next working day, so a period ending on a weekend or a
    holiday needs the statement of the next working day.

CONTENTS
    1. Identity            4. Reading a statement
    2. Statement columns   5. SLA
    3. Files needed        6. Calculations (which row goes in which summary table)
"""
import io
import re

import pandas as pd


# =====================================================================================================================
# 1. IDENTITY
# =====================================================================================================================
CODE = "IM"                       # short code: report file names, review decisions, run log
NAME = "I&M"                      # short name, e.g. "I&M bank statements"
FULL_NAME = "I&M Bank"            # e.g. "I&M Bank Reconciliation"
PROVIDER = "I_M"                  # payment_provider of this bank in the IremboPay payment file
CURRENCIES = ["RWF", "USD"]       # one statement per currency
FILE_TYPES = ["xls", "xlsx"]      # statement files accepted in the upload box
STATEMENT_SHEET = "I&M statement" # name of the statement sheet in the report

# IremboPay statuses that count as paid
COMPLETED_STATUSES = {"SETTLED", "PAID", "PAYOUT_INITIATED"}
# pending = not paid yet: payment_status NEW (20/10/2026: "expired" is not a payment status - an expired
# invoice was never paid, it stays NEW; the expiry is a note in the report)
PENDING_STATUSES = {"NEW"}


# =====================================================================================================================
# 2. STATEMENT COLUMNS (as they are named on the I&M statement)
# =====================================================================================================================
COLUMNS = {
    "posted_date": "Transaction Date",   # day the bank posted the payment
    "payment_date": "Value Date",        # day the payment took place: used for the period and the SLA
    "description": "Description",
    "reference": "Tran Id",              # bank reference of the row
    "amount": "Credit",                  # credits only
}

_IREMBO = "IREMBO PAYMENT"               # words that mark an IremboPay payment in the Description


# =====================================================================================================================
# 3. FILES NEEDED (used by the checks of step 3)
# =====================================================================================================================
MONDAY_AFTER_WEEKEND = True       # a period that ends on a weekend or a holiday also needs the statement of the next
                                  # working day (Monday, or later after a holiday: Settings → public holidays)
MISSING_DAYS_BLOCK = False        # 26/10/2026: a day with no statement only warns in step 3 (Run allowed): a day
                                  # may simply have had no transaction
MISSING_DAYS_NOTE = False         # 26/10/2026: the "No statement for ..." note is NOT added to step 6 nor to the
                                  # Slack message (the other banks keep it)
DAY_BEFORE_CURRENCIES = []        # 20/10/2026: no day-before recommendation - the Transaction Date is never before the
                                  # Value Date, so the day before holds no payment of the period
SEPARATE_CREDITED_AGAIN = True    # step 5: the extra credits of an invoice in a table of their own (20/10/2026)


# =====================================================================================================================
# 4. READING A STATEMENT
# =====================================================================================================================
def read_statement(content: bytes, name: str, excel_engine=None):
    """Return (rows, other_credits, info). Rows keep the statement's own column names (as text) plus Currency.
    info: currency, date_from / date_to (from the header, or None), first / last (Value Dates found).
    Raise ValueError with a plain message when the file is not an I&M statement."""
    try:
        raw = pd.read_excel(io.BytesIO(content) if isinstance(content, (bytes, bytearray)) else content,
                            header=None, dtype=str, engine=excel_engine).fillna("")
    except Exception:
        raise ValueError(f"{name}: this file cannot be opened as an Excel bank statement.") from None
    raw = raw.apply(lambda col: col.str.strip())

    # ---- the header row ("Transaction Date"), and where each column is
    header_rows = raw.index[raw.eq("Transaction Date").any(axis=1)]
    if len(header_rows) == 0:
        raise ValueError(f"{name}: this is not an I&M bank statement (no 'Transaction Date' header).")
    header_idx = header_rows[0]
    positions = {col_name: col for col, col_name in raw.loc[header_idx].items() if col_name}

    # ---- the currency, from the "Balance (RWF)" or "Balance (USD)" column name
    balance = [n for n in positions if n.startswith("Balance")]
    match = re.search(r"\((\w{3})\)", balance[0]) if balance else None
    if not match:
        raise ValueError(f"{name}: no 'Balance (RWF)' or 'Balance (USD)' column, so the currency is unknown.")
    currency = match.group(1).upper()
    wanted = list(COLUMNS.values())
    missing = [c for c in wanted if c not in positions]
    if missing:
        raise ValueError(f"{name}: the bank file is missing the column(s) {', '.join(missing)}.")

    def header_date(label):
        """The date written next to a label above the header (e.g. 'Date From(dd/MM/yyyy):')."""
        hit = raw.loc[:header_idx].isin([label])
        if not hit.any().any():
            return None
        r, c = next((r, c) for c in hit.columns for r in hit.index[hit[c]])
        after = [v for v in raw.loc[r, c + 1:] if v]
        return pd.to_datetime(after[0], errors="coerce") if after else None

    # ---- the rows: IremboPay credits only (the other credits are only counted)
    data = raw.loc[header_idx + 1:, [positions[c] for c in wanted]]
    data.columns = wanted
    credit = pd.to_numeric(data["Credit"].str.replace(",", "", regex=False), errors="coerce").fillna(0)
    is_irembo = data["Description"].str.contains(_IREMBO, case=False)
    other_credits = int(((credit > 0) & ~is_irembo).sum())
    data = data[(credit > 0) & is_irembo].copy()
    data["Currency"] = currency
    value_dates = pd.to_datetime(data["Value Date"], format="%d/%m/%Y", errors="coerce")
    info = {"currency": currency,
            "date_from": header_date("Date From(dd/MM/yyyy):"), "date_to": header_date("Date To(dd/MM/yyyy):"),
            "first": value_dates.min(), "last": value_dates.max()}
    return data.reset_index(drop=True), other_credits, info


def clean_statement(statement: pd.DataFrame) -> pd.DataFrame:
    """Typed columns + the invoice number (matched with IremboPay)."""
    statement = statement.copy()
    statement["invoice_number"] = statement["Description"].str.extract(rf"{_IREMBO}\W*(\d{{10,}})", flags=re.I)[0]
    statement["Credit"] = pd.to_numeric(statement["Credit"].str.replace(",", "", regex=False), errors="coerce")
    statement["Transaction Date"] = pd.to_datetime(statement["Transaction Date"], format="%d/%m/%Y", errors="coerce")
    statement["Value Date"] = pd.to_datetime(statement["Value Date"], format="%d/%m/%Y", errors="coerce")
    return statement


def channel(rows: pd.DataFrame) -> pd.Series:
    """Bank channel of a row (the word before IREMBO PAYMENT), shown in the review."""
    return rows["Description"].str.extract(rf"^\s*(\w+)\s+{_IREMBO}", flags=re.I)[0].fillna("BRANCH/OTHER")


# =====================================================================================================================
# 5. SLA: the paid_at day against the Value Date (0 days allowed by default; "Settings" in the side panel)
# =====================================================================================================================
def sla(rows: pd.DataFrame, sla_days: int):
    """(days difference, SLA status) per matched row: the paid_at day against the Value Date.
    Only I&M payments are checked; the others are "Not processed". A review decision (sla_override) wins."""
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
# 6. CALCULATIONS (one True/False per bank row). Processed = the invoices on both sides, each once (engine,
#    20/10/2026); these decide which of the other summary tables an invoice goes in.
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
    (20/10/2026): an invoice credited more than once - the earliest credit (Value Date, then Tran Id) counts, every
    other credit of that invoice goes to the step-5 table "Invoice credited more than once" to be viewed."""
    out = pd.Series("", index=rows.index, dtype=object)
    # 26/10/2026: also an invoice NOT found in IremboPay - its second credit goes to "credited more than once" too
    found = rows[rows["invoice_number"].notna()]
    for _, group in found.groupby("invoice_number"):
        if len(group) < 2:
            continue
        group = group.sort_values(["Value Date", "Tran Id"], kind="stable")
        kept = str(group["Tran Id"].iloc[0]).strip()
        out[group.index[1:]] = f"Invoice credited again: the credit counted is {kept}"
    return out
