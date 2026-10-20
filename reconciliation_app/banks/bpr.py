"""
BPR (BANQUE POPULAIRE DU RWANDA)
================================
Everything that is specific to BPR (the common rules are in engine.py). Every bank credit counts (unlike I&M and
Bank of Africa, there is no "Irembo wording" filter): a credit with no invoice number in its description is kept
too, and is Pending with status "no invoice number" (engine.py, the automatic classification).

The BPR statement (one Excel file per day, any sub-folder)
    - columns BOOKING_DATE, BANK_REFERENCE, DESCRIPTION, a "DOC ID" column (spelled differently, or missing, in
      some files: the invoice can be in DESCRIPTION instead), VALUE_DATE (with time; missing in some files),
      DEBIT, CREDIT, CLOSING_BALANCE
    - currency: "USD" in the file name, else RWF (one file per currency)
    - a credit followed by a debit with the same BANK_REFERENCE is a reversal, not a payment
    - the same day can be saved in two files: de-duplicated by Bank Reference + invoice + amount

Dates: VALUE_DATE (with its time) sets the period and the SLA; a file with no VALUE_DATE uses BOOKING_DATE instead
("Has time" says which, so the SLA is checked by day for those rows).

payment_provider in the IremboPay file: "BPR" (agents, internet banking, bank transfer) and "BPR_RW" (branch) -
both count as BPR (engine.providers()).

CONTENTS
    1. Identity            4. Reading a statement
    2. Statement columns    5. SLA
    3. Files needed         6. Calculations (which row goes in which summary table)
"""
import io
import re

import pandas as pd

# =====================================================================================================================
# 1. IDENTITY
# =====================================================================================================================
CODE = "BPR"
NAME = "BPR"
FULL_NAME = "BPR Bank"
PROVIDER = "BPR"                         # payment_provider of BPR in the IremboPay payment file
PROVIDERS = ["BPR", "BPR_RW"]             # both count as BPR (agents/internet banking/transfer, and branch)
CURRENCIES = ["RWF", "USD"]
FILE_TYPES = ["xlsx", "xls"]
STATEMENT_SHEET = "BPR statement"

COMPLETED_STATUSES = {"SETTLED", "PAID", "PAYOUT_INITIATED"}
PENDING_STATUSES = {"NEW", "EXPIRED"}


# =====================================================================================================================
# 2. STATEMENT COLUMNS (as they are named on the BPR statement)
# =====================================================================================================================
COLUMNS = {
    "posted_date": "Booking Date",
    "payment_date": "Value Date",        # VALUE_DATE, or BOOKING_DATE when the file has none ("Has time" says which)
    "description": "Description",
    "reference": "Bank Reference",
    "amount": "Credit",
}
DETAIL_COLUMNS = ["Has time"]
DATETIME_COLUMNS = ["Value Date"]
DUPLICATE_KEY = ["Currency", "Bank Reference", "invoice_number", "Credit"]

_REQUIRED_HEADERS = {"BOOKING_DATE", "BANK_REFERENCE", "DESCRIPTION", "DEBIT", "CREDIT"}
_INVOICE = re.compile(r"(?<!\d)(8[6-9]\d{10})(?!\d)")


# =====================================================================================================================
# 3. FILES NEEDED (used by the checks of step 3)
# =====================================================================================================================
MONDAY_AFTER_WEEKEND = False
DAY_BEFORE_CURRENCIES = []

# SLA: a fixed rule (at most 1 minute after VALUE_DATE), so the "SLA days" setting is hidden
SLA_DAYS_SETTING = False
SLA_DIFF_COLUMN = "sla_seconds_difference"


# =====================================================================================================================
# 4. READING A STATEMENT
# =====================================================================================================================
def read_statement(content: bytes, name: str, excel_engine=None):
    """Return (rows, other_credits, info); the currency comes from the file name. Raise ValueError with a plain
    message when the file is not a BPR statement."""
    try:
        raw = pd.read_excel(io.BytesIO(content) if isinstance(content, (bytes, bytearray)) else content,
                            sheet_name=0, dtype=str, engine=excel_engine)
    except Exception:
        raise ValueError(f"{name}: this file cannot be opened as a BPR statement.") from None
    raw.columns = [str(c).strip().upper().replace(" ", "") for c in raw.columns]
    missing = _REQUIRED_HEADERS - set(raw.columns)
    if missing:
        raise ValueError(f"{name}: this is not a BPR statement (missing {', '.join(sorted(missing))}).")

    # ---- the invoice number: in the "DOC ID" column (spelled differently in some files) or in the Description
    doc = [c for c in raw.columns if c.startswith("DOC") or c.startswith("UNNAMED")]
    text = raw["DESCRIPTION"].fillna("") + " " + (raw[doc[0]].fillna("") if doc else "")
    invoice = text.str.replace(r"\.0\b", "", regex=True).str.extract(_INVOICE)[0]

    credit = pd.to_numeric(raw["CREDIT"], errors="coerce").fillna(0)
    debit = pd.to_numeric(raw["DEBIT"], errors="coerce").fillna(0)
    value = pd.to_datetime(raw["VALUE_DATE"], errors="coerce") if "VALUE_DATE" in raw.columns \
        else pd.Series(pd.NaT, index=raw.index)
    booking = pd.to_datetime(raw["BOOKING_DATE"], errors="coerce")
    has_time = value.notna()
    currency = "USD" if "USD" in name.upper() else "RWF"

    # ---- a credit cancelled by a debit with the same Bank Reference (same file) is not a payment
    reversed_refs = set(raw.loc[debit.ne(0), "BANK_REFERENCE"].dropna())
    keep = (credit > 0) & ~raw["BANK_REFERENCE"].isin(reversed_refs)

    data = pd.DataFrame({
        "Booking Date": booking, "Value Date": value.where(has_time, booking), "Description": raw["DESCRIPTION"].fillna(""),
        "Bank Reference": raw["BANK_REFERENCE"], "Credit": credit, "Currency": currency,
        "Has time": has_time, "invoice_number": invoice,
    })
    data = data[keep].reset_index(drop=True)
    info = {"currency": currency, "date_from": None, "date_to": None,
            "first": data["Value Date"].min(), "last": data["Value Date"].max()}
    return data, 0, info


def clean_statement(statement: pd.DataFrame) -> pd.DataFrame:
    """Typed columns; the invoice number is already read in read_statement()."""
    s = statement.copy()
    s["invoice_number"] = s["invoice_number"].astype("string")
    s["Bank Reference"] = s["Bank Reference"].astype(str).str.strip()
    return s


def channel(rows: pd.DataFrame) -> pd.Series:
    """Every BPR row is shown the same way in the review (agents, internet banking, branch, bank transfer are all
    BPR; the statement does not say which)."""
    return pd.Series(NAME, index=rows.index)


# =====================================================================================================================
# 5. SLA: paid_at against VALUE_DATE, at most 1 minute after; a row with no VALUE_DATE is checked by day instead
# =====================================================================================================================
def sla(rows: pd.DataFrame, sla_days: int):
    """(seconds between VALUE_DATE and paid_at, SLA status): on time when the two, each rounded down to the
    minute, are at most 1 minute apart (a statement with no VALUE_DATE is checked by day instead). Only BPR
    payments are checked; the others are "Not processed". A review decision (sla_override) wins."""
    diff = (rows["paid_at"] - rows["Value Date"]).dt.total_seconds()
    within_a_minute = (rows["paid_at"].dt.floor("min") - rows["Value Date"].dt.floor("min")).abs() <= pd.Timedelta(minutes=1)
    same_day = rows["paid_at"].dt.date == rows["Value Date"].dt.date
    on_time = within_a_minute.where(rows["Has time"], same_day)
    status = pd.Series("Cannot check", index=rows.index, dtype=object)
    has_date = rows["paid_at"].notna() & rows["Value Date"].notna()
    status[has_date & on_time] = "On time"
    status[has_date & ~on_time] = "SLA violated"
    is_bpr = rows["payment_provider"].astype(str).str.upper().isin(PROVIDERS)
    override = rows["sla_override"].ne("") & is_bpr
    status[override] = rows.loc[override, "sla_override"]
    status[~is_bpr] = "Not processed"
    return diff, status


# =====================================================================================================================
# 6. CALCULATIONS (one True/False per row). Every BPR credit of the period is a processed payment.
# =====================================================================================================================
def received(rows: pd.DataFrame) -> pd.Series:
    """Notification received by Irembo: found, paid through BPR, and completed in IremboPay."""
    return (rows["found_in_irembopay"] & rows["payment_provider"].astype(str).str.upper().isin(PROVIDERS)
            & rows["payment_status"].isin(COMPLETED_STATUSES))


def pending(rows: pd.DataFrame) -> pd.Series:
    """Pending notification: the invoice is not paid yet in IremboPay (NEW, or EXPIRED)."""
    return rows["payment_status"].isin(PENDING_STATUSES)


def duplicate(rows: pd.DataFrame) -> pd.Series:
    """Duplicate payment: credited on BPR, but IremboPay says another provider was used."""
    return (rows["found_in_irembopay"] & rows["payment_provider"].notna()
            & ~rows["payment_provider"].astype(str).str.upper().isin(PROVIDERS)
            & ~rows["payment_status"].isin(PENDING_STATUSES))
