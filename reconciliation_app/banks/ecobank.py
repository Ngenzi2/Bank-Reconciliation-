"""
ECOBANK
=======
Everything that is specific to Ecobank (the common rules are in engine.py).

The Ecobank statement ("IREMBO PAY" export) is a log of payment REQUESTS, not of credits
    - Request Date = when the customer paid at Ecobank, in UTC; Kigali time = UTC + 2 h
    - one invoice can be tried many times (failed attempts, retries): one row per invoice is kept (section 5)
    - one file holds every currency (Ccy column); a file's day is read from its Request Dates, not its name
    - TACCT (transit-account) files are a different flow: they are refused

The report (approved 2026-10-25; before that, processed/notified/SLA/duplicate were counted by IremboPay's
paid_at instead - CLAUDE.md gap #1 - which could make Processed smaller than Notified + Pending + Duplicate)
    - processed = every bank row Ecobank itself answered SUCCESS for (Cba Rsp Msg), requested in the period - the
      bank statement is the source of truth, whether or not IremboPay ever matched the invoice
    - notified / SLA violated / pending / duplicate = how IremboPay answered each processed row
    - "Duplicate Payment" table = paid with another channel (each invoice once)

CONTENTS (same numbers as im.py; section 7 exists only for Ecobank)
    1. Identity            4. Reading a statement            7. Which rows go in the period and in each table
    2. Statement columns   5. Matching: one row per invoice
    3. Files needed        6. SLA + calculations
"""
import io
import warnings

import pandas as pd


# =====================================================================================================================
# 1. IDENTITY
# =====================================================================================================================
CODE = "ECO"
NAME = "Ecobank"
FULL_NAME = "Ecobank"
PROVIDER = "ECOBANK"                    # payment_provider of Ecobank in the IremboPay payment file
CURRENCIES = ["RWF", "USD"]
FILE_TYPES = ["xlsx", "xls", "csv"]
STATEMENT_SHEET = "Ecobank statement"

# IremboPay statuses
COMPLETED_STATUSES = {"SETTLED", "PAID", "PAYOUT_INITIATED"}
# pending = not paid yet: NEW, or EXPIRED (an invoice the IremboPay check of the review found expired: it is in the
# Pending table and keeps EXPIRED as its status in the report)
PENDING_STATUSES = {"NEW", "EXPIRED"}


# =====================================================================================================================
# 2. STATEMENT COLUMNS (as they are named on the Ecobank statement)
# =====================================================================================================================
COLUMNS = {
    "posted_date": None,                 # the statement has no separate posting date
    "payment_date": "Request Date",      # Kigali time (converted from UTC when the file is read)
    "description": "Bill Ref No",        # the invoice number
    "reference": "Cba Ref No",           # bank reference of the debit
    "amount": "Amount",
}
DETAIL_COLUMNS = ["Source Channel", "Cba Rsp Code", "Biller Rsp Code", "Biller Rsp Msg", "Attempts"]
DATETIME_COLUMNS = ["Request Date"]

UTC_OFFSET = pd.Timedelta(hours=2)       # Request Date is UTC; Kigali = UTC + 2 h
SUCCESS = "000"                          # response code of a successful debit / notification
_HEADER = {"Bill Ref No", "Request Date"}                 # the columns that make a file an Ecobank statement
_TACCT = {"TRANDATE", "NARRATION"}                        # the columns of a transit-account (TACCT) file
_KEEP = ["Source Channel", "Biller Code", "Bill Ref No", "Ccy", "Amount", "Request Date", "Cba Rsp Code",
         "Cba Rsp Msg", "Biller Rsp Code", "Biller Rsp Msg", "Cba Ref No", "Customer Ref No"]
# every column of a row: an export that repeats a row exactly is dropped (Source file excluded)
DUPLICATE_KEY = _KEEP + ["Currency"]


# =====================================================================================================================
# 3. FILES NEEDED (used by the checks of step 3; the settings are explained in banks/__init__.py OPTIONAL)
# =====================================================================================================================
MONDAY_AFTER_WEEKEND = False
DAY_BEFORE_CURRENCIES = []
ONE_FILE_PER_CURRENCY = False      # one statement holds RWF and USD
MISSING_DAYS_BLOCK = False         # a day without statement is a warning, Run is allowed
ACCEPT_EARLIER_DAYS = 31           # statements up to 31 days before: payments of the period may be requested then
REVERSE_CHECK = True               # step 3: IremboPay ECOBANK payments that are on no statement
PAYMENT_FILE_TYPES = ["csv", "xlsx"]
PAYMENT_OPTIONAL_COLUMNS = ["created_at"]

# an unresolved invoice that Ecobank retries in two weeks is pending in both: the weeks can add up to more
TOTALS_CHECK_SKIP = ["Pending notifications"]

# the review recognises an invoice by currency | Request Date | invoice: the Cba Ref No is Ecobank's own number, it is
# not in the IremboPay files and can be missing, so it must never stop a decision from matching (2026-10-12)
REVIEW_KEY_WITHOUT_REFERENCE = True

# SLA: a fixed rule (same minute), so the "SLA days" setting is hidden
SLA_DAYS_SETTING = False
SLA_DIFF_COLUMN = "sla_seconds_difference"


# =====================================================================================================================
# 4. READING A STATEMENT
# =====================================================================================================================
def _reference(value):
    """870602011522, 870602011522.0, ' 870602011522 ' and 8.70602011522e11 -> '870602011522'."""
    if pd.isna(value):
        return pd.NA
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    text = str(value).strip()
    if text.endswith(".0") and text[:-2].isdigit():
        text = text[:-2]
    return text or pd.NA


def _code(value):
    """Response codes 0, '0', 0.0, '000', 400.0, 'E92' -> '000', '400', 'E92'."""
    text = _reference(value)
    if pd.isna(text):
        return ""
    return text.zfill(3) if text.isdigit() else text.upper()


def read_statement(content: bytes, name: str, excel_engine=None):
    """Return (rows, other_credits, info); Request Date is converted to Kigali time.
    Raise ValueError with a plain message when the file is not an Ecobank IREMBO PAY statement."""
    source = io.BytesIO(content) if isinstance(content, (bytes, bytearray)) else content
    try:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")        # Ecobank exports have no style: the Excel reader warns on each file
            if name.lower().endswith(".csv"):
                raw = pd.read_csv(source, header=None, dtype=object)
            else:
                raw = pd.read_excel(source, header=None, dtype=object, engine=excel_engine)
    except Exception:
        raise ValueError(f"{name}: this file cannot be opened as an Ecobank statement.") from None

    # ---- the header row is found in each file (its position changes from one export to another)
    header_idx, names = None, set()
    for idx, row in raw.head(30).iterrows():
        names = {str(v).strip() for v in row if pd.notna(v)}
        if _HEADER <= names:
            header_idx = idx
            break
        if _TACCT <= names:
            raise ValueError(f"{name}: this is an Ecobank transit-account (TACCT) statement; it is not used.")
    if header_idx is None:
        raise ValueError(f"{name}: this is not an Ecobank IREMBO PAY statement (no 'Bill Ref No' and "
                         "'Request Date' columns).")

    # ---- the rows: references and codes cleaned, Request Date in Kigali time
    data = raw.iloc[header_idx + 1:].copy()
    data.columns = [str(c).strip() for c in raw.iloc[header_idx]]
    data = data.loc[:, ~data.columns.duplicated()]
    for col in _KEEP:
        if col not in data.columns:
            data[col] = pd.NA
    data = data[_KEEP].dropna(subset=["Bill Ref No"])
    data["Bill Ref No"] = data["Bill Ref No"].map(_reference)
    data = data.dropna(subset=["Bill Ref No"])
    data["Amount"] = pd.to_numeric(data["Amount"], errors="coerce")
    data["Ccy"] = data["Ccy"].fillna("RWF").astype(str).str.strip().str.upper()
    data["Currency"] = data["Ccy"]
    data["Request Date"] = pd.to_datetime(data["Request Date"], errors="coerce") + UTC_OFFSET
    for col in ("Cba Rsp Code", "Biller Rsp Code"):
        data[col] = data[col].map(_code)
    for col in ("Cba Ref No", "Customer Ref No"):
        data[col] = data[col].map(_reference)
    data = data.dropna(subset=["Request Date"])
    days = data["Request Date"].dt.normalize()
    info = {"currency": ", ".join(sorted(data["Currency"].unique())), "date_from": None, "date_to": None,
            "first": days.min(), "last": days.max()}
    return data.reset_index(drop=True), 0, info


def clean_statement(statement: pd.DataFrame) -> pd.DataFrame:
    """The invoice number is the Bill Ref No."""
    statement = statement.copy()
    statement["invoice_number"] = statement["Bill Ref No"].astype(str)
    return statement


def channel(rows: pd.DataFrame) -> pd.Series:
    """Ecobank channel of a row (AGENCYBANKING, ECOBANKMOBILEAPP, …), shown in the review."""
    return rows["Source Channel"].fillna("OTHER").astype(str)


# =====================================================================================================================
# 5. MATCHING: one row per invoice
# =====================================================================================================================
def one_row_per_invoice(rows: pd.DataFrame, payments: pd.DataFrame) -> pd.DataFrame:
    """The first Ecobank attempt, unless it failed and IremboPay knows the invoice: then the earliest
    successful attempt (Biller Rsp Code 000). Attempts = how many times Ecobank tried the invoice.
    A missing Ecobank amount is taken from IremboPay (as in the old tool)."""
    rows = rows.assign(Attempts=rows.groupby("invoice_number")["invoice_number"].transform("size"))
    irembo_amount = payments.drop_duplicates("invoice_number").set_index("invoice_number")["amount"]
    rows["Amount"] = rows["Amount"].fillna(rows["invoice_number"].map(irembo_amount))
    known = rows["invoice_number"].isin(set(payments["invoice_number"]))
    rows["_rank"] = (known & rows["Biller Rsp Code"].ne(SUCCESS)).astype(int)     # 1 = a failed attempt: kept last
    rows = rows.sort_values(["_rank", "Request Date"], kind="stable")
    return rows.drop_duplicates(subset="invoice_number").drop(columns="_rank")


# =====================================================================================================================
# 6. SLA + CALCULATIONS (one True/False per row)
# =====================================================================================================================
def sla(rows: pd.DataFrame, sla_days: int):
    """(seconds between request and paid_at, SLA status): met when both fall in the same minute.
    Only Ecobank payments are checked; the others are "Not processed". A review decision (sla_override) wins."""
    request, paid = rows["Request Date"], rows["paid_at"]
    diff = (paid - request).dt.total_seconds()
    status = pd.Series("Cannot check", index=rows.index, dtype=object)
    both = request.notna() & paid.notna()
    same_minute = request.dt.floor("min").eq(paid.dt.floor("min"))
    status[both & same_minute] = "On time"
    status[both & ~same_minute] = "SLA violated"
    override = rows["sla_override"].ne("") & rows["payment_provider"].eq(PROVIDER)
    status[override] = rows.loc[override, "sla_override"]
    status[~rows["payment_provider"].eq(PROVIDER)] = "Not processed"
    return diff, status


def bank_success(rows: pd.DataFrame) -> pd.Series:
    """Every invoice the bank itself answered SUCCESS for (Cba Rsp Msg - Ecobank's own response, not Biller Rsp
    Code which is IremboPay's answer to Ecobank): what "processed" starts from (CLAUDE.md: the bank statement is
    the source of truth, never the IremboPay export)."""
    return rows["Cba Rsp Msg"].astype(str).str.strip().str.upper().eq("SUCCESS")


def received(rows: pd.DataFrame) -> pd.Series:
    """Notification received by Irembo: paid through Ecobank and no longer NEW."""
    return (rows["found_in_irembopay"] & rows["payment_provider"].eq(PROVIDER)
            & ~rows["payment_status"].isin(PENDING_STATUSES))


def pending(rows: pd.DataFrame) -> pd.Series:
    """Pending notification: the invoice is still NEW in IremboPay."""
    return rows["payment_status"].isin(PENDING_STATUSES)


def duplicate(rows: pd.DataFrame) -> pd.Series:
    """Paid with another channel: Ecobank received it, but IremboPay says another provider was used."""
    return (rows["found_in_irembopay"] & rows["payment_provider"].notna() & rows["payment_provider"].ne(PROVIDER)
            & ~rows["payment_status"].isin(PENDING_STATUSES))


# =====================================================================================================================
# 7. WHICH ROWS GO IN THE PERIOD AND IN EACH TABLE (I&M uses the engine's default: by payment date)
#    Every table is windowed the same way, by when Ecobank requested the invoice (its own bank time, converted
#    to Kigali) - never by when IremboPay says it was paid: the bank statement is the source of truth
#    (CLAUDE.md gap #1, approved 2026-10-25). Processed = every bank SUCCESS row requested in the period, whether
#    or not IremboPay ever matched it.
# =====================================================================================================================
def _requested(rows, attempts, a, b):
    """Invoices Ecobank sent any attempt for in [a, b): a retried, unresolved invoice stays visible."""
    sent = attempts.loc[(attempts["Request Date"] >= a) & (attempts["Request Date"] < b), "invoice_number"]
    return rows["invoice_number"].isin(set(sent))


def in_period(rows, attempts, start, end) -> pd.Series:
    """A row belongs to the period when Ecobank requested it then - the bank's own Request Date, never IremboPay's
    paid_at."""
    return _requested(rows, attempts, start, end)


def tables(rows, attempts, a, b) -> dict:
    """Rows of the 5 tables for [a, b): all windowed together, by Request Date. Processed = every bank SUCCESS
    row Ecobank requested in the period; the other four tables are how IremboPay answered each processed row."""
    processed = _requested(rows, attempts, a, b) & bank_success(rows)
    return {"processed": processed,
            "received": processed & received(rows),
            "sla": processed & rows["sla_status"].eq("SLA violated"),
            "pending": processed & pending(rows),
            "duplicate": processed & duplicate(rows)}
