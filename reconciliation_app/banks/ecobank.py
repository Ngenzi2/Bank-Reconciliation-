"""
ECOBANK
=======
Everything that is specific to Ecobank (the common rules are in engine.py). The rules come from the old Ecobank
tool (~/Irembo reconciliation/Ecobank-Reconciliation) and give the same figures.

The Ecobank statement ("IREMBO PAY" export) is a log of payment REQUESTS, not of credits
    - Request Date = when the customer paid at Ecobank, in UTC; Kigali time = UTC + 2 h
    - one invoice can be tried many times (failed attempts, retries): one row per invoice is kept (section 5)
    - one file holds every currency (Ccy column); a file's day is read from its Request Dates, not its name
    - TACCT (transit-account) files are a different flow: they are refused

The report follows the old tool
    - processed (26/10/2026) = the requests Ecobank answered SUCCESS (Cba Rsp Code 000), by Request Date, each
      invoice once (a second SUCCESS = debited twice: Exception Resolution)
    - received = in IremboPay, paid (paid_at) through ECOBANK; another provider = Duplicate Payment
    - every invoice is looked up on IremboPay's checkout before the reconciliation (CHECKOUT_ALL)
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
DRIVE_STATEMENTS = "ecobank_statements_folder_id"   # 26/10/2026: step 2 can read the statements from this Drive
                                                    # folder (drive_settings.toml), picked by the days inside them
DRIVE_STATEMENTS_PREFIX = "IREMBO PAY"              # only files named like this are considered (the folder also
                                                    # has IREMBO_TACCT_TRANS_... and other exports: not statements)
STATEMENT_SHEET = "Ecobank statement"

# IremboPay statuses
COMPLETED_STATUSES = {"SETTLED", "PAID", "PAYOUT_INITIATED"}
# pending = not paid yet: payment_status NEW (20/10/2026: "expired" is not a payment status - an expired
# invoice was never paid, it stays NEW; the expiry is a note in the report)
PENDING_STATUSES = {"NEW"}


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
CHECKOUT_ALL = True                # 26/10/2026: every invoice goes through IremboPay's checkout before the run
SEPARATE_CREDITED_AGAIN = True     # step 5: an invoice debited twice (second SUCCESS) in a table of its own
PAYMENT_FILE_TYPES = ["csv", "xlsx"]
PAYMENT_OPTIONAL_COLUMNS = ["created_at"]


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
    """26/10/2026: every request Ecobank answered SUCCESS (Cba Rsp Code 000: the customer was debited) is kept - the
    first one counts, a second one is an extra debit (Exception Resolution); an invoice with no SUCCESS keeps its
    first attempt (not processed). Attempts = how many times Ecobank tried the invoice. A missing Ecobank amount is
    taken from IremboPay (as in the old tool).
    26/10/2026: a SUCCESS with the same invoice AND the same Cba Ref No as an earlier one is the same debit sent
    again by Ecobank (a resend), not a second debit: only the earliest is kept. Only a different Cba Ref No is an
    extra debit."""
    rows = rows.assign(Attempts=rows.groupby("invoice_number")["invoice_number"].transform("size"))
    irembo_amount = payments.drop_duplicates("invoice_number").set_index("invoice_number")["amount"]
    rows["Amount"] = rows["Amount"].fillna(rows["invoice_number"].map(irembo_amount))
    ok = rows["Cba Rsp Code"].eq(SUCCESS)
    has_ok = rows["invoice_number"].map(ok.groupby(rows["invoice_number"]).any())
    rows = rows.sort_values("Request Date", kind="stable")
    success = rows[ok[rows.index]]
    ref = success["Cba Ref No"].fillna("").astype(str).str.strip()
    resend = ref.ne("") & success.assign(_ref=ref).duplicated(subset=["invoice_number", "_ref"], keep="first")
    success = success[~resend]
    first_try = rows[~has_ok[rows.index].astype(bool)].drop_duplicates(subset="invoice_number")
    return pd.concat([success, first_try]).sort_values("Request Date", kind="stable")


def is_success(rows: pd.DataFrame) -> pd.Series:
    """Ecobank answered SUCCESS for the request: Cba Rsp Code 000 (the customer was debited)."""
    return rows["Cba Rsp Code"].fillna("").astype(str).eq(SUCCESS)


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


def received(rows: pd.DataFrame) -> pd.Series:
    """Notification received: in IremboPay, paid (paid_at filled, whatever the status) through Ecobank."""
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
    """Why a request FOUND in IremboPay still goes to Exception Resolution ("" = it does not). 26/10/2026: an
    invoice Ecobank answered SUCCESS more than once (debited twice) - the first SUCCESS counts, the others are
    checked."""
    out = pd.Series("", index=rows.index, dtype=object)
    # 26/10/2026: also an invoice NOT found in IremboPay - its second credit goes to "credited more than once" too
    found = rows[rows["invoice_number"].notna() & is_success(rows)]
    for _, group in found.groupby("invoice_number"):
        if len(group) < 2:
            continue
        group = group.sort_values("Request Date", kind="stable")
        kept = str(group["Cba Ref No"].iloc[0] if pd.notna(group["Cba Ref No"].iloc[0]) else group["Request Date"].iloc[0])
        out[group.index[1:]] = f"Invoice credited again: debited again by Ecobank; the request counted is {kept}"
    return out


# =====================================================================================================================
# 7. WHICH ROWS GO IN THE PERIOD AND IN EACH TABLE (I&M uses the engine's default: by payment date)
# =====================================================================================================================
def in_period(rows, attempts, start, end) -> pd.Series:
    """26/10/2026: a request belongs to the period of its Request Date (the bank date, Kigali time)."""
    return (rows["Request Date"] >= start) & (rows["Request Date"] < end)


def tables(rows, attempts, a, b) -> dict:
    """Rows of the 5 tables for [a, b) (26/10/2026): Processed = the requests Ecobank answered SUCCESS with a Request
    Date in [a, b), each invoice once; Received / Duplicate / Pending from IremboPay's paid_at and provider."""
    window = (rows["Request Date"] >= a) & (rows["Request Date"] < b)
    processed = window & is_success(rows)
    if "is_extra_credit" in rows:
        processed &= ~rows["is_extra_credit"]
    received_ = processed & received(rows)
    duplicate_ = processed & duplicate(rows) & ~received_
    pending_ = processed & pending(rows) & ~received_ & ~duplicate_
    return {"processed": processed, "received": received_,
            "sla": processed & rows["sla_status"].eq("SLA violated"),
            "pending": pending_, "duplicate": duplicate_}
