"""
BANK OF KIGALI (BK)
===================
Everything that is specific to BK (the common rules are in engine.py). It follows the BK notebook ("bk_recon_final &
weekly report"), with the 4 fixes decided on 03/11/2026 and the credit types agreed on 04/11/2026. I&M, Ecobank, BOA
and BPR are not used or changed.

The BK statement (the files BK sends for each Irembo collection account, .csv or .xlsx)
    - daily:   IREMBO_<account>_STMT_<YYYYMMDD>_<sent>_1.csv/xlsx
    - monthly: IREMBO_ACCOUNT_<account>_MONTHLY_STMT_<MONTH>_<YEAR>_<sent>.xlsx
    - the file name is never used (05/11/2026): the days a file covers are the dates INSIDE it - the H1 row (the
      statement's start), the BookingDate and ValueDate of every movement
    - 18 columns: RowNum, RECTYPE, Currency, AccountNumber, BookingDate, ValueDate, ProcessingDateTime,
      TransactionType, TheirReference, Source, OurReference, DebitAcct, CreditAcct, CreditAmt, DebitAmt, Narration,
      ApiId, Balance. They are found by NAME (some files write them in capitals, in another order, or call them
      RECORDNUM / ACCOUNT_CURRENCY).
    - RECTYPE: H1BKI = opening balance, D1BKIGRW = a movement, F1BKI = closing balance. Only the D1 rows are
      movements (fix 1: the notebook counted H1/F1 too).
    - one file = one account = one currency; only the 5 Irembo collection accounts are accepted (ACCOUNTS).
    - a file opened and saved in Excel loses the long numbers (1.00E+11): it is refused with a plain message.

The type of each row (the notebook's 7 rules + 3 added on 05/11/2026, in order; the last rule that fits wins; no
rule = Transaction)
    1. Source "Card Purchase Advices", or Narration has "Purchase-settlement"         POS TXN
    2. Source is one of BANK_TRANSFER_SOURCES                                         Bank transfer
    3. Source or Narration has "E- commerce purchase"                                 Card payment
    4. Narration has "Purchase-settlement advices"                                    Card payment
    +  Source "VISA SETTLEMENT"                                                       Card payment
    +  Source "In-House Cheques" / "Outward bulk Transfer On-Us", or Narration has
       "EKASH incoming transfer" (cheques, a returned transfer, a wallet transfer)    Bank transfer
    +  Narration has "Cash deposit-BK YACU" (fines paid in cash at an agent; not the
       Branch Teller bill payments, whose Narration is "Cash Deposit : <invoice>")    Cash deposit
    +  Source "Branch Teller", Narration "Cash Deposit" with no 12-digit invoice
       (a name instead: "Cash Deposit : NCGR LTD"; 09/11/2026)                        Cash deposit
    5. Narration has "SWIFT Payment"                                                  Bank transfer
    6. Narration has "SWIFT Payment - Charges"                                        Bank transfer Charges
    7. Narration has "AGENT REFUND"                                                   Refund
    Fix 2: rules 5 and 6 are in this order, so the SWIFT charges stay "Bank transfer Charges" (the notebook had them
    the other way round and turned every charge into a Bank transfer).

How each type is handled (agreed 04/11/2026)
    Transaction (an Irembo bill payment)   matched with IremboPay, checked (SLA, received...), reviewed when not found
    Card payment, POS TXN, Bank transfer,  counted in Processed, never matched nor reviewed, listed in the sheet
    Bank transfer Charges, Cash deposit    "Other credits (BK)" (Bank transfer + Charges also in "Bank transfer (BK)")
    Refund and every other debit           not counted (they are not payments), listed in "Debits (BK)"

Matching a Transaction with IremboPay (the notebook's two steps)
    1. OurReference = an IremboPay provider_reference (only a reference of ONE invoice)
    2. otherwise the invoice number is cut from the Narration, by Source (INVOICE_RULES); only a 12-digit number is
       kept (the notebook kept any 12 characters)
    Fix 3: a Transaction still not found goes to Exception Resolution (the notebook counted it as "other channels"),
    except one with no invoice number at all: never reviewed, straight to Pending notification (05/11/2026).

The figures (Transactions only, except Processed)
    Processed transaction     every credit whose ValueDate is in the period (all types), each invoice once
    Received notification     a Transaction paid with BK, payment_status not NEW (the notebook's rule)
                              fix 4: provider BK, or channel BANK_TRANSFER_BK (IremboPay records some BK transfers
                              under provider IREMBO)
    SLA violated              received, and the paid_at day is not the ValueDate day (0 days allowed by default)
    Pending notification      a Transaction whose invoice is NEW in IremboPay, or one with no invoice number at all
    Paid with other channels  a Transaction paid with another provider, payment_status not NEW
    An invoice credited twice: one credit counts (the one IremboPay recorded, else the first), like the other banks.

Dates: ValueDate sets the period and is compared with paid_at (SLA). BK gives weekend payments a weekend ValueDate.

CONTENTS
    1. Identity            4. Reading a statement        7. SLA
    2. Statement columns   5. Cleaning + the row types   8. Calculations + what goes to review
    3. Files needed        6. Matching with IremboPay    9. The extra report sheets
"""
import io
import re

import pandas as pd


# =====================================================================================================================
# 1. IDENTITY
# =====================================================================================================================
CODE = "BK"
NAME = "BK"
FULL_NAME = "Bank of Kigali"
PROVIDER = "BK"                          # payment_provider in the IremboPay payment file
PROVIDERS = {PROVIDER}
TRANSFER_CHANNEL = "BANK_TRANSFER_BK"    # fix 4: a BK bank transfer, whatever the provider IremboPay recorded
CURRENCIES = ["RWF", "USD", "EUR", "GBP"]
FILE_TYPES = ["csv", "xlsx", "xls"]
STATEMENT_SHEET = "BK statement"

# the 5 Irembo collection accounts and their currency (the only accounts reconciled; 03/11/2026)
ACCOUNTS = {"100081756273": "RWF", "100011044926": "RWF", "100081756467": "USD",
            "100081756548": "EUR", "100081756621": "GBP"}

COMPLETED_STATUSES = {"SETTLED", "PAID", "PAYOUT_INITIATED"}     # same as I&M
PENDING_STATUSES = {"NEW"}               # received = every status but NEW (the notebook's rule)


# =====================================================================================================================
# 2. STATEMENT COLUMNS (as they are named on the BK statement)
# =====================================================================================================================
COLUMNS = {
    "posted_date": "BookingDate",
    "payment_date": "ValueDate",         # sets the period and is compared with paid_at (SLA)
    "description": "Narration",
    "reference": "OurReference",         # = IremboPay provider_reference
    "amount": "CreditAmt",
}
_ALL = ["RowNum", "RECTYPE", "Currency", "AccountNumber", "BookingDate", "ValueDate", "ProcessingDateTime",
        "TransactionType", "TheirReference", "Source", "OurReference", "DebitAcct", "CreditAcct", "CreditAmt",
        "DebitAmt", "Narration", "ApiId", "Balance"]
_NEEDED = ["RECTYPE", "Currency", "AccountNumber", "BookingDate", "ValueDate", "Source", "OurReference", "CreditAmt",
           "DebitAmt", "Narration"]
# a column name with its case, spaces and "_" removed -> its name here (other spellings seen on BK files)
_NAMES = {re.sub(r"[^a-z]", "", c.lower()): c for c in _ALL}
_NAMES.update({"recordnum": "RowNum", "accountcurrency": "Currency", "accountno": "AccountNumber",
               "theirrefence": "TheirReference"})

TYPE = "BK type"
FOUND_BY = "Invoice found by"
DETAIL_COLUMNS = ["AccountNumber", "Source", "ProcessingDateTime", TYPE, FOUND_BY]
DATETIME_COLUMNS = ["ProcessingDateTime"]
# 05/11/2026: the statement sheet shows where each invoice comes from (a row linked by its reference has no invoice in
# its Narration: the invoice is IremboPay's, and provider_reference = OurReference proves the link)
STATEMENT_LINK_COLUMNS = [FOUND_BY, "provider_reference"]
# 05/11/2026: the decision (e.g. "Pending notification (invoice expired)") right after Result, not far on the right
FRONT_COLUMNS = ["Counted credit", "review_status", "review_comment"]
COUNTED_CREDIT_COLUMN = True      # a second credit shows the OurReference of the credit counted for its invoice
# the same row in two files (a daily and the monthly statement): dropped once
DUPLICATE_KEY = ["AccountNumber", "Currency", "OurReference", "BookingDate", "ValueDate", "Narration", "CreditAmt",
                 "DebitAmt"]


# =====================================================================================================================
# 3. FILES NEEDED (used by the checks of step 3)
# =====================================================================================================================
MONDAY_AFTER_WEEKEND = False             # weekend payments have a weekend ValueDate
DAY_BEFORE_CURRENCIES = []
ONE_FILE_PER_CURRENCY = False            # two RWF accounts: the files are checked per account (check_files)
MISSING_DAYS_BLOCK = False               # a day with no statement only warns (a day may have had no movement)
ALL_PROVIDER_REFERENCES = True           # every IremboPay transaction's provider_reference (matching step 1)
SEPARATE_CREDITED_AGAIN = True           # step 5: the extra credits of an invoice in a table of their own
# 05/11/2026: every credit with its IremboPay payment is in "Matched details"; "Payment details" holds the IremboPay
# payment file's own rows of the period (its columns, nothing from the bank)
DETAILS_SHEET = "Matched details"
IREMBOPAY_SHEET = "Payment details"
EXTRA_SHEETS = ["SLA violated (BK)", "Bank transfer (BK)", "Other credits (BK)", "Not found (BK)", "Debits (BK)"]


# =====================================================================================================================
# 4. READING A STATEMENT
# =====================================================================================================================
def _code(values: pd.Series) -> pd.Series:
    """An account / date / reference cell as text: '100081756273.0' -> '100081756273', 'nan' / None -> ''."""
    text = values.astype(object).where(values.notna(), "").astype(str).str.strip()
    return text.str.replace(r"^(\d+)\.0+$", r"\1", regex=True).replace({"nan": "", "NaN": "", "None": ""})


def _number(values: pd.Series) -> pd.Series:
    """'20,000' / '-1350' / '' -> float (NaN when empty)."""
    return pd.to_numeric(_code(values).str.replace(",", "", regex=False), errors="coerce")


def _days(values: pd.Series) -> pd.Series:
    """'20260715' (BK's format) or an Excel date -> Timestamp; anything else -> NaT."""
    text = _code(values)
    out = pd.to_datetime(text, format="%Y%m%d", errors="coerce")
    iso = out.isna() & text.str.match(r"^\d{4}-\d{2}-\d{2}")
    out[iso] = pd.to_datetime(text[iso].str[:10], format="%Y-%m-%d", errors="coerce")
    return out.astype("datetime64[ns]")


def _table(content, name: str, excel_engine):
    """The file as text cells, with the header row found (the first row that has RECTYPE)."""
    data = io.BytesIO(content) if isinstance(content, (bytes, bytearray)) else content
    try:
        if name.lower().endswith(".csv"):
            raw = pd.read_csv(data, header=None, dtype=str, keep_default_na=False, encoding_errors="replace")
        else:
            raw = pd.read_excel(data, header=None, dtype=str, engine=excel_engine)
    except Exception:
        raise ValueError(f"{name}: this file cannot be opened as a BK statement (.csv or .xlsx).") from None
    raw = raw.fillna("").astype(str).apply(lambda col: col.str.strip())
    is_header = raw.head(20).apply(lambda r: any(re.sub(r"[^a-z]", "", v.lower()) == "rectype" for v in r), axis=1)
    if not is_header.any():
        raise ValueError(f"{name}: this is not a BK statement (no RECTYPE column).")
    header_idx = is_header.idxmax()
    columns = [_NAMES.get(re.sub(r"[^a-z]", "", v.lower()), v) for v in raw.loc[header_idx]]
    body = raw.loc[header_idx + 1:].copy()
    body.columns = columns
    missing = [c for c in _NEEDED if c not in body.columns]
    if missing:
        raise ValueError(f"{name}: the BK statement is missing the column(s) {', '.join(missing)}.")
    return body.loc[:, ~body.columns.duplicated()].reset_index(drop=True)


def read_statement(content: bytes, name: str, excel_engine=None):
    """Return (rows, other_credits, info): every movement (D1 row) as text plus Currency, the number of credits that
    are not Irembo bill payments, and info (currency, the days covered, the account in "balances"). Raise ValueError
    with a plain message when the file is not a BK statement of one of the 5 collection accounts."""
    body = _table(content, name, excel_engine)
    rectype = body["RECTYPE"].str.upper()
    data = body[rectype.str.startswith("D1")].copy()
    if data.empty and not rectype.str.startswith(("H1", "F1")).any():
        raise ValueError(f"{name}: this is not a BK statement (no H1 / D1 / F1 rows).")

    accounts = set(_code(body.loc[rectype.str.match(r"^(H1|D1|F1)"), "AccountNumber"])) - {""}
    if any(re.search(r"E\+?\d", a, flags=re.I) for a in accounts):
        raise ValueError(f"{name}: the account number reads {sorted(accounts)[0]}: the file was opened and saved in "
                         "Excel, which changes the long numbers. Upload the file as downloaded from BK.")
    if len(accounts) != 1:
        raise ValueError(f"{name}: a BK statement must be for one account; this file has "
                         f"{len(accounts) or 'no'} account number(s).")
    account = accounts.pop()
    currencies = set(body.loc[rectype.str.match(r"^(H1|D1|F1)"), "Currency"].str.upper()) - {""}
    currency = currencies.pop() if len(currencies) == 1 else ACCOUNTS.get(account, "")
    if account not in ACCOUNTS:
        known = ", ".join(f"{a} ({c})" for a, c in ACCOUNTS.items())
        raise ValueError(f"{name}: account {account}" + (f" ({currency})" if currency else "")
                         + f" is not one of the 5 Irembo collection accounts ({known}); it is not reconciled.")
    if currency != ACCOUNTS[account]:
        raise ValueError(f"{name}: account {account} ({ACCOUNTS[account]}) is a {ACCOUNTS[account]} account, but "
                         f"the file says {currency}.")

    for col in [c for c in _ALL if c not in data.columns]:
        data[col] = ""
    data = data[_ALL].copy()
    data["Currency"] = currency
    data["AccountNumber"] = account
    value = _days(data["ValueDate"])
    if value.isna().any():
        raise ValueError(f"{name}: {int(value.isna().sum())} movement(s) have no readable ValueDate (YYYYMMDD).")
    credit = _number(data["CreditAmt"]).fillna(0)
    other_credits = int(((credit > 0) & row_types(data).ne("Transaction")).sum())
    framed = body[rectype.str.match(r"^(H1|D1|F1)")]
    dates = pd.concat([_days(framed["BookingDate"]), _days(framed["ValueDate"])]).dropna()
    if dates.empty:
        raise ValueError(f"{name}: no date inside the file (BookingDate / ValueDate), so the days it covers are "
                         "unknown.")
    info = {"currency": currency, "date_from": None, "date_to": None, "first": dates.min(), "last": dates.max(),
            "balances": {"account": account, "name": name}}
    return data.reset_index(drop=True), other_credits, info


# =====================================================================================================================
# 5. CLEANING + THE ROW TYPES
# =====================================================================================================================
BANK_TRANSFER_SOURCES = ["Incoming Trsf frm local banks", "Account Transfer", "Payroll Credit", "IB Account transfer",
                         "New App BK-BK Account Transfer"]
# 05/11/2026: credits that are not bill payments and that no notebook rule named (they were Transactions by default)
MORE_BANK_TRANSFER_SOURCES = ["In-House Cheques", "Outward bulk Transfer On-Us"]   # cheques, a returned transfer


def row_types(rows: pd.DataFrame) -> pd.Series:
    """The type of each row: the notebook's rules in order, the last one that fits wins (see the top of the file)."""
    source = rows["Source"].fillna("").astype(str).str.strip()
    narration = rows["Narration"].fillna("").astype(str)
    out = pd.Series("Transaction", index=rows.index, dtype=object)
    out[source.eq("Card Purchase Advices") | narration.str.contains("Purchase-settlement", regex=False)] = "POS TXN"
    out[source.isin(BANK_TRANSFER_SOURCES)] = "Bank transfer"
    out[source.eq("E- commerce purchase") | narration.str.contains("E- commerce purchase", regex=False)] = "Card payment"
    out[narration.str.contains("Purchase-settlement advices", regex=False)] = "Card payment"
    out[source.eq("VISA SETTLEMENT")] = "Card payment"                                         # 05/11/2026
    out[source.isin(MORE_BANK_TRANSFER_SOURCES)
        | narration.str.contains("EKASH incoming transfer", regex=False)] = "Bank transfer"      # 05/11/2026
    out[narration.str.contains("Cash deposit-BK YACU", regex=False)] = "Cash deposit"         # 05/11/2026: fines
    # 09/11/2026: a cash deposit at a branch with a name, not an invoice ("Cash Deposit : NCGR LTD") is not a bill
    # payment (the Branch Teller bill payments are "Cash Deposit : <invoice>")
    teller = source.eq("Branch Teller") & narration.str.contains("Cash Deposit", case=False, regex=False)
    out[teller & narration_invoice(rows)[0].eq("")] = "Cash deposit"
    out[narration.str.contains("SWIFT Payment", regex=False)] = "Bank transfer"
    out[narration.str.contains("SWIFT Payment - Charges", regex=False)] = "Bank transfer Charges"   # fix 2
    out[narration.str.contains("AGENT REFUND", regex=False)] = "Refund"
    return out


def clean_statement(statement: pd.DataFrame) -> pd.DataFrame:
    """Typed columns, the row type, and the invoice number cut from the Narration (matching step 2: the invoice is
    set in find_invoices, after step 1)."""
    s = statement.copy()
    for col in ("CreditAmt", "DebitAmt", "Balance"):
        s[col] = _number(s[col])
    s["CreditAmt"] = s["CreditAmt"].fillna(0.0)
    s["DebitAmt"] = s["DebitAmt"].fillna(0.0)
    s["BookingDate"], s["ValueDate"] = _days(s["BookingDate"]), _days(s["ValueDate"])
    s["ProcessingDateTime"] = pd.to_datetime(_code(s["ProcessingDateTime"]), format="%y%m%d%H%M%S", errors="coerce")
    for col in ("AccountNumber", "OurReference", "Source"):
        s[col] = _code(s[col])
    s[TYPE] = row_types(s)
    s["_narration_invoice"], s["_narration_rule"] = narration_invoice(s)
    s["invoice_number"] = pd.Series(pd.NA, index=s.index, dtype=object)
    s[FOUND_BY] = ""
    return s


def channel(rows: pd.DataFrame) -> pd.Series:
    """The BK channel of a row: its Source (shown in the review)."""
    return rows["Source"].fillna("").astype(str).replace("", "Other")


# =====================================================================================================================
# 6. MATCHING WITH IREMBOPAY
# =====================================================================================================================
SLASH_SOURCES = ["MTN Irembo payments", "Airtel Irembo payments", "New BK Yacu BILL PAYMENT IREMBO",
                 "New bk App Bill Payment IREMBO", "IREMBO NEW MOBISERVE", "BK Yacu Bill Payment IREMBO",
                 "Bill Payment IREMBO", "IREMBO SYBRIN", "IKOFI Bill Payment IREMBO", "Mobiserve Bill Payment IREMBO",
                 "Sybrin Bill Payment IREMBO", "IB Bill Payment IREMBO", "BK App Bill Payment IREMBO"]
_INVOICE = re.compile(r"^\d{12}$")


def narration_invoice(rows: pd.DataFrame):
    """(invoice, rule) cut from the Narration, by Source (the notebook's 3 ways; MTN is cut the third way):
        Bill payment sources (SLASH_SOURCES)  after the last "/", the first 12 characters
        Branch Teller                         after the last ":", the first 12 characters
        MTN Irembo payments                   after the first ":", spaces removed, the last 12 characters
    Only a 12-digit number is kept."""
    source = rows["Source"].fillna("").astype(str).str.strip()
    narration = rows["Narration"].fillna("").astype(str)
    inv = pd.Series("", index=rows.index, dtype=object)
    rule = pd.Series("", index=rows.index, dtype=object)
    slash = source.isin(SLASH_SOURCES) & narration.str.contains("/", regex=False)
    inv[slash] = narration[slash].str.rsplit("/", n=1).str[1].str.strip().str[:12]
    rule[slash] = "Narration (after the last /)"
    branch = source.eq("Branch Teller") & narration.str.contains(":", regex=False)
    inv[branch] = narration[branch].str.rsplit(":", n=1).str[1].str.strip().str[:12]
    rule[branch] = "Narration (after the last :)"
    mtn = source.eq("MTN Irembo payments") & narration.str.contains(":", regex=False)
    inv[mtn] = narration[mtn].str.split(":", n=1).str[1].str.replace(" ", "", regex=False).str[-12:]
    rule[mtn] = "Narration (after the first :)"
    ok = inv.str.match(_INVOICE)
    return inv.where(ok, ""), rule.where(ok, "")


def find_invoices(statement: pd.DataFrame, payments: pd.DataFrame) -> pd.DataFrame:
    """(engine hook) The invoice of each Transaction credit: 1. OurReference = an IremboPay provider_reference (a
    reference of ONE invoice only), 2. otherwise the one cut from the Narration (kept even when IremboPay's file does
    not have it: it is then reviewed with its number). The way it was found is kept ("Invoice found by")."""
    s = statement.copy()
    refs = payments[["invoice_number"]].copy()
    refs["_r"] = (payments["_provider_refs"].fillna("").astype(str) if "_provider_refs" in payments
                  else pd.Series("", index=payments.index))
    refs["_r"] = refs["_r"] + "|" + payments["provider_reference"].fillna("").astype(str).str.strip()
    refs = refs.assign(_r=refs["_r"].str.split("|")).explode("_r")
    refs = refs[refs["_r"].ne("")].drop_duplicates()
    per_ref = refs.groupby("_r")["invoice_number"].nunique()
    by_ref = refs[refs["_r"].map(per_ref).eq(1)].drop_duplicates("_r").set_index("_r")["invoice_number"]

    todo = s[TYPE].eq("Transaction") & (s["CreditAmt"] > 0)
    from_ref = s["OurReference"].map(by_ref)
    step1 = todo & from_ref.notna()
    s.loc[step1, "invoice_number"] = from_ref[step1]
    s.loc[step1, FOUND_BY] = "OurReference = provider_reference"
    step2 = todo & ~step1 & s["_narration_invoice"].ne("")
    s.loc[step2, "invoice_number"] = s.loc[step2, "_narration_invoice"]
    s.loc[step2, FOUND_BY] = s.loc[step2, "_narration_rule"]
    return s


# =====================================================================================================================
# 7. SLA: the paid_at day against the ValueDate (0 days allowed by default, "Settings" in the side panel)
# =====================================================================================================================
def _transaction(rows: pd.DataFrame) -> pd.Series:
    return rows[TYPE].eq("Transaction")


def _is_bk(rows: pd.DataFrame) -> pd.Series:
    """Paid with BK: provider BK, or a BK bank transfer (channel BANK_TRANSFER_BK, fix 4)."""
    provider = rows["payment_provider"].fillna("").astype(str).str.strip().str.upper()
    chan = (rows["payment_channel"].fillna("").astype(str).str.strip().str.upper() if "payment_channel" in rows
            else pd.Series("", index=rows.index))
    return provider.isin(PROVIDERS) | chan.eq(TRANSFER_CHANNEL)


def _status(rows: pd.DataFrame) -> pd.Series:
    return rows["payment_status"].fillna("").astype(str).str.strip().str.upper()


def sla(rows: pd.DataFrame, sla_days: int):
    """(days difference, SLA status) per row: the IremboPay paid_at day against the ValueDate. Only the Transactions
    paid with BK are checked; the other Transactions are "Not processed", the other types "Not applicable".
    A review decision (sla_override) wins."""
    diff = (rows["paid_at"].dt.normalize() - rows["ValueDate"]).dt.days
    status = pd.Series("Cannot check", index=rows.index, dtype=object)
    has_date = diff.notna()
    within = diff.abs() <= sla_days
    status[has_date & within] = "On time"
    status[has_date & ~within] = "SLA violated"
    bk = _is_bk(rows)
    override = rows["sla_override"].ne("") & bk
    status[override] = rows.loc[override, "sla_override"]
    status[~bk] = "Not processed"
    status[~_transaction(rows)] = "Not applicable"
    return diff, status


# =====================================================================================================================
# 8. CALCULATIONS + WHAT GOES TO REVIEW (one True/False per credit; Processed = every credit, each invoice once)
# =====================================================================================================================
def received(rows: pd.DataFrame) -> pd.Series:
    """Notification received: a Transaction found in IremboPay, paid with BK, payment_status not NEW."""
    status = _status(rows)
    return (_transaction(rows) & rows["found_in_irembopay"].astype(bool) & _is_bk(rows)
            & status.ne("") & ~status.isin(PENDING_STATUSES))


def pending(rows: pd.DataFrame) -> pd.Series:
    """Pending notification: a Transaction whose invoice is NEW in IremboPay, or a bill payment with no invoice number
    at all (05/11/2026: the money is on the statement, IremboPay has no record of it - its notification is awaited)."""
    found = rows["found_in_irembopay"].astype(bool)
    return _transaction(rows) & ((found & _status(rows).isin(PENDING_STATUSES))
                                 | (~found & rows["invoice_number"].isna()))


def duplicate(rows: pd.DataFrame) -> pd.Series:
    """Paid with other channels: a Transaction found in IremboPay, paid with another provider, status not NEW."""
    status = _status(rows)
    return (_transaction(rows) & rows["found_in_irembopay"].astype(bool) & ~_is_bk(rows)
            & status.ne("") & ~status.isin(PENDING_STATUSES))


def in_period(rows, attempts, start, end) -> pd.Series:
    """The credits whose ValueDate is in the period (a debit is never a payment: it is listed in "Debits (BK)")."""
    return (rows["CreditAmt"] > 0) & rows["ValueDate"].between(start, end, inclusive="left")


NO_INVOICE = "Pending notification (no invoice number)"


def counted_not_reviewed(rows: pd.DataFrame) -> pd.Series:
    """(engine hook) Why a credit is counted in Processed but never matched nor reviewed ("" = handled normally):
    card payments, POS, bank transfers and their charges are not Irembo bill payments; a bill payment with no
    invoice number (nothing after "IREMBO :", reference not in IremboPay) is never reviewed (04/11/2026: nobody can
    look it up) and goes straight to Pending notification (05/11/2026, see pending)."""
    t = rows[TYPE].fillna("Transaction").astype(str)
    out = ("Not an Irembo bill payment (" + t + ")").where(t.ne("Transaction"), "")
    out[t.eq("Transaction") & rows["invoice_number"].isna()] = NO_INVOICE
    return out


RESULTS = ["Matched", "Second credit", "Not found", "No invoice number", "Not a bill payment"]


def row_result(rows: pd.DataFrame) -> pd.Series:
    """(engine hook, 05/11/2026) What happened to each credit, in one word (the "Result" column of Matched details):
        Matched              a bill payment matched with an IremboPay payment
        Second credit        another credit of an invoice already counted (not counted again)
        Not found            a bill payment whose invoice is not in IremboPay's file (decided in Exception Resolution;
                             what was decided is in review_status)
        No invoice number    a bill payment with no invoice at all (straight to Pending notification)
        Not a bill payment   a card payment, POS, bank transfer, charge or cash deposit (counted, not checked)"""
    lost = rows["originally_not_found"] if "originally_not_found" in rows else ~rows["found_in_irembopay"]
    out = pd.Series("Matched", index=rows.index, dtype=object)
    out[lost.astype(bool)] = "Not found"
    out[lost.astype(bool) & rows["invoice_number"].isna()] = "No invoice number"
    out[rows["is_extra_credit"].astype(bool)] = "Second credit"
    out[~_transaction(rows)] = "Not a bill payment"
    return out


def review_reasons(rows: pd.DataFrame) -> pd.Series:
    """Why a credit FOUND in IremboPay still goes to Exception Resolution ("" = it does not). Only one case, like the
    other banks: an invoice credited more than once - the credit IremboPay recorded counts (OurReference =
    provider_reference), otherwise the earliest; every other credit of that invoice goes to the step-5 table
    "Invoice credited more than once"."""
    out = pd.Series("", index=rows.index, dtype=object)
    t = rows[_transaction(rows) & rows["invoice_number"].notna()]
    t = t[t["invoice_number"].duplicated(keep=False)]
    if t.empty:
        return out
    t = t.assign(_r=t[FOUND_BY].ne("OurReference = provider_reference").astype(int))
    t = t.sort_values(["invoice_number", "_r", "ValueDate", "ProcessingDateTime", "OurReference"], kind="stable")
    counted = t.drop_duplicates("invoice_number")
    kept = counted.set_index("invoice_number")["OurReference"].astype(str)
    others = t.index.difference(counted.index)
    out[others] = "Invoice credited again: the credit counted is " + t.loc[others, "invoice_number"].map(kept)
    return out


def check_files(statements: list) -> tuple:
    """(days covered, lines, blocking): the files of each of the 5 accounts. An account with no file blocks the run
    (the notebook always reconciles the 5); a day with no file only warns (a day may have had no movement)."""
    covered, lines, blocking = set(), [], []
    per_account = {}
    for f in statements:
        days = set(pd.date_range(f.first, f.last))
        covered |= days
        account = (getattr(f, "balances", None) or {}).get("account", "")
        per_account.setdefault(account, set()).update(days)
    if not covered:
        return covered, lines, blocking
    every_day = set(pd.date_range(min(covered), max(covered)))
    for account, currency in ACCOUNTS.items():
        label = f"Account {account} ({currency})"
        if account not in per_account:
            lines.append((False, f"{label}: no file added → upload its statement"))
            blocking.append(f"{label}: no file added.")
            continue
        gaps = sorted(every_day - per_account[account])
        if gaps:
            days = ", ".join(sorted({f"{d:%d/%m}" for d in gaps}))
            lines.append((False, f"{label}: no file for {days}"))
        else:
            lines.append((True, f"{label}: {min(every_day):%d/%m} to {max(every_day):%d/%m} covered"))
    return covered, lines, blocking


# =====================================================================================================================
# 9. THE EXTRA REPORT SHEETS (the notebook's sheets, with names that say what they hold)
# =====================================================================================================================
def _time_diff(delta: pd.Series) -> pd.Series:
    """|paid_at - ValueDate| as the notebook writes it: '05hrs 12mins 03sec 01days' (empty when unknown)."""
    def one(d):
        if pd.isna(d):
            return ""
        s = int(abs(d).total_seconds())
        days, s = divmod(s, 86400)
        hours, s = divmod(s, 3600)
        mins, secs = divmod(s, 60)
        parts = [f"{hours:02d}hrs" if hours else "", f"{mins:02d}mins" if mins else "", f"{secs:02d}sec" if secs else "",
                 f"{days:02d}days" if days else ""]
        return " ".join(p for p in parts if p) or "0sec"
    return delta.map(one)


def extra_sheets(rep, attempts, start, end, whole) -> dict:
    """(engine hook) The BK sheets of the report:
        SLA violated (BK)    the SLA violated of the period, with Time diff and Charge applicable (notebook)
        Bank transfer (BK)   the Bank transfer and Bank transfer Charges credits of the period
        Other credits (BK)   every credit that is not an Irembo bill payment (counted in Processed, not checked)
        Not found (BK)       the Transactions not found in IremboPay, with what the review decided
        Debits (BK)          the debits of the period, refunds included (not counted: not payments)"""
    stmt = ["Currency", "AccountNumber", "BookingDate", "ValueDate", "ProcessingDateTime", "Source", "OurReference",
            "Narration", "CreditAmt", TYPE]
    late = rep[whole["sla"]].copy()
    late["Time diff"] = _time_diff(late["paid_at"] - late["ValueDate"])
    late["Charge applicable"] = "Yes"
    sla_cols = ["invoice_number", "Currency", "CreditAmt", "amount", "Source", "OurReference", "payment_status",
                "payment_channel", "name", "ValueDate", "paid_at", "Time diff", "Charge applicable"]
    others = rep[rep[TYPE].ne("Transaction")]
    transfers = others[others[TYPE].isin(["Bank transfer", "Bank transfer Charges"])]
    lost = rep[_transaction(rep) & rep.get("originally_not_found", ~rep["found_in_irembopay"]).astype(bool)
               & rep["invoice_number"].notna()]            # no invoice number: in Pending notification (05/11/2026)
    lost_cols = stmt + ["invoice_number", FOUND_BY, "found_in_irembopay", "payment_status", "payment_provider",
                        "review_status", "review_comment", "Source file"]
    debit = attempts[(attempts["CreditAmt"] <= 0) & attempts["ValueDate"].between(start, end, inclusive="left")].copy()
    debit[TYPE] = debit[TYPE].where(debit[TYPE].ne("Transaction"), "Debit")      # money out, not a bill payment
    debit_cols = ["Currency", "AccountNumber", "BookingDate", "ValueDate", "ProcessingDateTime", "Source",
                  "OurReference", "Narration", "DebitAmt", TYPE, "Source file"]

    def cols(df, wanted):
        return df[[c for c in wanted if c in df.columns]]

    return {"SLA violated (BK)": cols(late, sla_cols).sort_values("ValueDate"),
            "Bank transfer (BK)": cols(transfers, stmt + ["Source file"]).sort_values("ValueDate"),
            "Other credits (BK)": cols(others, stmt + ["Source file"]).sort_values([TYPE, "ValueDate"]),
            "Not found (BK)": cols(lost, lost_cols).sort_values("ValueDate"),
            "Debits (BK)": cols(debit, debit_cols).sort_values("ValueDate")}
