"""
RECONCILIATION ENGINE
=====================
The rules that are the SAME FOR EVERY BANK. The app (app.py) calls these functions; the developer notebook
(I_M_reconciliation.ipynb) uses the review functions of section 7.

How the code is split
    app.py           the screens (6 steps + side panel); no reconciliation rules
    engine.py        this file: the common rules
    banks/<bank>.py  what is specific to one bank: statement format, matching, SLA, calculations
                     (banks/__init__.py lists the banks and explains how to add one)

What happens in one reconciliation (the order of this file)
    1. choose the bank and the period (day / week / month) ............ sections 1, 3
    2. read the uploaded files in memory (statements + IremboPay file) . section 5
    3. check that every day is covered before Run ...................... section 6
    4. run: match the bank rows with IremboPay, review the rows that are
       not found, check the SLA, build the 5 summary tables ............ sections 7, 9
    5. compare with the previous report (optional) .................... section 8
    6. write the Excel report, log the run, make the PDF ............... sections 10, 11, 13

What is saved (nothing else: uploaded files are never kept)
    output/<day|week|month>/<CODE>_recon_summary_<period>.xlsx   the reports  (CODE = IM, ECO, ...)
    output/.review_store.csv                                       the review decisions (shared with the notebook)
    output/run_log.csv                                             one line per run

CONTENTS
    1.  Errors and banks
    2.  Settings
    3.  Periods
    4.  File names and columns
    5.  Reading the uploaded files
    6.  Checks before Run
    7.  Review decisions
    8.  Comparison with the previous report
    9.  The reconciliation run
    10. Run log
    11. Writing the Excel report
    12. Saved reports and decisions (side panel)
    13. PDF (summary only)
    14. Custom report (step 6: sections, charts and layout chosen by the user)

Words used in the code
    statement row   one line of a bank statement (one payment, or for Ecobank one payment attempt)
    payment date    the bank's date of the payment (I&M: Value Date, Ecobank: Request Date); sets the period
    not found       a statement row whose invoice is not in the IremboPay file: it must be reviewed
    block           one summary of 5 tables (a day, a week or a month); a monthly report has one block per week
                    + one for the month
"""
from __future__ import annotations

import hashlib
import io
import itertools
import re
import time
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import banks


# =====================================================================================================================
# 1. ERRORS AND BANKS
#    ReconError = a problem shown to the user in plain words. The banks come from banks/ (see banks/__init__.py).
# =====================================================================================================================
class ReconError(Exception):
    """A problem explained in plain words for the user."""


BANKS = banks.load()                       # {code: bank module}; raises ValueError if a bank file is incomplete
DEFAULT_BANK = banks.ACTIVE[0].CODE        # I&M: used when no bank is given (e.g. by the notebook)


def get_bank(bank=None):
    """A bank module from its code (or the module itself); None = the default bank (I&M)."""
    if bank is None:
        return BANKS[DEFAULT_BANK]
    if not isinstance(bank, str):
        return bank
    if bank not in BANKS:
        raise ReconError(f"Unknown bank '{bank}'.")
    return BANKS[bank]


def bank_name(code: str) -> str:
    """'IM' -> 'I&M' (an unknown code is returned as it is)."""
    return BANKS[code].NAME if code in BANKS else code


def option(bank, name, default):
    """Optional bank setting (see banks/__init__.py OPTIONAL); the default is how I&M works."""
    return getattr(get_bank(bank), name, default)


def providers(bank) -> set:
    """The payment_provider value(s) that count as this bank: {PROVIDER}, or PROVIDERS when a bank uses more than
    one code for itself (e.g. BPR: "BPR" and "BPR_RW")."""
    bank = get_bank(bank)
    return {p.upper() for p in option(bank, "PROVIDERS", [bank.PROVIDER])}


def present(*cols):
    """Columns that the bank has (a bank may have no posted date)."""
    return [c for c in cols if c]


# =====================================================================================================================
# 2. SETTINGS
#    Column names of the IremboPay file, the review choices, the report sheets and file names.
# =====================================================================================================================
# ---- re-sends: a bank statement row whose invoice was already credited this many days before the period is still
#      read (to recognise the re-send), but excluded from Processed (see run(), step C)
RESEND_LOOKBACK_DAYS = 45

# ---- IremboPay payment file
TEXT_COLUMNS = ["invoice_number", "transaction_id", "provider_reference", "transaction_reference"]  # kept as text
DATE_COLUMNS = ["created_at", "paid_at", "settled_at", "expiry_at"]                                  # read as dates
PAYMENT_REQUIRED = ["invoice_number", "amount", "currency", "payment_status", "paid_at",
                    "payment_provider", "transaction_status", "created_at", "name"]                  # must be there

# ---- review (step 4 of the app): the 5 choices and the columns the user fills in
TABLE_CHOICES = ["Notification received", "Pending notification", "Paid with other channel",
                 "Failed notification", "Exclude"]
# Pending notification: two statuses besides the bank's own PENDING_STATUSES (dianah branch, automatic
# classification, services/auto_review.py) - a bank credit with no invoice number, and one whose IremboPay
# checkout answer could not be had (it is still Pending; "Needs a look" shows it, but it never blocks the report)
NO_INVOICE_STATUS = "no invoice number"
COULD_NOT_CHECK_STATUS = "Could not check"
EDIT_COLS = ["Reviewed (Yes/No)", "Report table", "SLA violated (Yes/No)", "Correct invoice number",
             "payment_status", "payment_provider", "paid_at (YYYY-MM-DD HH:MM:SS)", "Merchant name", "Comment"]
SUGGEST_COL = "Suggestion from"            # filled when an earlier review already decided the invoice
STORE_COLS = ["bank", "period", "key", "invoice_number"] + EDIT_COLS + ["saved_at"]   # columns of the review store
STORE_NAME = ".review_store.csv"

# ---- report: the 5 summary tables, the sheets, the file names
TABLE_TITLES = ["Overall processed payments", "Overall notifications received by Irembo", "SLA Violated",
                "Pending notifications", "Duplicate Payment"]
SUMMARY_SHEET = {"day": "Summary-Daily", "week": "Summary-Weekly", "month": "Summary-Monthly"}
# the 6th table of each summary block (2026-10-12): the notifications received, by merchant and currency
MERCHANT_TITLE = "Payment per merchant"
OLD_SHEETS = {"Payment per merchant", "Debited more than once"}   # Ecobank sheets of reports made before: removed
NOT_FOUND_SHEET = "not found in IremboPay"     # only in old reports (removed when they are made again)
REPORT_SHEETS = ({"Summary-Daily", "Summary-Weekly", "Summary-Monthly", "Pending notification",
                  "Duplicate payment other channel", NOT_FOUND_SHEET, "Payment details"}
                 | {b.STATEMENT_SHEET for b in BANKS.values()})          # sheets written by the app (the rest = user's)
REPORT_SUFFIX = "_recon_summary_"              # report file = <CODE>_recon_summary_<period label>.xlsx

# ---- fast Excel reading when python-calamine is installed (Setup installs it); pandas' default otherwise
try:
    import python_calamine  # noqa: F401
    EXCEL_ENGINE = "calamine"
except ImportError:
    EXCEL_ENGINE = None


# =====================================================================================================================
# 3. PERIODS
#    A period is a day, a week (Monday to Sunday) or a month; `end` is the midnight after the last day (NOT included).
#    Only a period that has ended can be reconciled (period_not_ended).
# =====================================================================================================================
@dataclass
class Period:
    report_type: str           # "day", "week" or "month"
    start: pd.Timestamp        # first day
    end: pd.Timestamp          # not included
    unit: str                  # "Day", "Week" or "Month" (used in "VS prev <unit>")

    @property
    def label(self):
        """'01-09-2026_to_30-09-2026' (first day to last day): used in report file names."""
        return f"{self.start:%d-%m-%Y}_to_{self.last_day:%d-%m-%Y}"

    @property
    def store_key(self):
        """'month:01-09-2026_to_01-10-2026': the period of a review decision. It keeps the day AFTER the period as
        its end: the format of the decisions already saved."""
        return f"{self.report_type}:{self.start:%d-%m-%Y}_to_{self.end:%d-%m-%Y}"

    @property
    def text(self):
        """'01/09/2026 - 30/09/2026': used in the run log."""
        return f"{self.start:%d/%m/%Y} - {self.last_day:%d/%m/%Y}"

    @property
    def last_day(self):
        return self.end - pd.Timedelta(days=1)

    @property
    def name(self):
        """How the period is shown to the user: 'September 2026', 'week 07/09 - 13/09/2026', ..."""
        return {"day": f"{self.start:%A %d %B %Y}",
                "week": f"week {self.start:%d/%m} - {self.last_day:%d/%m/%Y}",
                "month": f"{self.start:%B %Y}"}[self.report_type]

    @property
    def day_before(self):
        return self.start - pd.Timedelta(days=1)

    @property
    def monday_after(self):
        """Monday after the period when it ends on a weekend (late weekend credits are posted on Monday)."""
        if self.last_day.dayofweek >= 5:
            return self.last_day + pd.Timedelta(days=7 - self.last_day.dayofweek)
        return None

    def previous(self) -> "Period":
        """The period just before (used by the comparison)."""
        return make_period(self.report_type, {"day": self.start - pd.Timedelta(days=1),
                                              "week": self.start - pd.Timedelta(days=7),
                                              "month": self.start - pd.offsets.MonthBegin(1)}[self.report_type])


def make_period(report_type: str, date) -> Period:
    """The day / week / month that contains `date`."""
    day = pd.Timestamp(date).normalize()
    if report_type == "day":
        return Period("day", day, day + pd.Timedelta(days=1), "Day")
    if report_type == "week":
        start = day - pd.Timedelta(days=day.dayofweek)          # back to Monday
        return Period("week", start, start + pd.Timedelta(days=7), "Week")
    if report_type == "month":
        start = day.replace(day=1)
        return Period("month", start, start + pd.offsets.MonthBegin(1), "Month")
    raise ReconError('Report type must be "day", "week" or "month".')


def period_not_ended(period: Period, today=None) -> str | None:
    """Why a period cannot be reconciled yet, or None when it has ended. A period ends at midnight after its last
    day: a day the next day, a week (Monday to Sunday) on the next Monday, a month on the 1st of the next month."""
    today = pd.Timestamp(today if today is not None else pd.Timestamp.today()).normalize()
    if period.end <= today:
        return None
    what = {"day": f"The day {period.start:%A %d/%m/%Y}",
            "week": f"The week {period.start:%d/%m/%Y} – {period.last_day:%d/%m/%Y}",
            "month": f"The month {period.start:%B %Y}"}[period.report_type]
    return f"{what} has not ended yet.\nIt can be reconciled from {period.end:%A %d/%m/%Y}."


def extra_monday(period: Period, bank=None):
    """Monday after the period, when the bank needs it (period ending on a weekend), else None."""
    return period.monday_after if get_bank(bank).MONDAY_AFTER_WEEKEND else None


# =====================================================================================================================
# 4. FILE NAMES AND COLUMNS
# =====================================================================================================================
def report_path(base_dir: Path, period: Period, bank=None, root: Path | None = None) -> Path:
    """Where the report of a bank and period is written: <root>/<day|week|month>/<CODE>_recon_summary_<label>.xlsx.
    root: output/ by default; the app gives its session's temporary folder (reports are kept on Google Drive)."""
    code = get_bank(bank).CODE
    root = Path(root) if root is not None else Path(base_dir) / "output"
    return root / period.report_type / f"{code}{REPORT_SUFFIX}{period.label}.xlsx"


def report_bank(path: Path) -> str:
    """Bank code from a report file name ('IM_recon_summary_...' -> 'IM')."""
    return Path(path).stem.split(REPORT_SUFFIX)[0]


def report_label(path: Path) -> str:
    """Period label from a report file name ('IM_recon_summary_01-09-2026_to_30-09-2026' -> '01-09-2026_to_...')."""
    return Path(path).stem.split(REPORT_SUFFIX, 1)[-1]


def bank_cols(bank=None) -> list:
    """Columns of a bank row shown in the review (not found in IremboPay)."""
    c = get_bank(bank).COLUMNS
    return (["Currency"] + present(c["payment_date"], c["posted_date"], c["description"], c["reference"], c["amount"])
            + ["invoice_number", "Bank channel", "Reason", "Source file"])


# =====================================================================================================================
# 5. READING THE UPLOADED FILES (in memory, never saved)
#    A bank statement is read by the bank's own read_statement(); the IremboPay file is the same for every bank.
# =====================================================================================================================
_FILE_ORDER = itertools.count()            # read order (section 9: across several payment files, the newest wins)


@dataclass
class UploadedFile:
    name: str
    kind: str                  # "bank" or "payments"
    digest: str                # fingerprint of the content: the same file added twice is recognised
    data: pd.DataFrame
    currency: str = ""
    first: pd.Timestamp | None = None      # first day covered
    last: pd.Timestamp | None = None       # last day covered
    rows: int = 0
    other_credits: int = 0                 # credits on the statement that are not IremboPay payments
    bank: str = ""                         # bank code (bank files only)
    providers: frozenset | None = None     # payment_provider values found (payment files only, e.g. {"I_M"})
    reference: bool = False    # bank file read only to spot re-sends (section 9); never required, never counted
    order: float = field(default_factory=lambda: next(_FILE_ORDER))   # read order, or a file's own mtime (newer wins)
    taken_at: pd.Timestamp | None = None   # payments file: when it was downloaded (folder-found files only)


def file_digest(content: bytes) -> str:
    """Fingerprint of a file's content (two files with the same content have the same fingerprint)."""
    return hashlib.sha1(content).hexdigest()


def read_bank_file(name: str, content: bytes, bank=None) -> UploadedFile:
    """Read one bank statement with the bank's rules; the days it covers come from its header dates, else from
    the payment dates inside."""
    bank = get_bank(bank)
    try:
        data, other_credits, info = bank.read_statement(content, name, EXCEL_ENGINE)
    except ValueError as err:                                     # the bank file explains what is wrong
        raise ReconError(str(err)) from None
    data["Source file"] = name
    first = info["date_from"] if pd.notna(info["date_from"]) else info["first"]
    last = info["date_to"] if pd.notna(info["date_to"]) else info["last"]
    if pd.isna(first) or pd.isna(last):
        raise ReconError(f"{name}: no dates found in this bank file.")
    return UploadedFile(name, "bank", file_digest(content), data, info["currency"],
                        pd.Timestamp(first).normalize(), pd.Timestamp(last).normalize(), len(data), other_credits,
                        bank.CODE)


def read_payment_file(name: str, content: bytes, bank=None) -> UploadedFile:
    """Read the IremboPay payment file (.csv, or .xlsx when the bank allows it); the days it covers come from
    paid_at."""
    is_excel = name.lower().endswith((".xlsx", ".xls")) and "xlsx" in option(bank, "PAYMENT_FILE_TYPES", ["csv"])
    try:
        if is_excel:
            df = pd.read_excel(io.BytesIO(content), dtype={c: str for c in TEXT_COLUMNS}, engine=EXCEL_ENGINE)
        else:
            df = pd.read_csv(io.BytesIO(content), dtype={c: str for c in TEXT_COLUMNS})
    except Exception:
        kind = "an Excel" if is_excel else "a CSV"
        raise ReconError(f"{name}: this file cannot be read as {kind} payment file.") from None
    optional = option(bank, "PAYMENT_OPTIONAL_COLUMNS", [])
    missing = [c for c in PAYMENT_REQUIRED if c not in df.columns and c not in optional]
    if missing:
        raise ReconError(f"{name}: this is not an IremboPay payment file (missing {', '.join(missing)}).")
    for c in DATE_COLUMNS:
        if c in df:
            df[c] = pd.to_datetime(df[c], errors="coerce")
    df["amount"] = pd.to_numeric(df["amount"], errors="coerce")
    df["invoice_number"] = df["invoice_number"].astype(str).str.strip()
    if "created_at" not in df:                      # optional for some banks: paid_at is used instead
        df["created_at"] = df["paid_at"]
    paid = df["paid_at"].dropna()
    if paid.empty:
        raise ReconError(f"{name}: no paid_at dates found in this payment file.")
    providers = frozenset(df["payment_provider"].dropna().astype(str).str.strip().str.upper())
    return UploadedFile(name, "payments", file_digest(content), df, "",
                        paid.min().normalize(), paid.max().normalize(), len(df), providers=providers)


# =====================================================================================================================
# 6. CHECKS BEFORE RUN (step 3 of the app)
#    accept_file: is an uploaded file usable for this period?   check_period: is every day covered?
# =====================================================================================================================
@dataclass
class Check:
    blocking: list = field(default_factory=list)     # Run is blocked
    warnings: list = field(default_factory=list)
    lines: list = field(default_factory=list)        # (ok: bool, text) shown as a checklist


def _ranges(days):
    """[1/9, 2/9, 3/9, 7/9] -> [(1/9, 3/9), (7/9, 7/9)]: missing days shown as ranges."""
    days, out = sorted(days), []
    for d in days:
        if out and (d - out[-1][1]).days == 1:
            out[-1][1] = d
        else:
            out.append([d, d])
    return [(a, b) for a, b in out]


def accept_file(period: Period, f: UploadedFile, accepted: list, bank=None) -> str | None:
    """Reason to reject an uploaded file, or None if it is accepted."""
    bank = get_bank(bank)
    if any(a.digest == f.digest for a in accepted):
        return f"{f.name}: the same file was already added."
    if f.kind == "bank" and f.bank != bank.CODE:
        return f"{f.name}: this is a {bank_name(f.bank)} statement, not {bank.NAME}."
    if f.kind == "payments" and f.providers is not None and not (providers(bank) & f.providers):
        return f"{f.name}: this IremboPay file has no {bank.NAME} payments (payment_provider {bank.PROVIDER})."
    if f.reference:            # read only to spot re-sends (engine.RESEND_LOOKBACK_DAYS back): no window to check
        return None
    # the days a file may cover: statements from the day before (late payments) to the Monday after a weekend
    if f.kind == "bank":
        window_start, window_end = period.day_before, extra_monday(period, bank) or period.last_day
        earlier = option(bank, "ACCEPT_EARLIER_DAYS", 0)      # payments of the period requested before it
        if earlier:
            window_start = period.start - pd.Timedelta(days=earlier)
    else:
        window_start, window_end = period.start, period.last_day
    if f.last < window_start or f.first > window_end:
        return (f"{f.name}: this file covers {f.first:%d/%m/%Y} to {f.last:%d/%m/%Y}, outside "
                f"{period.name} ({period.start:%d/%m/%Y} to {period.last_day:%d/%m/%Y}).")
    return None


def check_period(period: Period, files: list, bank=None, today=None) -> Check:
    """The checklist of step 3: which days are covered by statements, what is missing, what is recommended.
    today: None = the real today (another day only for testing)."""
    bank = get_bank(bank)
    chk = Check()
    today = pd.Timestamp(today if today is not None else pd.Timestamp.today()).normalize()
    statements = [f for f in files if f.kind == "bank" and not f.reference]
    pays = [f for f in files if f.kind == "payments"]
    not_ended = period_not_ended(period, today)
    if not_ended:
        chk.lines.append((False, not_ended.replace("\n", " ")))
        chk.blocking.append(not_ended)
    monday = extra_monday(period, bank)
    needed = list(pd.date_range(period.start, period.last_day))    # every day: only ended periods are reconciled
    if monday is not None:                                         # always needed, even when not over yet
        needed.append(monday)

    # ---- statements: every needed day must be covered (per currency, or once when one file holds all currencies)
    per_currency = option(bank, "ONE_FILE_PER_CURRENCY", True)
    block_missing = option(bank, "MISSING_DAYS_BLOCK", True)
    groups = ([(cur, f"Bank {cur}", f"the {cur} statement", [f for f in statements if f.currency == cur])
               for cur in bank.CURRENCIES] if per_currency
              else [(None, f"{bank.NAME} statements", f"the {bank.NAME} statement", statements)])
    for cur, label, what, mine in groups:
        covered = set()
        for f in mine:
            covered |= set(pd.date_range(f.first, f.last))
        if not mine:
            chk.lines.append((False, f"{label}: no file added → upload {what}(s) for "
                                     f"{period.start:%d/%m} to {(monday or period.last_day):%d/%m/%Y}"))
            chk.blocking.append(f"{label}: no file added.")
        else:
            gaps = [d for d in needed if d not in covered]
            for a, b in _ranges(gaps):
                txt = f"{a:%A %d/%m}" if a == b else f"{a:%d/%m} to {b:%d/%m}"
                why = ""
                if monday is not None and a <= monday <= b:
                    why = " (Monday after the weekend: late weekend payments are posted on Monday"
                    why += (f"; Monday's statement is available from {monday + pd.Timedelta(days=1):%A %d/%m})"
                            if monday >= today else ")")
                chk.lines.append((False, f"{label}: {txt} missing → upload {what} for "
                                         f"{'that day' if a == b else 'those days'}{why}"))
            if gaps and block_missing:
                chk.blocking.append(f"{label}: some days are missing.")
            elif gaps:                                         # some banks: a warning only, Run stays possible
                chk.warnings.append(f"{label}: {sum(1 for _ in gaps)} day(s) have no statement. Run is allowed, but "
                                    "payments of those days will be missing from the report. Upload them if you can.")
            else:
                chk.lines.append((True, f"{label}: {period.start:%d/%m} to {max(needed):%d/%m} covered"))
        # the statement of the day before may hold payments of the first day (recommended, never blocking)
        if cur in bank.DAY_BEFORE_CURRENCIES and period.day_before < today and period.day_before not in covered:
            chk.warnings.append(f"Recommended: upload the {cur} statement for {period.day_before:%d/%m/%Y}. "
                                f"It may contain payments with {bank.COLUMNS['payment_date']} {period.start:%d/%m/%Y}.")

    # ---- reverse check (some banks): IremboPay payments of this bank that are on no statement
    if option(bank, "REVERSE_CHECK", False) and statements and pays:
        _reverse_check(period, bank, statements, pays, chk)

    # ---- the IremboPay payment file
    if not pays:
        chk.lines.append((False, "Payments: no IremboPay payment file added"))
        chk.blocking.append("Payments: no file added.")
    else:
        first = min(f.first for f in pays)
        last = max(f.last for f in pays)
        chk.lines.append((True, f"Payments: {first:%d/%m} to {last:%d/%m} ({sum(f.rows for f in pays):,} payments)"))
        if first > period.start or last < period.last_day:
            chk.warnings.append(f"The payment file looks short: it covers {first:%d/%m} to {last:%d/%m}, the period is "
                                f"{period.start:%d/%m} to {period.last_day:%d/%m}. Invoices outside it will show as not found.")
        taken = [f.taken_at for f in pays if f.taken_at is not None]
        if taken and statements:
            newest_taken, newest_statement = max(taken), max(f.last for f in statements)
            if newest_taken.normalize() < newest_statement:
                chk.warnings.append(f"The IremboPay export was taken {newest_taken:%d/%m/%Y %H:%M}, before the "
                                    f"newest statement ({newest_statement:%d/%m/%Y}): late payments may look like "
                                    "duplicates. Take a fresh export if you can.")

    # ---- currency: a statement row's currency must match IremboPay's currency for the same invoice
    if statements and pays:
        _currency_check(statements, pays, bank, chk)
    return chk


def _currency_check(statements: list, pays: list, bank, chk: Check):
    """A statement row whose invoice is paid in IremboPay with another currency (e.g. a USD file read as RWF)."""
    stmt = bank.clean_statement(pd.concat([f.data for f in statements], ignore_index=True))
    pay = pd.concat([f.data for f in pays], ignore_index=True).drop_duplicates("invoice_number", keep="last")
    known = stmt["invoice_number"].map(pay.set_index("invoice_number")["currency"])
    wrong = stmt[known.notna() & known.ne(stmt["Currency"])]
    if len(wrong):
        by_file = wrong["Source file"].value_counts()
        shown = ", ".join(f"{name} ({n})" for name, n in by_file.items())
        chk.warnings.append(f"{len(wrong)} payment(s) don't match IremboPay's currency for the same invoice: "
                            f"{shown}. Check the file (a USD file may be read as RWF, or the other way round).")


def _reverse_check(period: Period, bank, statements: list, pays: list, chk: Check):
    """IremboPay payments made through this bank in the period that are on no uploaded statement."""
    on_statement = set(bank.clean_statement(pd.concat([f.data for f in statements], ignore_index=True))["invoice_number"])
    p = pd.concat([f.data for f in pays], ignore_index=True)
    p = p[p["payment_provider"].astype(str).str.upper().isin(providers(bank)) & p["payment_status"].isin(bank.COMPLETED_STATUSES)
          & p["paid_at"].between(period.start, period.end, inclusive="left")]
    missing = p[~p["invoice_number"].isin(on_statement)].drop_duplicates("invoice_number")
    if missing.empty:
        chk.lines.append((True, f"IremboPay {bank.NAME} payments: all {len(p):,} found on the statements"))
        return
    days = missing["paid_at"].dt.normalize().value_counts().sort_index()
    shown = ", ".join(f"{d:%d/%m} ({n})" for d, n in list(days.items())[:8]) + (" …" if len(days) > 8 else "")
    chk.lines.append((False, f"IremboPay {bank.NAME} payments: {len(missing):,} of {len(p):,} are on no statement"))
    chk.warnings.append(f"{len(missing):,} IremboPay {bank.NAME} payment(s) paid in the period are on no uploaded "
                        f"statement, paid on: {shown}. Upload the statements of the days they were requested "
                        "(the same day or a few days before).")


# =====================================================================================================================
# 7. REVIEW DECISIONS (step 4 of the app; also used by the developer notebook)
#    Rows not found in IremboPay are reviewed by the user. The decisions are saved in output/.review_store.csv,
#    one line per bank row: bank | period | key | the user's choices. Decisions of other periods are suggested again.
# =====================================================================================================================
def store_path(base_dir: Path) -> Path:
    return Path(base_dir) / "output" / STORE_NAME


def key_of(df, bank=None):
    """Key of a bank row in the review store: currency | payment date | bank reference | invoice."""
    bank = get_bank(bank)
    c = bank.COLUMNS
    if option(bank, "REVIEW_KEY_WITHOUT_REFERENCE", False):          # Ecobank: the reference is not used
        reference = pd.Series("", index=df.index)
    else:
        reference = df[c["reference"]]
        reference = reference.astype(object).where(reference.notna(), "")   # no reference: blank, the row still has
    return (df["Currency"].astype(str) + "|"                                # a key and gets its decision
            + pd.to_datetime(df[c["payment_date"]]).dt.strftime("%Y-%m-%d") + "|"
            + reference.astype(str).str.strip() + "|"
            + df["invoice_number"].fillna("").astype(str).str.strip())


def is_yes(value):
    return str(value).strip().lower() in ("yes", "y")


def load_store(base_dir: Path, bank=None) -> pd.DataFrame:
    """All review decisions (bank=None) or those of one bank. Rows saved before banks existed are I&M."""
    migrate_review_files(base_dir)
    path = store_path(base_dir)
    if not path.exists():
        return pd.DataFrame(columns=STORE_COLS)
    store = pd.read_csv(path, dtype=str).fillna("")
    for c in STORE_COLS:
        if c not in store.columns:
            store[c] = ""
    store.loc[store["bank"].eq(""), "bank"] = "IM"
    # banks whose review ignores the reference: decisions saved with one still match (currency | date | | invoice)
    for code in [k for k, b in BANKS.items() if option(b, "REVIEW_KEY_WITHOUT_REFERENCE", False)]:
        rows = store["bank"].eq(code)
        if rows.any():                                    # currency | date (the first two parts) + the invoice from
            parts = store.loc[rows, "key"].str.split("|", n=2, expand=True)   # its own column: a reference or an
            if parts.shape[1] == 3:                                            # invoice may itself contain "|"
                store.loc[rows, "key"] = (parts[0] + "|" + parts[1] + "||"
                                          + store.loc[rows, "invoice_number"].astype(str).str.strip())
    store = store[STORE_COLS]
    return store if bank is None else store[store["bank"].eq(get_bank(bank).CODE)]


def _write_store(base_dir: Path, store: pd.DataFrame, what="save again"):
    """Save the whole review store (a clear message if the file is open in Excel)."""
    path = store_path(base_dir)
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        store.to_csv(path, index=False)
    except PermissionError:
        raise ReconError(f"The review store is open in another program. Close it and {what}.") from None


def save_decisions(base_dir: Path, period: Period, review: pd.DataFrame, bank=None):
    """Write the decisions of one period (replaces that period's earlier decisions for this bank)."""
    bank = get_bank(bank)
    store = load_store(base_dir)
    rows = review.copy()
    rows["key"] = key_of(rows, bank)
    rows["bank"] = bank.CODE
    rows["period"] = period.store_key
    rows["saved_at"] = pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")
    rows["invoice_number"] = rows["invoice_number"].fillna("").astype(str)
    for c in EDIT_COLS:
        rows[c] = rows[c].fillna("").astype(str).replace({"None": "", "nan": ""})
    rows = rows[rows[EDIT_COLS].apply(lambda r: any(v.strip() for v in r), axis=1)]     # skip rows left empty
    store = store[~(store["bank"].eq(bank.CODE) & store["period"].eq(period.store_key))]
    _write_store(base_dir, pd.concat([store, rows[STORE_COLS]], ignore_index=True))


def migrate_review_files(base_dir: Path):
    """One time: copy the decisions of the old Excel review files into the store, then move those files away."""
    base_dir = Path(base_dir)
    path = store_path(base_dir)
    old_files = sorted((base_dir / "output").glob("*/review/IM_missing_review_*.xlsx"))
    if path.exists() or not old_files:
        return
    rows = []
    for f in old_files:
        try:
            h = pd.read_excel(f, dtype=str).fillna("")
        except Exception:
            continue
        if h.empty:
            continue
        for c in EDIT_COLS:
            if c not in h.columns:
                h[c] = ""
        h = h[h[EDIT_COLS].apply(lambda r: any(str(v).strip() for v in r), axis=1)].copy()
        if h.empty:
            continue
        h["key"] = key_of(h, "IM")
        h["bank"] = "IM"
        h["period"] = f"{f.parent.parent.name}:{f.stem.replace('IM_missing_review_', '')}"
        h["saved_at"] = pd.Timestamp(f.stat().st_mtime, unit="s").strftime("%Y-%m-%d %H:%M:%S")
        rows.append(h[STORE_COLS])
    path.parent.mkdir(parents=True, exist_ok=True)
    (pd.concat(rows, ignore_index=True) if rows else pd.DataFrame(columns=STORE_COLS)).to_csv(path, index=False)
    keep = base_dir / "output" / ".old_review_files"
    for f in old_files:
        target = keep / f.parent.parent.name
        target.mkdir(parents=True, exist_ok=True)
        f.rename(target / f.name)
    for d in {f.parent for f in old_files}:
        try:
            d.rmdir()
        except OSError:
            pass


def _period_title(store_key: str) -> str:
    """'month:01-08-2026_to_01-09-2026' -> 'August 2026 review' (shown in the 'Suggested from' column)."""
    try:
        kind, label = store_key.split(":")
        start = pd.to_datetime(label.split("_to_")[0], format="%d-%m-%Y")
        return f"{start:%B %Y} review" if kind == "month" else f"{kind} review of {start:%d/%m/%Y}"
    except Exception:
        return store_key


def review_rows(base_dir: Path, period: Period, not_found: pd.DataFrame, bank=None):
    """Rows to review for this period: this period's decisions, or suggestions from other periods (same bank).
    Returns (review table, number of suggestions)."""
    bank = get_bank(bank)
    store = load_store(base_dir, bank)
    review = not_found[bank_cols(bank)].copy().reset_index(drop=True)
    review["_key"] = key_of(review, bank)
    for c in EDIT_COLS + [SUGGEST_COL]:
        review[c] = ""
    mine = store[store["period"] == period.store_key].drop_duplicates("key", keep="last").set_index("key")
    others = store[(store["period"] != period.store_key) & store["Reviewed (Yes/No)"].map(is_yes)]
    by_key = others.drop_duplicates("key", keep="last").set_index("key")
    by_inv = (others[others["invoice_number"].ne("")]
              .drop_duplicates("invoice_number", keep="last").set_index("invoice_number"))
    suggested = 0
    for i, r in review.iterrows():
        if r["_key"] in mine.index:                       # 1. decided for this period: use it as it is
            for c in EDIT_COLS:
                review.at[i, c] = mine.loc[r["_key"], c]
            continue
        inv = str(r["invoice_number"]) if pd.notna(r["invoice_number"]) else ""
        source = by_key.loc[r["_key"]] if r["_key"] in by_key.index else (by_inv.loc[inv] if inv in by_inv.index else None)
        if source is not None:                            # 2. decided in another period: a suggestion to confirm
            for c in EDIT_COLS:
                if c != "Reviewed (Yes/No)":
                    review.at[i, c] = source[c]
            review.at[i, SUGGEST_COL] = _period_title(source["period"])
            suggested += 1
    return review, suggested


def apply_review(merged: pd.DataFrame, not_found: pd.DataFrame, payments: pd.DataFrame, review: pd.DataFrame,
                 bank=None):
    """Apply the review decisions to the matched rows (same logic in the app and the notebook).
    Returns (merged, not_found_with_status, counts, messages)."""
    bank = get_bank(bank)
    amount = bank.COLUMNS["amount"]
    merged = merged.copy()
    current = merged["period"].eq("current")
    merged["originally_not_found"] = ~merged["found_in_irembopay"]
    for c in ("review_status", "review_table", "review_comment", "sla_override"):
        merged[c] = ""
    merged["_key"] = key_of(merged, bank)
    pay_by_invoice = payments.set_index("invoice_number")
    payment_fields = [c for c in payments.columns if c != "invoice_number"]
    counts, messages, exclude_keys = {"Not reviewed": 0}, [], []

    for _, r in review.iterrows():
        rows = current & merged["originally_not_found"] & merged["_key"].eq(r["_key"])
        if not is_yes(r["Reviewed (Yes/No)"]):
            merged.loc[rows, "review_status"] = "Not reviewed"
            counts["Not reviewed"] += 1
            continue
        table_choice = str(r["Report table"]).strip()
        if table_choice and table_choice not in TABLE_CHOICES:
            table_choice = ""
        correct_inv = str(r["Correct invoice number"]).strip()
        manual = {"payment_status": str(r["payment_status"]).strip().upper(),
                  "payment_provider": str(r["payment_provider"]).strip(),
                  "paid_at": str(r["paid_at (YYYY-MM-DD HH:MM:SS)"]).strip(),
                  "name": str(r["Merchant name"]).strip()}

        # ---- a corrected invoice number: take its IremboPay record
        if correct_inv:
            if correct_inv in pay_by_invoice.index:
                record = pay_by_invoice.loc[correct_inv]
                for col in payment_fields:
                    merged.loc[rows, col] = record[col]
                merged.loc[rows, "found_in_irembopay"] = True
            else:
                messages.append(("warning", f"Invoice {correct_inv} (typed in the review) is not in the payment file; "
                                            "the typed details are used instead."))
        # ---- details typed by the user (status, provider, paid_at, merchant)
        for col, value in manual.items():
            if value:
                merged.loc[rows, col] = pd.to_datetime(value, errors="coerce") if col == "paid_at" else value
        if any(manual.values()):
            merged.loc[rows, "found_in_irembopay"] = True
            merged.loc[rows, "amount"] = merged.loc[rows, "amount"].fillna(merged.loc[rows, amount])
            merged.loc[rows, "currency"] = merged.loc[rows, "currency"].fillna(merged.loc[rows, "Currency"])

        # ---- the Report table chosen: puts the row in the right summary table
        if table_choice == "Exclude":
            exclude_keys.append(r["_key"])
        elif table_choice == "Notification received":
            merged.loc[rows, "payment_provider"] = bank.PROVIDER
            if manual["payment_status"] not in bank.COMPLETED_STATUSES:
                merged.loc[rows, "payment_status"] = "PAID"
            merged.loc[rows, "found_in_irembopay"] = True
        elif table_choice == "Pending notification":
            merged.loc[rows, "payment_provider"] = bank.PROVIDER
            # NEW, unless the status typed in the review is one the bank counts as pending (Ecobank: EXPIRED too),
            # or one of the two automatic statuses above (manual["payment_status"] is upper-cased; these are not)
            sentinels = {s.upper(): s for s in (NO_INVOICE_STATUS, COULD_NOT_CHECK_STATUS)}
            status = manual["payment_status"]
            if status in option(bank, "PENDING_STATUSES", {"NEW"}):
                merged.loc[rows, "payment_status"] = status
            elif status in sentinels:
                merged.loc[rows, "payment_status"] = sentinels[status]
            else:
                merged.loc[rows, "payment_status"] = "NEW"
            merged.loc[rows, "found_in_irembopay"] = True
        elif table_choice == "Paid with other channel":
            provider = manual["payment_provider"]
            merged.loc[rows, "payment_provider"] = provider if provider and provider != bank.PROVIDER else "OTHER"
            if manual["payment_status"] in ("", "NEW"):
                merged.loc[rows, "payment_status"] = "PAID"
            merged.loc[rows, "found_in_irembopay"] = True
        elif table_choice == "Failed notification":
            for col in payment_fields:
                merged.loc[rows, col] = None
            merged.loc[rows, "found_in_irembopay"] = False
        if table_choice in ("Notification received", "Pending notification", "Paid with other channel"):
            merged.loc[rows, "amount"] = merged.loc[rows, "amount"].fillna(merged.loc[rows, amount])
            merged.loc[rows, "currency"] = merged.loc[rows, "currency"].fillna(merged.loc[rows, "Currency"])

        # ---- what the row became (shown in the report), and an SLA decided by hand
        if table_choice:
            status = table_choice
        elif merged.loc[rows, "found_in_irembopay"].any():
            status = f"Matched to invoice {correct_inv}" if correct_inv in pay_by_invoice.index else "Resolved manually"
        else:
            status = "Failed notification"
        counts[status] = counts.get(status, 0) + 1
        merged.loc[rows, "review_status"] = status
        merged.loc[rows, "review_table"] = table_choice
        merged.loc[rows, "review_comment"] = r["Comment"]
        sla = str(r["SLA violated (Yes/No)"]).strip().lower()
        if sla in ("yes", "y", "no", "n"):
            merged.loc[rows, "sla_override"] = "SLA violated" if sla.startswith("y") else "On time"
    merged["paid_at"] = pd.to_datetime(merged["paid_at"], errors="coerce")

    nf = not_found.copy()
    nf["_key"] = key_of(nf, bank)
    info = merged.drop_duplicates("_key").set_index("_key")[["review_status", "review_comment"]]
    nf = nf.join(info, on="_key").drop(columns="_key")
    if exclude_keys:                                      # "Exclude": the row leaves the report
        merged = merged[~(merged["period"].eq("current") & merged["_key"].isin(exclude_keys))].reset_index(drop=True)
        messages.append(("info", f"{len(exclude_keys)} row(s) excluded from the report, as chosen in the review."))
    return merged, nf, counts, messages


# =====================================================================================================================
# 8. COMPARISON WITH THE PREVIOUS REPORT ("Compare with previous period" in step 1)
#    The previous report of the same bank is read back; each table gets a "VS prev <unit> (%)" column.
# =====================================================================================================================
def read_report_counts(path: Path, sheet: str, bank=None):
    """({table title: {"RWF": n, "USD": n, "Total": n}}, heading, rows not reviewed) from the first block of a
    summary sheet of a report made by this app or the notebook."""
    wb = load_workbook(path, data_only=True)
    try:
        if sheet not in wb.sheetnames:
            raise ReconError(f"{path.name}: this report has no {sheet} sheet.")
        ws = wb[sheet]
        heading = ws.cell(1, 1).value or ""
        counts = {}
        for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 30)):
            for c in row:
                if c.value in TABLE_TITLES and c.value not in counts:
                    vals = {}
                    for i, cur in enumerate(get_bank(bank).CURRENCIES + ["Total"]):
                        v = ws.cell(c.row + 2 + i, c.column + 1).value
                        vals[cur] = int(v) if isinstance(v, (int, float)) else 0
                    counts[c.value] = vals
        not_reviewed = max(wb[NOT_FOUND_SHEET].max_row - 1, 0) if NOT_FOUND_SHEET in wb.sheetnames else 0
    finally:
        wb.close()
    missing = [t for t in TABLE_TITLES if t not in counts]
    if missing:
        raise ReconError(f"{path.name}: this report does not have the expected tables ({', '.join(missing)}).")
    return counts, heading, not_reviewed


def find_previous_report(base_dir: Path, period: Period, bank=None):
    """(path or None, previous Period): the previous report of the same bank."""
    rename_old_reports(base_dir)
    prev = period.previous()
    path = report_path(base_dir, prev, bank)
    return (path if path.exists() else None), prev


def save_uploaded_report(base_dir: Path, period: Period, name: str, content: bytes, bank=None,
                         root: Path | None = None) -> Path:
    """Check an uploaded report for `period` and keep it in the report store (root: see report_path)."""
    tmp = (Path(root) if root is not None else Path(base_dir) / "output") / f".upload_{file_digest(content)}.xlsx"
    tmp.parent.mkdir(parents=True, exist_ok=True)
    tmp.write_bytes(content)
    try:
        wb = load_workbook(tmp, read_only=True)
        sheets = set(wb.sheetnames)
        wb.close()
        mine = get_bank(bank)
        made_here = [kind for kind, sheet in SUMMARY_SHEET.items() if sheet in sheets]
        if not made_here or not any(b.STATEMENT_SHEET in sheets for b in BANKS.values()):
            raise ReconError(f"{name}: this is not a reconciliation report made by this app.")
        if mine.STATEMENT_SHEET not in sheets:                       # 1. the right bank
            other = next(b.NAME for b in BANKS.values() if b.STATEMENT_SHEET in sheets)
            raise ReconError(f"{name}: this report is for {other}, not {mine.NAME}.")
        if SUMMARY_SHEET[period.report_type] not in sheets:          # 2. the right kind (day / week / month)
            kinds = {"day": "a day", "week": "a week", "month": "a month"}
            is_kind = "month" if "month" in made_here else made_here[-1]
            raise ReconError(f"{name}: this is {kinds[is_kind]} report; {period.name} needs {kinds[period.report_type]} "
                             "report.")
        _, heading, _ = read_report_counts(tmp, SUMMARY_SHEET[period.report_type], bank)
        if f"{period.start:%d/%m/%Y}" not in heading:                # 3. the right period
            raise ReconError(f"{name}: this report is for '{heading}', not for {period.name}.")
    except ReconError as err:
        tmp.unlink(missing_ok=True)
        raise ReconError(str(err).replace(tmp.name, name)) from None
    except Exception:
        tmp.unlink(missing_ok=True)
        raise ReconError(f"{name}: this is not a reconciliation report made by this app.") from None
    target = report_path(base_dir, period, bank, root)
    target.parent.mkdir(parents=True, exist_ok=True)
    tmp.replace(target)
    return target


def counts_of(tables: dict) -> dict:
    """{table title: {currency or "Total": number of transactions}} of one block."""
    return {title: {("Total" if pd.isna(c) else c): int(n) for c, n in zip(t["Currency"], t["Number of Transactions"])}
            for title, t in tables.items()}


def merchant_counts(table) -> dict:
    """{(merchant, currency): number of transactions} of a Payment per merchant table ("All" included)."""
    if table is None:
        return {}
    return {(str(m), str(c)): int(n) for m, c, n in zip(table["Merchant name"], table["Currency"],
                                                       table["Number of Transactions"])}


def add_vs_merchants(table: pd.DataFrame, prev, unit: str) -> pd.DataFrame:
    """Add 'VS prev <unit> (%)' to a Payment per merchant table: the change of each merchant (same currency).
    prev None = nothing to compare (empty column); a merchant missing before (or a previous report without
    the table, prev {}) = "N/A"."""
    t = table.copy()
    col = f"VS prev {unit} (%)"
    if prev is None:
        t[col] = None
        return t
    out = []
    for m, c, n in zip(t["Merchant name"], t["Currency"], t["Number of Transactions"]):
        before = prev.get((str(m), str(c)), 0)
        out.append("N/A" if not before else f"{(n - before) / before * 100:.2f}%")
    t[col] = out
    return t


def add_vs_column(tables: dict, prev_counts, unit: str) -> dict:
    """Add 'VS prev <unit> (%)' (change in the number of transactions); empty when there is nothing to compare."""
    col = f"VS prev {unit} (%)"
    out = {}
    for title, t in tables.items():
        t = t.copy()
        if prev_counts is None:
            t[col] = None
        else:
            prev = prev_counts[title]
            t[col] = [("N/A" if not prev["Total" if pd.isna(c) else c]
                       else f"{(n - prev['Total' if pd.isna(c) else c]) / prev['Total' if pd.isna(c) else c] * 100:.2f}%")
                      for c, n in zip(t["Currency"], t["Number of Transactions"])]
        out[title] = t
    return out


# =====================================================================================================================
# 9. THE RECONCILIATION RUN ("Run reconciliation" in step 3)
#    run() goes through steps A to I below. It stops after step D when rows still need a review.
# =====================================================================================================================
@dataclass
class RunResult:
    period: Period
    bank: str = ""
    messages: list = field(default_factory=list)      # (level, text): "info", "warning" or "success"
    bank_rows: int = 0                                # statement rows in the period
    not_found: int = 0                                # rows not found in IremboPay
    review: pd.DataFrame | None = None                # the rows to review (step 4 of the app)
    not_reviewed: int = 0
    stopped_for_review: bool = False                  # True: no report yet, some rows are not reviewed
    blocks: list = field(default_factory=list)        # [(sheet, heading, tables)]
    report_file: Path | None = None
    timings: dict = field(default_factory=dict)       # seconds per step (written in the run log)
    compared_with: str = ""
    ran_at: pd.Timestamp = field(default_factory=pd.Timestamp.now)       # when the run was made
    by_currency: dict = field(default_factory=dict)   # {currency: {"rows", "matched", "not_found", "amount"}}
    review_counts: dict = field(default_factory=dict) # what the reviewed rows became: {"Pending notification": 6, ...}
    sheets: dict = field(default_factory=dict)        # the detail sheets of the report (used by the custom report)
    daily: dict = field(default_factory=dict)         # {day: counts_of(the 5 tables of that day)} (custom charts)
    merchants: dict = field(default_factory=dict)     # {block heading: Payment per merchant table}
    previous_counts: dict | None = None               # "Compare": counts_of() the previous report (Slack message)

    @property
    def reconciled_on(self) -> pd.Timestamp:
        """The day of the run."""
        return self.ran_at.normalize()

    @property
    def matched(self) -> int:
        """Bank rows found in the IremboPay file by invoice number (before the review)."""
        return self.bank_rows - self.not_found

    def add(self, level, text):
        self.messages.append((level, text))


def run(base_dir: Path, period: Period, files: list, sla_days: int = 0, compare: bool = False,
        require_review: bool = True, bank=None, report_dir: Path | None = None,
        previous_report: Path | None = None) -> RunResult:
    """Reconcile the uploaded files for one bank and period, and write the report.
    report_dir: where the Excel is written (default output/; the app gives a temporary folder).
    previous_report: the report compared with (default: looked for in output/).
    Steps: A payments  B statement  C match  D review  E SLA  F summary tables  G totals check  H comparison
    I report + run log."""
    bank = get_bank(bank)
    c = bank.COLUMNS
    posted, paid_on, desc, ref, amount = (c["posted_date"], c["payment_date"], c["description"],
                                          c["reference"], c["amount"])
    base_dir = Path(base_dir)
    output_dir = base_dir / "output"
    start, end = period.start, period.end
    res = RunResult(period, bank.CODE)
    t0 = time.perf_counter()

    not_ended = period_not_ended(period, res.reconciled_on)
    if not_ended:
        raise ReconError(not_ended.replace("\n", " "))
    statements = [f.data for f in files if f.kind == "bank"]
    pay_files = [f for f in files if f.kind == "payments"]
    if not statements:
        raise ReconError("No bank file was added.")
    if not pay_files:
        raise ReconError("No IremboPay payment file was added.")

    # ---- A. payments: one per invoice. Within one file, the successful transaction wins (else the latest one: a
    #         file can list retries of the same invoice). Across several files, the newest file wins (its order) ----
    parts = []
    for f in pay_files:
        d = f.data.copy()
        d["_file_order"] = f.order
        parts.append(d)
    payments = pd.concat(parts, ignore_index=True)
    payments["_successful"] = payments["transaction_status"].astype(str).str.upper().eq("SUCCESSFUL")
    payments = payments.sort_values(["_file_order", "invoice_number", "_successful", "created_at"],
                                    ascending=[True, True, False, False])
    payments = payments.drop_duplicates(["_file_order", "invoice_number"], keep="first")
    payments = (payments.sort_values("_file_order", ascending=True)
                .drop_duplicates("invoice_number", keep="last")
                .drop(columns=["_successful", "_file_order"]).reset_index(drop=True))

    # ---- B. bank statement: all files together, repeated rows dropped (by default the period is set by the
    #         payment date); some banks keep one row per invoice ----
    statement = bank.clean_statement(pd.concat(statements, ignore_index=True))
    statement = statement.drop_duplicates(
        subset=option(bank, "DUPLICATE_KEY", None) or ["Currency"] + present(ref, posted, paid_on, desc, amount)
    ).reset_index(drop=True)
    attempts = statement                                  # every statement row (some banks keep one per invoice)
    if hasattr(bank, "one_row_per_invoice"):
        statement = bank.one_row_per_invoice(statement, payments).reset_index(drop=True)
    res.timings["read"] = time.perf_counter() - t0

    # ---- C. match: every statement row with its IremboPay payment (by invoice number) ----
    t = time.perf_counter()
    merged = statement.merge(payments, on="invoice_number", how="left", indicator=True)
    merged["found_in_irembopay"] = merged["_merge"].eq("both")
    merged = merged.drop(columns="_merge")

    # ---- re-send: an invoice already credited on an earlier day (within RESEND_LOOKBACK_DAYS, read by the
    #      reference files) is not Processed again; neither is one IremboPay already shows paid before this period
    #      (a bank that already keeps one row per invoice, e.g. Ecobank, never has an earlier day here) ----
    first_seen = statement.dropna(subset=["invoice_number"]).groupby("invoice_number")[paid_on].min()
    first_day = merged["invoice_number"].map(first_seen).dt.normalize()
    is_resend = (merged[paid_on].dt.normalize() > first_day) | merged["paid_at"].lt(start).fillna(False)

    if hasattr(bank, "in_period"):                        # the bank's own rule
        in_period = bank.in_period(merged, attempts, start, end)
    else:                                                 # default: payment date inside the period
        in_period = merged[paid_on].between(start, end, inclusive="left")
    in_period = in_period & ~is_resend
    merged["period"] = None
    merged.loc[in_period, "period"] = "current"
    current = merged["period"].eq("current")
    if not current.any():
        raise ReconError(f"The bank files have no Irembo payments with a {paid_on} in {period.name}.")
    no_invoice = merged.loc[current, "invoice_number"].isna().sum()
    if no_invoice:
        res.add("warning", f"{no_invoice} bank row(s) have no invoice number in the description.")
    res.bank_rows = int(current.sum())
    per_cur = merged[current].groupby("Currency").agg(rows=("found_in_irembopay", "size"),
                                                      matched=("found_in_irembopay", "sum"), amount=(amount, "sum"))
    order = [cur for cur in bank.CURRENCIES if cur in per_cur.index] + [cur for cur in per_cur.index
                                                                         if cur not in bank.CURRENCIES]
    res.by_currency = {cur: {"rows": int(per_cur.at[cur, "rows"]), "matched": int(per_cur.at[cur, "matched"]),
                             "not_found": int(per_cur.at[cur, "rows"] - per_cur.at[cur, "matched"]),
                             "amount": float(per_cur.at[cur, "amount"])} for cur in order}
    not_found = merged[current & ~merged["found_in_irembopay"]].copy()
    not_found["Bank channel"] = bank.channel(not_found)
    not_found["Reason"] = ["No invoice number in description" if pd.isna(inv) else "Invoice not in IremboPay payment file"
                           for inv in not_found["invoice_number"]]
    not_found = not_found[bank_cols(bank)].sort_values(paid_on)
    res.not_found = len(not_found)
    res.timings["match"] = time.perf_counter() - t

    # ---- D. review: apply the saved decisions; stop here when some rows are not reviewed yet ----
    t = time.perf_counter()
    review, suggested = review_rows(base_dir, period, not_found, bank)
    if suggested:
        res.add("info", f"{suggested} invoice(s) were already decided in an earlier review: the decision is filled in "
                        "as a suggestion. Check it and set Reviewed = Yes.")
    merged, not_found, counts, msgs = apply_review(merged, not_found, payments, review, bank)
    res.messages += msgs
    res.review = review.drop(columns="_key").reset_index(drop=True)
    res.not_reviewed = counts["Not reviewed"]
    res.review_counts = {k: v for k, v in counts.items() if v and k != "Not reviewed"}
    res.timings["review"] = time.perf_counter() - t
    if require_review and res.not_reviewed:
        res.stopped_for_review = True
        _log_run(output_dir, res, sla_days)
        return res

    # ---- E. SLA (the bank's rule) ----
    t = time.perf_counter()
    merged["paid_date"] = merged["paid_at"].dt.normalize()
    sla_col = option(bank, "SLA_DIFF_COLUMN", "sla_days_difference")
    merged[sla_col], merged["sla_status"] = bank.sla(merged, sla_days)

    # ---- F. summary: flags per row (the bank's calculations) + a reason when a row is not a clean match ----
    rep = (merged[merged["period"] == "current"].sort_values(present(paid_on, posted, ref))
           .reset_index(drop=True))
    has_invoice = rep["invoice_number"].notna()
    rep["times_on_statement"] = rep.groupby("invoice_number")["invoice_number"].transform("size")
    rep["is_extra_credit"] = has_invoice & rep.duplicated("invoice_number", keep="first")
    rep["paid_other_channel"] = bank.duplicate(rep)
    rep["notification_received"] = bank.received(rep)
    # Pending regardless of the bank's own rule: the two automatic statuses above are never Notified or Duplicate
    rep["is_pending"] = bank.pending(rep) | rep["payment_status"].isin({NO_INVOICE_STATUS, COULD_NOT_CHECK_STATUS})
    found_ = rep["found_in_irembopay"]
    provider_ = rep["payment_provider"].astype(object).where(rep["payment_provider"].notna(), "none").astype(str)
    status_ = rep["payment_status"].astype(object).where(rep["payment_status"].notna(), "none").astype(str)
    amount_diff = rep["amount"].notna() & (rep["amount"] - rep[amount]).abs().gt(0.005)
    currency_diff = rep["currency"].notna() & rep["currency"].ne(rep["Currency"])
    parts = [
        ("Other channel (" + provider_ + ")").where(
            found_ & ~rep["payment_provider"].astype(str).str.upper().isin(providers(bank)), ""),
        ("Not settled (" + status_ + ")").where(found_ & ~rep["payment_status"].isin(bank.COMPLETED_STATUSES), ""),
        ("Amount difference (bank " + rep[amount].map("{:,.2f}".format) + " vs IremboPay "
         + rep["amount"].fillna(0).map("{:,.2f}".format) + ")").where(found_ & amount_diff, ""),
        ("Currency difference (bank " + rep["Currency"].astype(str) + " vs IremboPay "
         + rep["currency"].astype(object).where(rep["currency"].notna(), "").astype(str) + ")").where(found_ & currency_diff, ""),
    ]
    combined = parts[0]
    for extra in parts[1:]:
        combined = combined + "; " + extra
    combined = combined.str.replace("(?:; )+", "; ", regex=True).str.strip("; ")
    rep["Reason"] = combined.where(found_, "Not found in IremboPay")

    def table_rows(a, b):
        """Rows of each of the 5 tables for [a, b): the bank's own rule, or by payment date (default). The two
        automatic statuses are always Pending, even for a bank whose own rule would not say so."""
        window = (rep[paid_on] >= a) & (rep[paid_on] < b)
        if hasattr(bank, "tables"):
            m = bank.tables(rep, attempts, a, b)
        else:
            m = {"processed": window, "received": window & rep["notification_received"],
                "sla": window & rep["sla_status"].eq("SLA violated"), "pending": window & rep["is_pending"],
                "duplicate": window & rep["paid_other_channel"]}
        m["pending"] = m["pending"] | (window & rep["payment_status"].isin({NO_INVOICE_STATUS, COULD_NOT_CHECK_STATUS}))
        return m

    def build_summary5(a, b):
        """The 5 summary tables for [a, b): one row per currency + a total row (count only, no amount).
        % = share of the processed payments of the same currency."""
        m = table_rows(a, b)
        df = rep[m["processed"]]
        processed = df.groupby("Currency").size()

        def table(part, with_pct=True):
            rows = []
            for cur in bank.CURRENCIES:
                sub = part[part["Currency"] == cur]
                base = processed.get(cur, 0)
                rows.append({"Currency": cur, "Number of Transactions": len(sub),
                             "Total Amount": sub[amount].sum(), "%": len(sub) / base if base else 0})
            rows.append({"Currency": None, "Number of Transactions": len(part), "Total Amount": None,
                         "%": len(part) / len(df) if len(df) else 0})
            out = pd.DataFrame(rows)
            return out if with_pct else out.drop(columns="%")

        return {
            "Overall processed payments": table(df, with_pct=False),
            "Overall notifications received by Irembo": table(rep[m["received"]]),
            "SLA Violated": table(rep[m["sla"]]),
            "Pending notifications": table(rep[m["pending"]]),
            "Duplicate Payment": table(rep[m["duplicate"]]),
        }

    def build_merchants(a, b):
        """Payment per merchant for [a, b): the notifications received (the same rows as that table), one row per
        merchant and currency - amounts are never added across currencies -, most transactions first, then an
        "All" row per currency. % = share of the currency's notifications received."""
        part = rep[table_rows(a, b)["received"]]
        name = part["name"].astype(object) if "name" in part else pd.Series(None, index=part.index, dtype=object)
        name = name.where(name.notna() & name.astype(str).str.strip().ne(""), "(no merchant)").astype(str).str.strip()
        rows = []
        for cur in list(bank.CURRENCIES) + sorted(set(part["Currency"].dropna()) - set(bank.CURRENCIES)):
            sub = part[part["Currency"] == cur]
            if sub.empty:
                continue
            g = (pd.DataFrame({"m": name[sub.index], "amt": sub[amount]}).groupby("m")
                 .agg(n=("m", "size"), amt=("amt", "sum")).reset_index().sort_values(["n", "m"], ascending=[False, True]))
            total = len(sub)
            rows += [{"Merchant name": m, "Currency": cur, "Number of Transactions": int(n), "Total Amount": amt,
                      "%": n / total} for m, n, amt in zip(g["m"], g["n"], g["amt"])]
            rows.append({"Merchant name": "All", "Currency": cur, "Number of Transactions": total,
                         "Total Amount": sub[amount].sum(), "%": 1.0})
        if not rows:
            rows = [{"Merchant name": "All", "Currency": bank.CURRENCIES[0], "Number of Transactions": 0,
                     "Total Amount": 0.0, "%": 0.0}]
        return pd.DataFrame(rows)

    # one block for a day or a week; for a month one block per week (Monday to Sunday) + one for the month.
    # Headings show the last day of the block (the end date itself is not included).
    one_day = pd.Timedelta(days=1)
    if period.report_type == "day":
        res.blocks = [("Summary-Daily", f"Day: {start:%d/%m/%Y}", build_summary5(start, end))]
        res.merchants[res.blocks[-1][1]] = build_merchants(start, end)
    else:
        if period.report_type == "week":
            weeks = [(start, end)]
        else:
            cuts = [start] + [d for d in pd.date_range(start + pd.Timedelta(days=1), end - pd.Timedelta(days=1))
                              if d.dayofweek == 0]
            weeks = list(zip(cuts, cuts[1:] + [end]))
        for i, (a, b) in enumerate(weeks, 1):
            res.blocks.append(("Summary-Weekly", f"Week {i}: {a:%d/%m/%Y} - {b - one_day:%d/%m/%Y}",
                               build_summary5(a, b)))
            res.merchants[res.blocks[-1][1]] = build_merchants(a, b)
        if period.report_type == "month":
            res.blocks.append(("Summary-Monthly", f"Month: {start:%B %Y} ({start:%d/%m/%Y} - {period.last_day:%d/%m/%Y})",
                               build_summary5(start, end)))
            res.merchants[res.blocks[-1][1]] = build_merchants(start, end)

    # the same counts for each day (only used by the charts of a custom report)
    res.daily = {d: counts_of(build_summary5(d, d + one_day)) for d in pd.date_range(start, period.last_day)}

    # ---- G. totals check: the weeks must add up to the month ----
    if period.report_type == "month":
        month = counts_of(res.blocks[-1][2])
        weeks_sum = {}
        for _, _, tables in res.blocks[:-1]:
            for title, d in counts_of(tables).items():
                for cur, n in d.items():
                    weeks_sum.setdefault(title, {}).setdefault(cur, 0)
                    weeks_sum[title][cur] += n
        skip = option(bank, "TOTALS_CHECK_SKIP", [])     # tables whose weeks may overlap by design
        diffs = [f"{title} {cur}: weeks {weeks_sum[title][cur]:,} vs month {n:,}"
                 for title, d in month.items() for cur, n in d.items()
                 if weeks_sum[title][cur] != n and title not in skip]
        if diffs:
            res.add("warning", "Totals check: the weekly blocks do not add up to the month: " + "; ".join(diffs))
        else:
            res.add("success", "Totals check: the weekly blocks add up to the month.")

    # ---- H. comparison with the previous period of the same bank (only when ticked) ----
    if compare:
        prev_path, prev = find_previous_report(base_dir, period, bank)
        if previous_report is not None:
            prev_path = Path(previous_report)
        if prev_path is None:
            raise ReconError(f"No {bank.NAME} report found for {prev.name}. Upload it or make it first.")
        prev_counts, _, prev_unreviewed = read_report_counts(prev_path, SUMMARY_SHEET[period.report_type], bank)
        res.compared_with = f"{prev.name} (report of {pd.Timestamp(prev_path.stat().st_mtime, unit='s'):%d/%m/%Y})"
        res.previous_counts = prev_counts
        if prev_unreviewed:
            res.add("warning", f"{prev.name} was not reviewed: its {prev_unreviewed} missing invoice(s) are counted as "
                               "not found. The comparison may not be exact.")
        prev_merchants = read_merchant_counts(prev_path, SUMMARY_SHEET[period.report_type])   # {} when it has none
        week_blocks = [b for b in res.blocks if b[0] == "Summary-Weekly"]
        new_blocks = []
        for block in res.blocks:
            sheet_name, heading, tables = block
            if sheet_name == "Summary-Weekly" and period.report_type == "month":
                i = week_blocks.index(block)
                prev_tables = counts_of(week_blocks[i - 1][2]) if i > 0 else None      # week 1: no comparison
                new_blocks.append((sheet_name, heading, add_vs_column(tables, prev_tables, "Week")))
                prev_m = merchant_counts(res.merchants.get(week_blocks[i - 1][1])) if i > 0 else None
                res.merchants[heading] = add_vs_merchants(res.merchants[heading], prev_m, "Week")
            else:
                new_blocks.append((sheet_name, heading, add_vs_column(tables, prev_counts, period.unit)))
                res.merchants[heading] = add_vs_merchants(res.merchants[heading], prev_merchants, period.unit)
        res.blocks = new_blocks
    res.timings["summary"] = time.perf_counter() - t

    # ---- I. report: the detail sheets, the Excel file, the run log ----
    t = time.perf_counter()
    statement_cols = ["Currency"] + present(posted, paid_on, desc, ref, amount)
    detail_cols = statement_cols + option(bank, "DETAIL_COLUMNS", []) + [
        "invoice_number", "found_in_irembopay", "amount", "currency", "payment_status",
        "payment_provider", "payment_channel", "transaction_status", "name", "created_at",
        "paid_at", "settled_at", "provider_reference", "transaction_reference",
        sla_col, "sla_status", "review_status", "review_comment",
        "notification_received", "times_on_statement", "is_extra_credit", "Reason", "Source file"]
    detail = rep[[col for col in detail_cols if col in rep.columns]]
    whole = table_rows(start, end)
    sheets = {"Pending notification": detail[whole["pending"]].sort_values(paid_on),
              "Duplicate payment other channel": detail[whole["duplicate"]].sort_values(["payment_provider", paid_on]),
              bank.STATEMENT_SHEET: rep[statement_cols + ["invoice_number", "Source file"]]}
    if hasattr(bank, "extra_sheets"):
        sheets.update(bank.extra_sheets(rep, attempts, start, end, whole))
    sheets["Payment details"] = detail
    res.sheets = sheets
    report_file = report_path(base_dir, period, bank, report_dir)
    report_file.parent.mkdir(parents=True, exist_ok=True)
    _write_report(report_file, res.blocks, sheets, bank, res.reconciled_on, res.merchants)
    res.report_file = report_file
    res.timings["report"] = time.perf_counter() - t
    res.add("success", f"Report made: {report_file.name}" if report_dir is not None
            else f"Report saved: output/{period.report_type}/{report_file.name}")
    _log_run(output_dir, res, sla_days)
    return res


# =====================================================================================================================
# 10. RUN LOG (output/run_log.csv: one line per run, shown in "Recent runs" in the side panel)
# =====================================================================================================================
def _log_run(output_dir: Path, res: RunResult, sla_days: int):
    """Add one line for this run: when, bank, period, results, and the seconds spent per step."""
    main = next((b for b in res.blocks if b[0] in ("Summary-Monthly", "Summary-Daily")),
                res.blocks[0] if res.blocks else None)
    totals = counts_of(main[2]) if main else {}

    def total(title):
        return totals.get(title, {}).get("Total", "")

    row = {"run_at": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S"),
           "reconciled_on": res.reconciled_on.strftime("%Y-%m-%d"), "bank": res.bank,
           "report_type": res.period.report_type, "period": res.period.text,
           "bank_rows": res.bank_rows, "not_found": res.not_found, "not_reviewed": res.not_reviewed,
           "result": "waiting for review" if res.stopped_for_review else "report made",
           "processed": total("Overall processed payments"),
           "received": total("Overall notifications received by Irembo"),
           "sla_violated": total("SLA Violated"), "pending": total("Pending notifications"),
           "duplicate": total("Duplicate Payment"), "sla_days": sla_days,
           "compared_with": res.compared_with,
           "report_file": res.report_file.name if res.report_file else ""}
    for step in ("read", "match", "review", "summary", "report"):
        row[f"seconds_{step}"] = round(res.timings.get(step, 0.0), 2)
    log = output_dir / "run_log.csv"
    output_dir.mkdir(parents=True, exist_ok=True)
    new = pd.DataFrame([row]).astype(str)
    if log.exists():
        old = pd.read_csv(log, dtype=str)
        if "bank" not in old.columns:                 # runs logged before banks existed are I&M
            old["bank"] = "IM"
        new = pd.concat([old, new], ignore_index=True)
        new = new[list(row) + [c for c in new.columns if c not in row]]
    new.to_csv(log, index=False)


def recent_runs(base_dir: Path, n: int = 8, bank=None) -> pd.DataFrame:
    """The last n runs (of one bank when bank is given), newest first."""
    log = Path(base_dir) / "output" / "run_log.csv"
    if not log.exists():
        return pd.DataFrame()
    try:
        runs = pd.read_csv(log, dtype=str).fillna("")
        if "bank" not in runs.columns:
            runs["bank"] = "IM"
        runs.loc[runs["bank"].eq(""), "bank"] = "IM"
        if bank is not None:
            runs = runs[runs["bank"].eq(get_bank(bank).CODE)]
        return runs.tail(n).iloc[::-1].reset_index(drop=True)
    except Exception:
        return pd.DataFrame()


# =====================================================================================================================
# 11. WRITING THE EXCEL REPORT
#    Summary sheets: each block = a heading + the 5 tables placed like the Google Sheet layout (see GS_SLOTS).
#    Detail sheets: plain tables with a bold header row, filters and number formats.
# =====================================================================================================================
HEADER_FILL = PatternFill("solid", start_color="D9E1F2")
THIN = Side(style="thin", color="000000")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
BOLD = Font(bold=True)
GS_TITLE_FILL = PatternFill("solid", start_color="D0DEE3")
# where each table goes in a block: (column, slot) -> column 1 = A or 6 = F; slot 0/1/2 = rows 3, 11, 19 of the block
GS_SLOTS = {"Overall processed payments": (1, 0), "Overall notifications received by Irembo": (6, 0),
            "SLA Violated": (1, 1), "Pending notifications": (6, 1), "Duplicate Payment": (1, 2)}
MERCHANT_SLOT = (6, 2)                       # Payment per merchant: right of "Duplicate Payment"
GS_STEP = 8                                  # rows between two slots
GS_BLOCK = 2 + 3 * GS_STEP                   # rows taken by one block (heading + 3 slots)
GS_WIDTHS = {"A": 40, "B": 26, "C": 17, "D": 12, "E": 18, "F": 40, "G": 26, "H": 17, "I": 14, "J": 18, "K": 18}
RECONCILED_CELL = (1, 6)                     # F1 of each summary sheet: "Reconciled on: Thursday 08/10/2026"
RECONCILED_LABEL = "Reconciled on: "


def _style_table_sheet(ws, bank):
    """Format a detail sheet: bold header, frozen first row, filters, column widths, number and date formats."""
    c = get_bank(bank).COLUMNS
    c_amount, c_posted, c_paid = c["amount"], c["posted_date"], c["payment_date"]
    with_time = option(bank, "DATETIME_COLUMNS", [])
    for cell in ws[1]:
        cell.font, cell.fill, cell.border = BOLD, HEADER_FILL, BOX
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for col_cells in ws.columns:
        header = col_cells[0].value
        width = max(len(str(c.value)) if c.value is not None else 0 for c in col_cells[:200])
        ws.column_dimensions[get_column_letter(col_cells[0].column)].width = min(max(width + 2, 10), 60)
        for cell in col_cells[1:]:
            if header in (c_amount, "amount"):
                cell.number_format = "#,##0.00"
            elif header in with_time:
                cell.number_format = "yyyy-mm-dd hh:mm:ss"
            elif header in (c_posted, c_paid):
                cell.number_format = "yyyy-mm-dd"
            elif header in DATE_COLUMNS:
                cell.number_format = "yyyy-mm-dd hh:mm:ss"


def _write_table(ws, top, left, title, table):
    """One summary table: a merged title row, the column names, then one row per currency + the total row."""
    ncols = len(table.columns)
    ws.merge_cells(start_row=top, start_column=left, end_row=top, end_column=left + ncols - 1)
    t = ws.cell(top, left, title)
    t.font, t.alignment = Font(size=14), Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[top].height = 22
    for j, name in enumerate(table.columns):
        ws.cell(top, left + j).fill = GS_TITLE_FILL
        ws.cell(top, left + j).border = BOX
        ws.cell(top + 1, left + j, name).border = BOX
    for i, rec in enumerate(table.itertuples(index=False), start=2):
        for j, value in enumerate(rec):
            c = ws.cell(top + i, left + j, None if pd.isna(value) else value)
            c.border = BOX
            name = table.columns[j]
            if name in ("Number of Transactions", "Total Amount"):
                c.number_format = "#,##0"
            elif name == "%":
                c.number_format = "0.00%"
            elif name.startswith("VS prev"):
                c.alignment = Alignment(horizontal="right")


def _write_summary(ws, blocks, reconciled_on=None, merchants=None):
    """A summary sheet: the day it was reconciled (F1), then the blocks one under the other (heading, the 5 tables,
    and Payment per merchant in the 6th place). A block is GS_BLOCK rows, more when its merchant table is long."""
    for col, w in GS_WIDTHS.items():
        ws.column_dimensions[col].width = w
    if reconciled_on is not None:
        ws.cell(*RECONCILED_CELL, f"{RECONCILED_LABEL}{reconciled_on:%A %d/%m/%Y}").font = Font(size=12, italic=True)
    base = 1
    for heading, tables in blocks:
        ws.cell(base, 1, heading).font = Font(size=14, bold=True)
        for title, table in tables.items():
            col, slot = GS_SLOTS[title]
            _write_table(ws, base + 2 + slot * GS_STEP, col, title, table)
        height = GS_BLOCK
        table = (merchants or {}).get(heading)
        if table is not None:
            col, slot = MERCHANT_SLOT
            top = base + 2 + slot * GS_STEP
            _write_table(ws, top, col, MERCHANT_TITLE, table)
            height = max(GS_BLOCK, top - base + 2 + len(table) + 3)       # a long table pushes the next block down
        base += height


def _open_report(path: Path):
    """Excel writer: a new file, or the existing one kept when the user added sheets of their own."""
    try:
        if path.exists():
            existing = load_workbook(path, read_only=True)
            own = [n for n in existing.sheetnames if n not in REPORT_SHEETS]
            existing.close()
            if own:
                return pd.ExcelWriter(path, engine="openpyxl", mode="a", if_sheet_exists="replace")
        return pd.ExcelWriter(path, engine="openpyxl")
    except PermissionError:
        raise ReconError(f"{path.name} is open in Excel. Close it and run again.") from None


def _write_report(path: Path, blocks, sheets, bank=None, reconciled_on=None, merchants=None):
    """Write the whole report: summary sheets first, then the detail sheets."""
    summary_sheets = []
    for sheet, _, _ in blocks:
        if sheet not in summary_sheets:
            summary_sheets.append(sheet)
    try:
        with _open_report(path) as writer:
            for sheet in summary_sheets:                       # empty first, so the summaries are the first tabs
                pd.DataFrame().to_excel(writer, sheet_name=sheet, index=False)
            for name, df in sheets.items():
                df.to_excel(writer, sheet_name=name, index=False)
                _style_table_sheet(writer.sheets[name], bank)
            for sheet in summary_sheets:
                _write_summary(writer.sheets[sheet], [(h, t) for s, h, t in blocks if s == sheet], reconciled_on,
                               merchants)
            for old in {NOT_FOUND_SHEET} | OLD_SHEETS:             # left over in a report made before
                if old in writer.book.sheetnames and old not in sheets:
                    del writer.book[old]
    except PermissionError:
        raise ReconError(f"{path.name} is open in Excel. Close it and run again.") from None


NO_TX_STATUS = "No Transactions Recorded"
NO_TX_NOTE = ("No transactions were recorded for the selected reporting period. Consequently, no payment matching or "
              "reconciliation was performed for this period.")
STATUS_FILL, STATUS_INK = PatternFill("solid", start_color="FFF3D6"), "7A4E00"      # the app's amber notice


def previous_column(period: Period) -> str:
    """The name of the column with the previous period's counts: 'Previous week' / 'Previous month' / 'Previous day'."""
    return f"Previous {period.unit.lower()}"


def read_previous_counts(path: Path, period: Period, bank=None) -> dict:
    """{table title: {currency: n, "Total": n}} of the previous period's report, for the "Previous week" column:
    a standard report (its first summary block), or a No-Transaction Report (its one table)."""
    bank = get_bank(bank)
    sheet = SUMMARY_SHEET[period.report_type]
    wb = load_workbook(path, data_only=True)
    try:
        ws = wb[sheet] if sheet in wb.sheetnames else None
        no_tx = ws is not None and any(ws.cell(r, 1).value == "Reconciliation status" for r in range(1, ws.max_row + 1))
        if no_tx:
            counts = {title: {cur: 0 for cur in bank.CURRENCIES} for title in GS_SLOTS}
            for r in range(1, ws.max_row + 1):
                title, cur, n = ws.cell(r, 1).value, ws.cell(r, 2).value, ws.cell(r, 3).value
                if title in counts and cur in counts[title] and isinstance(n, (int, float)):
                    counts[title][cur] = int(n)
            for d in counts.values():
                d["Total"] = sum(d.values())
            return counts
    finally:
        wb.close()
    return read_report_counts(path, sheet, bank)[0]


def no_transaction_rows(bank=None, previous: dict | None = None) -> list:
    """The rows of a no-transaction Reconciliation Summary (Excel and the app's preview): the 5 categories in the
    standard report's order, at 0 for each currency of the bank; then the previous period's number of transactions
    (read_previous_counts) and the change against it, as the standard reports' 'VS prev' column: "-100.00%";
    "No change" when the previous period had none too; "N/A" when its report was not found."""
    rows = []
    for title in GS_SLOTS:
        for cur in get_bank(bank).CURRENCIES:
            before = None if previous is None else int((previous.get(title) or {}).get(cur, 0))
            change = ("N/A" if before is None else "No change" if before == 0
                      else f"{(0 - before) / before * 100:.2f}%")
            rows.append([title, cur, 0, 0, "N/A" if before is None else before, change])
    return rows


def no_transaction_excel(period: Period, bank=None, made_at=None, previous: dict | None = None) -> bytes:
    """The report of a period in which the bank had no transactions, in the standard report's look (same title
    fill, borders and number formats): one Reconciliation Summary table (every category at 0 for each currency of
    the bank; the last column = the previous period's transactions)
    and the Reconciliation status under it."""
    from openpyxl import Workbook
    bank = get_bank(bank)
    made_at = pd.Timestamp.now() if made_at is None else pd.Timestamp(made_at)
    wb = Workbook()
    ws = wb.active
    ws.title = SUMMARY_SHEET[period.report_type]
    for col, w in {"A": 40, "B": 14, "C": 24, "D": 18, "E": 18, "F": 20}.items():
        ws.column_dimensions[col].width = w
    ws.cell(1, 1, f"{bank.FULL_NAME} – {period.name[0].upper() + period.name[1:]}").font = Font(size=14, bold=True)
    ws.cell(2, 1, f"Reporting period: {period.start:%d/%m/%Y} to {period.last_day:%d/%m/%Y}").font = Font(italic=True)
    ws.cell(2, 3, f"Generated on: {made_at:%A %d/%m/%Y}").font = Font(italic=True)

    # ---- the summary table: title row, column names, one row per category and currency
    cols = ["Category", "Currency", "Number of Transactions", "Total Amount", previous_column(period),
            f"VS prev {period.unit} (%)"]
    top = 4
    ws.merge_cells(start_row=top, start_column=1, end_row=top, end_column=len(cols))
    t = ws.cell(top, 1, "Reconciliation Summary")
    t.font, t.alignment = Font(size=14), Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[top].height = 22
    for j, name in enumerate(cols, 1):
        ws.cell(top, j).fill, ws.cell(top, j).border = GS_TITLE_FILL, BOX
        cell = ws.cell(top + 1, j, name)
        cell.font, cell.border = BOLD, BOX
    r = top + 2
    for values in no_transaction_rows(bank, previous):
        for j, v in enumerate(values, 1):
            cell = ws.cell(r, j, v)
            cell.border = BOX
            if j in (3, 4) or (j == 5 and isinstance(v, int)):
                cell.number_format = "#,##0"
            elif j in (5, 6):
                cell.alignment = Alignment(horizontal="right")
        r += 1

    # ---- the reconciliation status, under the table
    r += 1
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=len(cols))
    ws.cell(r, 1, "Reconciliation status").font = BOLD
    for j in range(1, len(cols) + 1):
        ws.cell(r, j).fill, ws.cell(r, j).border = GS_TITLE_FILL, BOX
    for i, (text, font) in enumerate([(NO_TX_STATUS, Font(bold=True, size=12, color=STATUS_INK)),
                                      (NO_TX_NOTE, Font(color=STATUS_INK))], 1):
        ws.merge_cells(start_row=r + i, start_column=1, end_row=r + i, end_column=len(cols))
        cell = ws.cell(r + i, 1, text)
        cell.font, cell.alignment = font, Alignment(wrap_text=True, vertical="center", indent=1)
        for j in range(1, len(cols) + 1):
            ws.cell(r + i, j).fill, ws.cell(r + i, j).border = STATUS_FILL, BOX
    ws.row_dimensions[r + 1].height, ws.row_dimensions[r + 2].height = 22, 34
    ws.page_setup.orientation, ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = "portrait", 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# =====================================================================================================================
# 12. SAVED REPORTS AND DECISIONS (the side panel: list, view, delete)
# =====================================================================================================================
def rename_old_reports(base_dir: Path):
    """One time: reports named up to the day AFTER the period ('..._01-09-2026_to_01-10-2026.xlsx') are renamed to
    end on the last day ('..._01-09-2026_to_30-09-2026.xlsx'). A name already taken is left as it is."""
    for kind in SUMMARY_SHEET:
        for p in (Path(base_dir) / "output" / kind).glob(f"*{REPORT_SUFFIX}*_to_*.xlsx"):
            try:
                a, b = (pd.to_datetime(x, format="%d-%m-%Y") for x in report_label(p).split("_to_"))
            except ValueError:
                continue
            period = make_period(kind, a)
            if p.name.startswith("~$") or period.start != a or b != period.end:     # not an old name
                continue
            target = p.with_name(f"{report_bank(p)}{REPORT_SUFFIX}{period.label}.xlsx")
            if not target.exists():
                try:
                    p.rename(target)
                except OSError:
                    pass


def saved_reports(base_dir: Path) -> list:
    """[(kind, path)] of every bank (old report names are renamed first)."""
    rename_old_reports(base_dir)
    out = []
    for kind in ("month", "week", "day"):
        for p in sorted((Path(base_dir) / "output" / kind).glob(f"*{REPORT_SUFFIX}*.xlsx"), reverse=True):
            if not p.name.startswith("~$"):                    # "~$..." = Excel's lock file of an open report
                out.append((kind, p))
    return out


def period_title(kind: str, label: str) -> str:
    """'month', '01-09-2026_to_30-09-2026' -> 'Month · September 2026'. Only the first day is read, so the labels
    that end on the day after the period (review decisions, old reports and run log lines) give the same title."""
    try:
        a = pd.to_datetime(label.split("_to_")[0], format="%d-%m-%Y")
        last = make_period(kind, a).last_day
        if kind == "month":
            return f"Month · {a:%B %Y}"
        if kind == "week":
            return f"Week · {a:%d/%m} – {last:%d/%m/%Y}"
        return f"Day · {a:%d/%m/%Y}"
    except Exception:
        return f"{kind} · {label}"


def report_title(path: Path, with_bank=True) -> str:
    """'I&M · Month · September 2026' (or 'Month · September 2026')."""
    path = Path(path)
    kind = path.parent.name if path.parent.name in SUMMARY_SHEET else "month"
    title = period_title(kind, report_label(path))
    return f"{bank_name(report_bank(path))} · {title}" if with_bank else title


def report_list(base_dir: Path, bank=None) -> list:
    """[(title, path)] newest period first: every bank, or only one bank (then the title has no bank name)."""
    code = None if bank is None else get_bank(bank).CODE
    items = []
    for kind, path in saved_reports(base_dir):
        if code is not None and report_bank(path) != code:
            continue
        try:
            start = pd.to_datetime(report_label(path).split("_to_")[0], format="%d-%m-%Y")
        except Exception:
            start = pd.Timestamp(0)
        items.append((start, kind, report_title(path, with_bank=code is None), path))
    items.sort(key=lambda x: (x[0], x[1], x[2]), reverse=True)
    return [(title, path) for _, _, title, path in items]


def delete_report(base_dir: Path, path: Path):
    """Delete one report file (only files inside output/ can be deleted)."""
    path = Path(path).resolve()
    output = (Path(base_dir) / "output").resolve()
    if output not in path.parents or path.suffix != ".xlsx":
        raise ReconError("Only reports made by this app can be deleted here.")
    try:
        path.unlink(missing_ok=True)
    except PermissionError:
        raise ReconError(f"{path.name} is open in Excel. Close it and try again.") from None


def review_list(base_dir: Path, bank=None) -> list:
    """[(title, bank code, store_key, number of decisions, last saved)] newest period first: every bank, or only
    one bank (then the title has no bank name)."""
    store = load_store(base_dir, bank)
    out = []
    for (code, key), grp in store.groupby(["bank", "period"]):
        kind, _, label = key.partition(":")
        try:
            start = pd.to_datetime(label.split("_to_")[0], format="%d-%m-%Y")
        except Exception:
            start = pd.Timestamp(0)
        title = period_title(kind, label) if bank is not None else f"{bank_name(code)} · {period_title(kind, label)}"
        out.append((start, title, code, key, len(grp),
                    grp["saved_at"].max()))
    out.sort(key=lambda x: (x[0], x[1]), reverse=True)
    return [x[1:] for x in out]


def delete_review(base_dir: Path, store_key: str, bank=None):
    """Delete the decisions of one period of one bank (those invoices will need a new review)."""
    code = get_bank(bank).CODE
    store = load_store(base_dir)
    _write_store(base_dir, store[~(store["bank"].eq(code) & store["period"].eq(store_key))], "try again")


def _read_merchant_table(ws, top: int, col: int) -> pd.DataFrame:
    """The Payment per merchant table whose title is at (top, col): its header row, then rows down to a blank."""
    headers, c = [], col
    while ws.cell(top + 1, c).value:
        headers.append(ws.cell(top + 1, c).value)
        c += 1
    rows, r = [], top + 2
    while ws.cell(r, col).value not in (None, ""):
        rows.append([ws.cell(r, col + j).value for j in range(len(headers))])
        r += 1
    t = pd.DataFrame(rows, columns=headers)
    if len(t):
        t["Number of Transactions"] = pd.to_numeric(t["Number of Transactions"], errors="coerce").fillna(0).astype(int)
        t["Total Amount"] = pd.to_numeric(t["Total Amount"], errors="coerce")
        t["%"] = pd.to_numeric(t["%"], errors="coerce").fillna(0.0)
    return t


def read_merchant_counts(path: Path, sheet: str) -> dict:
    """{(merchant, currency): transactions} of the first Payment per merchant table of a summary sheet; {} when
    the report has none (made before 12/10/2026): the comparison then says N/A."""
    wb = load_workbook(path, data_only=True)
    try:
        if sheet not in wb.sheetnames:
            return {}
        ws = wb[sheet]
        for r in range(1, ws.max_row + 1):
            col, _ = MERCHANT_SLOT
            if ws.cell(r, col).value == MERCHANT_TITLE:
                return merchant_counts(_read_merchant_table(ws, r, col))
        return {}
    finally:
        wb.close()


def read_report(path: Path) -> dict:
    """Read a saved report back: {"report_type", "bank", "title", "blocks": [(sheet, heading, tables)], "not_found",
    "reconciled_on" ('Thursday 08/10/2026', or "" for reports made before it was recorded),
    "merchants": {heading: Payment per merchant table} (empty for reports made before 12/10/2026)}.
    Works for reports made by this app (same layout as _write_summary)."""
    path = Path(path)
    kind = path.parent.name if path.parent.name in SUMMARY_SHEET else "month"
    code = report_bank(path)
    currencies = BANKS[code].CURRENCIES if code in BANKS else []
    wb = load_workbook(path, data_only=True)
    try:
        blocks, reconciled_on, merchants = [], "", {}
        for sheet in ("Summary-Daily", "Summary-Weekly", "Summary-Monthly"):
            if sheet not in wb.sheetnames:
                continue
            ws = wb[sheet]
            stamp = str(ws.cell(*RECONCILED_CELL).value or "")
            if stamp.startswith(RECONCILED_LABEL):
                reconciled_on = stamp[len(RECONCILED_LABEL):]
            # a block = a heading in column A with "Overall processed payments" two rows below (found by looking:
            # a block is taller when its merchant table is long; reports made before have the same layout)
            starts = [r for r in range(1, ws.max_row + 1)
                      if ws.cell(r, 1).value and ws.cell(r + 2, 1).value == "Overall processed payments"]
            for base in starts:
                heading = ws.cell(base, 1).value
                tables = {}
                for title, (col, slot) in GS_SLOTS.items():
                    top = base + 2 + slot * GS_STEP
                    if ws.cell(top, col).value != title:
                        continue
                    headers = []
                    c = col
                    while (c < col + 5 and ws.cell(top + 1, c).value       # a table has 5 columns at most: stop at
                           and (c == col or ws.cell(top + 1, c).value != "Currency")):   # the next table's columns
                        headers.append(ws.cell(top + 1, c).value)
                        c += 1
                    n_rows = len(currencies) + 1
                    if not currencies:                                    # unknown bank: rows down to "Total"
                        while ws.cell(top + 1 + n_rows, col).value not in (None, "") and n_rows < 20:
                            n_rows += 1
                    rows = [[ws.cell(top + 2 + i, col + j).value for j in range(len(headers))]
                            for i in range(n_rows)]
                    t = pd.DataFrame(rows, columns=headers)
                    t["Number of Transactions"] = pd.to_numeric(t["Number of Transactions"], errors="coerce").fillna(0).astype(int)
                    t["Total Amount"] = pd.to_numeric(t["Total Amount"], errors="coerce")
                    if "%" in t:
                        t["%"] = pd.to_numeric(t["%"], errors="coerce").fillna(0.0)
                    tables[title] = t
                if len(tables) == len(GS_SLOTS):
                    blocks.append((sheet, heading, tables))
                    col, slot = MERCHANT_SLOT
                    if ws.cell(base + 2 + slot * GS_STEP, col).value == MERCHANT_TITLE:
                        merchants[heading] = _read_merchant_table(ws, base + 2 + slot * GS_STEP, col)
        not_found = 0                                             # rows that went through the review
        if "Payment details" in wb.sheetnames:
            ws = wb["Payment details"]
            headers = [c.value for c in ws[1]]
            if "review_status" in headers:
                col = headers.index("review_status")
                not_found = sum(1 for row in ws.iter_rows(min_row=2, values_only=True) if row[col])
    finally:
        wb.close()
    if not blocks:
        raise ReconError(f"{path.name}: this report cannot be shown (it was not made by this app).")
    return {"report_type": kind, "bank": bank_name(code), "full_name": BANKS[code].FULL_NAME if code in BANKS else code,
            "title": period_title(kind, report_label(path)), "blocks": blocks, "not_found": not_found,
            "reconciled_on": reconciled_on, "merchants": merchants}


# =====================================================================================================================
# 13. PDF (summary only): page 1 = key figures + the main summary; then one page per week
# =====================================================================================================================
def report_pdf(path: Path) -> bytes:
    """Summary-only PDF of a saved report: key figures + main summary on page 1, then one page per week."""
    from reportlab.graphics.charts.barcharts import VerticalBarChart
    from reportlab.graphics.shapes import Drawing
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.platypus import KeepTogether, LongTable, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
    from xml.sax.saxutils import escape

    rep = read_report(path)
    teal, green = colors.HexColor("#123239"), colors.HexColor("#99E378")
    title_fill, grid = colors.HexColor("#D0DEE3"), colors.HexColor("#9AA8AD")
    warning_fill, warning_stroke = colors.HexColor("#FFF3D6"), colors.HexColor("#D4A72C")
    muted = colors.HexColor("#5B6B70")
    made = pd.Timestamp(Path(path).stat().st_mtime, unit="s")
    page_w, page_h = landscape(A4)
    margin = 14 * mm
    usable = page_w - 2 * margin
    h1 = ParagraphStyle("h1", fontName="Helvetica-Bold", fontSize=17, textColor=green, leading=22)
    sub = ParagraphStyle("sub", fontName="Helvetica", fontSize=9.5, textColor=colors.HexColor("#D6E4E6"), leading=12)
    h2 = ParagraphStyle("h2", fontName="Helvetica-Bold", fontSize=12.5, textColor=teal, leading=16, spaceAfter=4)
    h3 = ParagraphStyle("h3", fontName="Helvetica-Bold", fontSize=9.2, textColor=teal, leading=12, spaceAfter=3)
    small = ParagraphStyle("small", fontName="Helvetica", fontSize=8.1, textColor=muted, leading=10)
    cell = ParagraphStyle("cell", fontName="Helvetica", fontSize=8.2, leading=10)

    def title_band(subtitle, warning=False):
        """Balanced title band with the current report metadata."""
        label = (rep.get("full_name") or "Bank Reconciliation").strip()
        t = Table([[Paragraph(f"{escape(label)} Reconciliation", h1)], [Paragraph(subtitle, sub)]], colWidths=[usable])
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), teal),
            ("LEFTPADDING", (0, 0), (-1, -1), 12),
            ("RIGHTPADDING", (0, 0), (-1, -1), 12),
            ("TOPPADDING", (0, 0), (-1, -1), 8),
            ("BOTTOMPADDING", (0, -1), (-1, -1), 8),
            ("LINEBEFORE", (0, 0), (0, -1), 5, green if not warning else warning_stroke),
            ("TEXTCOLOR", (0, 0), (-1, -1), colors.white),
        ]))
        return t

    def fmt_value(v):
        if v is None or (not isinstance(v, str) and pd.isna(v)):
            return ""
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            if abs(v) >= 1000:
                return f"{v:,.0f}"
            return f"{v:,.2f}" if isinstance(v, float) and not v.is_integer() else f"{v:,.0f}"
        return str(v)

    def one_table(title, t, width, warning=False):
        """One summary table with consistent numeric alignment and a clear total row."""
        cols = list(t.columns)
        head = {"Number of Transactions": "Transactions", "Total Amount": "Total amount", "%": "% of processed"}
        rows = [[title] + [""] * (len(cols) - 1), [head.get(c, c) for c in cols]]
        for rec in t.itertuples(index=False):
            out = []
            for c, v in zip(cols, rec):
                if c == "Currency":
                    out.append("Total" if pd.isna(v) else v)
                elif c == "Number of Transactions":
                    out.append(f"{int(v):,}")
                elif c == "Total Amount":
                    out.append("" if pd.isna(v) else f"{v:,.0f}")
                elif c == "%":
                    out.append(f"{v:.2%}")
                else:
                    out.append("" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v))
            rows.append(out)
        first = 0.24 * width
        widths = [first] + [(width - first) / (len(cols) - 1)] * (len(cols) - 1)
        tb = LongTable(rows, colWidths=widths, repeatRows=2)
        title_bg = warning_fill if warning else title_fill
        title_line = warning_stroke if warning else grid
        style = [
            ("SPAN", (0, 0), (-1, 0)),
            ("BACKGROUND", (0, 0), (-1, 0), title_bg),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, 0), 9.2),
            ("TEXTCOLOR", (0, 0), (-1, 0), teal if not warning else colors.HexColor("#563B00")),
            ("ALIGN", (0, 0), (-1, 0), "CENTER"),
            ("FONTNAME", (0, 1), (-1, 1), "Helvetica-Bold"),
            ("FONTSIZE", (0, 1), (-1, -1), 8.2),
            ("FONTNAME", (-1, 2), (-1, -1), "Helvetica-Bold"),
            ("ALIGN", (1, 2), (-1, -1), "RIGHT"),
            ("GRID", (0, 0), (-1, -1), 0.45, title_line),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ("ROWBACKGROUNDS", (0, 2), (-1, -1), [colors.white, colors.HexColor("#F7FAFB")]),
        ]
        if warning:
            style.extend([
                ("LINEABOVE", (0, 1), (-1, 1), 1.4, warning_stroke),
                ("TEXTCOLOR", (0, 2), (-1, -1), colors.HexColor("#563B00")),
            ])
        tb.setStyle(TableStyle(style))
        return tb

    def summary_grid(tables):
        """The 5 summary tables, arranged in a compact 2-column grid."""
        half = (usable - 8 * mm) / 2
        order = [["Overall processed payments", "Overall notifications received by Irembo"],
                 ["SLA Violated", "Pending notifications"], ["Duplicate Payment", None]]
        rows = [[one_table(a, tables[a], half, warning=(a == "SLA Violated")),
                 one_table(b, tables[b], half, warning=(b == "SLA Violated")) if b else ""]
                for a, b in order]
        g = Table(rows, colWidths=[half + 4 * mm, half + 4 * mm])
        g.setStyle(TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
            ("LEFTPADDING", (0, 0), (-1, -1), 0),
            ("RIGHTPADDING", (0, 0), (-1, -1), 8 * mm),
        ]))
        return g

    def distribution_chart(tables, width, height):
        """Small compact distribution chart for received, pending, duplicate and not-found totals."""
        counts = counts_of(tables)
        labels = ["Received", "Pending", "Duplicate", "Not found"]
        values = [
            counts.get("Overall notifications received by Irembo", {}).get("Total", 0),
            counts.get("Pending notifications", {}).get("Total", 0),
            counts.get("Duplicate Payment", {}).get("Total", 0),
            rep.get("not_found", 0),
        ]
        if sum(values) <= 0:
            return None
        d = Drawing(width, height)
        bc = VerticalBarChart()
        bc.x = 22
        bc.y = 16
        bc.height = max(52, height - 30)
        bc.width = width - 32
        bc.data = [values]
        bc.categoryAxis.categoryNames = labels
        bc.categoryAxis.labels.fontName = "Helvetica"
        bc.categoryAxis.labels.fontSize = 7
        bc.categoryAxis.strokeColor = grid
        bc.categoryAxis.visibleTicks = False
        bc.valueAxis.valueMin = 0
        bc.valueAxis.visibleGrid = True
        bc.valueAxis.gridStrokeColor = colors.HexColor("#E3ECEF")
        bc.valueAxis.strokeColor = grid
        bc.valueAxis.labels.fontName = "Helvetica"
        bc.valueAxis.labels.fontSize = 7
        bc.valueAxis.labelTextFormat = lambda v: f"{int(v):,}"
        bc.bars[0].fillColor = teal
        bc.bars[0].strokeColor = colors.HexColor("#122E36")
        bc.groupSpacing = 18
        d.add(bc)
        return d

    def kpis(tables, not_found):
        """The six key figures, styled as a compact KPI grid."""
        def total(title):
            t = tables[title]
            row = t[t["Currency"].isna()]
            return int(row["Number of Transactions"].iloc[0]), (row["%"].iloc[0] if "%" in t else None)

        processed, _ = total("Overall processed payments")
        _, rec_pct = total("Overall notifications received by Irembo")
        sla, _ = total("SLA Violated")
        values = [
            ("Processed", f"{processed:,}"),
            ("Received", f"{rec_pct:.2%}"),
            ("SLA violated", f"{sla:,}"),
            ("Pending (NEW)", f"{total('Pending notifications')[0]:,}"),
            ("Duplicate", f"{total('Duplicate Payment')[0]:,}"),
            ("Not found", f"{not_found:,}"),
        ]

        cards = []
        for label, value in values:
            cards.append(Table([[Paragraph(escape(label), small)], [Paragraph(escape(value), ParagraphStyle("kpi", fontName="Helvetica-Bold", fontSize=14, textColor=teal, leading=18))]], colWidths=[usable / 6]))
            cards[-1].setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, -1), colors.white),
                ("BOX", (0, 0), (-1, -1), 0.6, grid),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ]))
        row = Table([cards], colWidths=[usable / 6] * 6)
        row.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0)]))
        return row

    main = next((b for b in rep["blocks"] if b[0] in ("Summary-Monthly", "Summary-Daily")), rep["blocks"][0])
    weeks = [b for b in rep["blocks"] if b[0] == "Summary-Weekly" and b is not main]
    when = f"reconciled on {rep['reconciled_on']}" if rep["reconciled_on"] else f"made on {made:%d/%m/%Y %H:%M}"
    report_title = rep.get("title") or "Report"
    subtitle = f"{report_title} · {when}"

    story = [title_band(subtitle)]
    story += [Spacer(1, 5 * mm), kpis(main[2], rep["not_found"])]

    chart = distribution_chart(main[2], usable * 0.52, 70 * mm)
    if chart is not None:
        chart_block = Table([[Paragraph("Transaction distribution", h3), ""], [chart, ""]], colWidths=[usable * 0.52, usable * 0.48])
        chart_block.setStyle(TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ]))
        story += [Spacer(1, 5 * mm), chart_block]

    story += [Spacer(1, 5 * mm), Paragraph(main[1], h2), KeepTogether(summary_grid(main[2]))]

    for _, heading, tables in weeks:
        story += [PageBreak(), title_band(f"{report_title} · {heading.split(':')[0]}", warning=heading.startswith("Week") and "SLA" in str(tables))]
        story += [Spacer(1, 5 * mm), Paragraph(heading, h2), KeepTogether(summary_grid(tables))]

    def footer(canvas, doc):
        """Page number and report identity on every page."""
        canvas.saveState()
        canvas.setFont("Helvetica-Bold", 8)
        canvas.setFillColor(teal)
        canvas.drawString(margin, 8 * mm, f"{report_title} · {rep.get('full_name') or 'Report'}")
        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(muted)
        canvas.drawRightString(page_w - margin, 8 * mm, f"Page {doc.page}")
        canvas.restoreState()

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=landscape(A4), leftMargin=margin, rightMargin=margin,
                            topMargin=margin, bottomMargin=margin, title=f"{rep.get('full_name') or 'Reconciliation'} {report_title}")
    doc.build(story, onFirstPage=footer, onLaterPages=footer)
    return buf.getvalue()


# =====================================================================================================================
# 14. CUSTOM REPORT (step 6, "Custom report": the user picks the sections, how each is shown, and the layout)
#    custom_content() builds the content once from the run; the app's preview, custom_pdf() and custom_excel() all
#    draw that same content, so they always match. The standard report in output/ is never changed.
#    Charts are drawn in the PDF, the preview and the Excel file (real Excel charts, with their numbers beside).
# =====================================================================================================================
TABLE, BAR, LINE, KPI = "Table", "Bar chart", "Line chart", "KPI summary"
CHART_TEAL = "#123239"           # bars and lines of the PDF charts (IremboPay teal: 13.6:1 contrast on white paper)
# key: (title, the summary table it is read from or None, the ways it can be shown: the first is the default)
CUSTOM_SECTIONS = {
    "overview": ("Reconciliation Overview", None, [KPI]),
    "payments": ("Payment Summary", "Overall processed payments", [TABLE, BAR, LINE]),
    "received": ("Notifications Received", "Overall notifications received by Irembo", [TABLE, BAR, LINE]),
    "sla": ("SLA Violations", "SLA Violated", [TABLE, BAR, LINE]),
    "pending": ("Pending Notifications", "Pending notifications", [TABLE, BAR, LINE]),
    "duplicates": ("Duplicate Payments", None, [TABLE]),                       # credited twice on the statement
    "other_channel": ("Other Channel Payments", "Duplicate Payment", [TABLE, BAR, LINE]),   # paid another way
    "merchants": ("Payment per Merchant", None, [TABLE, BAR]),       # notifications received, by merchant
}
# report type: (description, {section: how it is shown}) - choosing a type fills in the sections and views
REPORT_TYPES = {
    "Full Report": ("Detailed reconciliation",
                    {k: v[2][0] for k, v in CUSTOM_SECTIONS.items() if k != "other_channel"}),
    "Management Summary": ("KPIs & trends", {"overview": KPI, "payments": TABLE, "received": BAR, "sla": BAR,
                                             "pending": TABLE}),
    "Pending & Duplicates Report": ("Pending, duplicate & other-channel payments",
                                    {"pending": TABLE, "duplicates": TABLE, "other_channel": TABLE}),
}
CUSTOM_INCLUDE = {"kpis": "Summary KPIs", "currency": "Currency breakdown", "timestamp": "Reconciliation timestamp"}
# once the report is made, the reviewed rows are simply in the table their decision chose: the custom report does not
# list or count them again (the decisions are seen in Exception Resolution only)


def chart_groups(period: Period) -> list:
    """How a chart can group the transactions: by week (a month), by day (a week or a month), by currency."""
    return ((["Week"] if period.report_type == "month" else []) + (["Day"] if period.report_type != "day" else [])
            + ["Currency"])


def _summary_frame(t: pd.DataFrame) -> pd.DataFrame:
    """A summary table for the reader: Currency (Total), Transactions, Total amount, % of processed."""
    out = pd.DataFrame({"Currency": t["Currency"].fillna("Total").values,
                        "Transactions": t["Number of Transactions"].astype(int).values,
                        "Total amount": t["Total Amount"].values})
    if "%" in t:
        out["% of processed"] = t["%"].values
    return out


def _list_frame(df: pd.DataFrame, cols: dict) -> pd.DataFrame:
    """Some columns of a detail sheet, renamed for the reader ({column: name}); missing columns are skipped."""
    keep = {c: n for c, n in cols.items() if c in df.columns}
    return df[list(keep)].rename(columns=keep).reset_index(drop=True)


def _series(res: RunResult, main, weeks, table_title: str, view: dict):
    """([(label, transactions)], group used) for a chart: by week, by day or by currency."""
    col = "Total" if view["currency"] == "All" else view["currency"]
    group = view["group"]
    if group == "Week" and weeks:
        return [(f"Week {i}", counts_of(t)[table_title].get(col, 0)) for i, (_, _, t) in enumerate(weeks, 1)], group
    if group == "Day" and len(res.daily) > 1:
        return [(f"{d:%d/%m}", v[table_title].get(col, 0)) for d, v in res.daily.items()], group
    return [(k, n) for k, n in counts_of(main[2])[table_title].items() if k != "Total"], "Currency"


def custom_content(res: RunResult, opts: dict, bank=None) -> dict:
    """Everything a custom report shows, from the run and the options chosen in the app:
    opts = {"type", "sections": [keys], "views": {key: {"as", "group", "currency"}}, "orientation", "density",
            "include": {"kpis", "currency", "timestamp": True/False}}"""
    bank = get_bank(bank)
    c = bank.COLUMNS
    period = res.period
    paid_on, amount, ref = c["payment_date"], c["amount"], c["reference"]
    main = next((b for b in res.blocks if b[0] in ("Summary-Monthly", "Summary-Daily")), res.blocks[0])
    weeks = [b for b in res.blocks if b[0] == "Summary-Weekly"] if period.report_type == "month" else []
    inc = opts["include"]
    out = {"title": f"{bank.FULL_NAME} Reconciliation Report", "period": period.name[0].upper() + period.name[1:],
           "type": opts["type"], "orientation": opts["orientation"], "density": opts["density"],
           "stamp": f"Reconciled on {res.ran_at:%d %b %Y · %H:%M}" if inc.get("timestamp") else "",
           "kpis": [], "currency_lines": [], "exceptions_line": "", "sections": [],
           "file_stem": f"{bank.CODE}_{opts['type'].lower().replace(' ', '_')}_{period.label}"}
    chosen = [k for k in CUSTOM_SECTIONS if k in opts["sections"]]

    # ---- the overview: key figures and one line per currency
    if "overview" in chosen:
        if inc.get("kpis"):
            rate = res.matched / res.bank_rows if res.bank_rows else 0
            out["kpis"] = [("Processed", f"{res.bank_rows:,}"), ("Matched", f"{res.matched:,}"),
                           ("Match rate", f"{rate:.2%}")]
        if inc.get("currency"):
            out["currency_lines"] = [f"{cur} · {d['rows']:,} rows · {d['matched']:,} matched"
                                     for cur, d in res.by_currency.items()]

    # ---- the other sections: a summary table or a chart, and the list of invoices when there is one
    details = res.sheets.get("Payment details", pd.DataFrame())
    base = {paid_on: paid_on, "Currency": "Currency", "invoice_number": "Invoice", amount: "Amount"}
    for key in chosen:
        if key == "overview":
            continue
        title, table_title, views = CUSTOM_SECTIONS[key]
        view = {"as": views[0], "group": chart_groups(period)[0], "currency": "All", **opts["views"].get(key, {})}
        if view["as"] not in views:
            view["as"] = views[0]
        sec = {"key": key, "title": title, "as": view["as"], "table": None, "series": [], "group": "",
               "currency": view["currency"], "list": None, "list_title": ""}
        if key == "merchants":                            # the Payment per merchant table of the main block
            m = res.merchants.get(main[1])
            if m is not None and len(m):
                if view["as"] == TABLE:
                    sec["table"] = pd.DataFrame({"Merchant": m["Merchant name"].values, "Currency": m["Currency"].values,
                                                 "Transactions": m["Number of Transactions"].astype(int).values,
                                                 "Total amount": m["Total Amount"].values,
                                                 "% of received": [f"{v:.2%}" for v in m["%"]]})
                else:                                     # the 10 merchants with the most transactions
                    part = m[m["Merchant name"].ne("All")]
                    if view["currency"] != "All":
                        part = part[part["Currency"].eq(view["currency"])]
                    top = (part.groupby("Merchant name")["Number of Transactions"].sum()
                           .sort_values(ascending=False).head(10))
                    sec["series"], sec["group"] = [(str(k), int(v)) for k, v in top.items()], "Merchant"
            out["sections"].append(sec)
            continue
        if table_title and view["as"] == TABLE:
            sec["table"] = _summary_frame(main[2][table_title])
        elif table_title:
            sec["series"], sec["group"] = _series(res, main, weeks, table_title, view)
        # the invoice list goes with the "Table" view only: a section shown as a chart shows the chart alone
        if key == "pending" and view["as"] == TABLE:
            sec["list"] = _list_frame(res.sheets.get("Pending notification", pd.DataFrame()),
                                      {**base, "payment_status": "Status", "name": "Merchant"})
            sec["list_title"] = "Pending invoices"
        elif key == "other_channel" and view["as"] == TABLE:
            sec["list"] = _list_frame(res.sheets.get("Duplicate payment other channel", pd.DataFrame()),
                                      {**base, "payment_provider": "Provider", "payment_status": "Status"})
            sec["list_title"] = "Invoices paid with another channel"
        elif key == "duplicates":
            if "Debited more than once" in res.sheets:                       # Ecobank's own sheet
                sec["list"] = _list_frame(res.sheets["Debited more than once"],
                                          {**base, "Cba Ref No": "Bank reference", "Debit": "Debit"})
            elif "times_on_statement" in details:
                sec["list"] = _list_frame(details[details["times_on_statement"] > 1],
                                          {**base, ref: "Bank reference", "times_on_statement": "Times on statement"})
            sec["list_title"] = "Invoices credited more than once"
        out["sections"].append(sec)
    return out


def layout_rows(sections: list) -> list:
    """How the sections are placed (preview and PDF): charts that follow each other go side by side, two per row,
    each in its own card; the other sections take the whole width.
    -> [("charts", [section, section]) or ("charts", [section]) or ("section", section)]"""
    rows, run = [], []
    for sec in sections + [None]:
        if sec is not None and sec["series"]:
            run.append(sec)
            continue
        rows += [("charts", run[i:i + 2]) for i in range(0, len(run), 2)]
        run = []
        if sec is not None:
            rows.append(("section", sec))
    return rows


def series_label(sec: dict) -> str:
    """'Transactions by week' (+ ' · RWF' when one currency is chosen)."""
    return (f"Transactions by {sec['group'].lower()}"
            + (f" · {sec['currency']}" if sec["currency"] != "All" and sec["group"] != "Currency" else ""))


def cell_text(name: str, value) -> str:
    """One value of a custom report table as text (PDF and the app's preview): amounts, %, dates, counts."""
    if value is None or (not isinstance(value, str) and pd.isna(value)):
        return ""
    if name == "% of processed":
        return f"{value:.2%}"
    if isinstance(value, pd.Timestamp):
        return f"{value:%d/%m/%Y %H:%M}" if (value.hour or value.minute) else f"{value:%d/%m/%Y}"
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return f"{value:,.0f}"
    return str(value)


# ---- the custom report as Excel: the same look as the PDF (IremboPay teal band, KPI cards, styled tables) and the
#      charts as real Excel charts, drawn from the numbers written on the section's sheet
XL_TEAL, XL_GREEN, XL_WARN = "123239", "99E378", "D4A72C"
XL_INK, XL_MUTED, XL_RULE, XL_HAIR = "1D2B2F", "5B6B70", "9AA8AD", "E1E8EA"
XL_TITLE_FILL, XL_TOTAL_FILL, XL_WARN_FILL = "E8F0F2", "F2F6F7", "FFF3D6"
NO_AMOUNT_XL = "—"            # the Total row has no amount: currencies are never added together


def _xl_fill(color: str) -> PatternFill:
    return PatternFill("solid", start_color=color, end_color=color)


def _xl_band(ws, title: str, sub: str, last_col: int, fs: float, back: str = ""):
    """The teal title band across the top of a sheet (+ a link back to the Overview)."""
    fill = _xl_fill(XL_TEAL)
    for r in (1, 2, 3):
        for c in range(1, last_col + 1):
            ws.cell(r, c).fill = fill
    ws.row_dimensions[1].height, ws.row_dimensions[2].height, ws.row_dimensions[3].height = 8, 26, 20
    ws.cell(2, 2, title).font = Font(size=fs + 7, bold=True, color="FFFFFF")
    ws.cell(3, 2, sub).font = Font(size=fs, color=XL_GREEN)
    if back:
        cell = ws.cell(2, last_col, "← Overview")
        cell.hyperlink, cell.font = f"#'{back}'!A1", Font(size=fs, color="FFFFFF", underline="single")
        cell.alignment = Alignment(horizontal="right")
    ws.column_dimensions["A"].width = 2


def _xl_page(ws, content: dict):
    """Print setup: the orientation chosen in the app, one page wide, no grid lines on screen."""
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = content["orientation"].lower()
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.oddFooter.center.text = f"{content['title']} · {content['period']} · Page &P of &N"


def _xl_table(ws, top: int, left: int, df: pd.DataFrame, fs: float, title: str = "", total_row: bool = False,
              warning: bool = False, filters: bool = False) -> int:
    """A table like the PDF's: optional title row, tinted header with a teal rule, hairlines between rows, numbers
    on the right, bold tinted Total row (its amount is '—': currencies are never added together).
    Returns the first free row after it."""
    line = XL_WARN if warning else XL_TEAL
    hair, rule = Side(style="thin", color=XL_HAIR), Side(style="medium", color=line)
    cols = list(df.columns)
    numeric = {c for c in cols if pd.api.types.is_numeric_dtype(df[c]) or c == "% of processed"
               or (len(df) and df[c].astype(str).str.endswith("%").all())}
    if title:
        cell = ws.cell(top, left, title)
        cell.font = Font(size=fs + 1, bold=True, color=XL_INK)
        cell.border = Border(left=Side(style="thick", color=line))
        top += 1
    for j, name in enumerate(cols):
        cell = ws.cell(top, left + j, name)
        cell.font = Font(size=fs - 0.5, bold=True, color=XL_MUTED)
        cell.fill = _xl_fill(XL_WARN_FILL if warning else XL_TITLE_FILL)
        cell.border = Border(bottom=rule, top=hair)
        cell.alignment = Alignment(horizontal="right" if name in numeric else "left", vertical="center",
                                   wrap_text=True)
        width = max([len(str(name))] + [len(cell_text(name, v)) for v in df[name].head(300)])
        letter = get_column_letter(left + j)
        ws.column_dimensions[letter].width = max(ws.column_dimensions[letter].width or 0, min(max(width + 3, 11), 45))
    ws.row_dimensions[top].height = 20
    for i, rec in enumerate(df.itertuples(index=False), 1):
        is_total = total_row and str(rec[0]) in ("Total", "All")
        for j, (name, value) in enumerate(zip(cols, rec)):
            empty = value is None or (not isinstance(value, str) and pd.isna(value))
            if empty and is_total and name == "Total amount":
                value, empty = NO_AMOUNT_XL, False
            cell = ws.cell(top + i, left + j, None if empty else (value.to_pydatetime() if isinstance(value, pd.Timestamp)
                                                                else value))
            cell.font = Font(size=fs, bold=is_total, color="7A4E00" if warning else XL_INK)
            cell.border = Border(bottom=hair, top=rule if is_total else None)
            if is_total:
                cell.fill = _xl_fill(XL_TOTAL_FILL)
            cell.alignment = Alignment(horizontal="right" if name in numeric else "left", vertical="center")
            if name == "% of processed":
                cell.number_format = "0.00%"
            elif isinstance(value, pd.Timestamp):
                cell.number_format = "dd/mm/yyyy hh:mm" if (value.hour or value.minute) else "dd/mm/yyyy"
            elif isinstance(value, (int, float)) and not isinstance(value, bool):
                cell.number_format = "#,##0"
    last = top + len(df)
    if filters and len(df):
        ws.auto_filter.ref = f"{get_column_letter(left)}{top}:{get_column_letter(left + len(cols) - 1)}{last}"
    return last + 2



def _xl_chart(sec: dict, ws, head: int, n: int):
    """A one-series Excel bar or line chart of the section's numbers (written on `ws` at rows head..head+n, columns
    B:C), teal like the PDF (amber for SLA), values on the bars when there are few, no legend."""
    from openpyxl.chart import BarChart, LineChart, Reference
    from openpyxl.chart.label import DataLabelList
    color = XL_WARN if sec["key"] == "sla" else XL_TEAL
    ch = BarChart() if sec["as"] == BAR else LineChart()
    if sec["as"] == BAR:
        ch.type, ch.gapWidth = "col", 60
    ch.add_data(Reference(ws, min_col=3, min_row=head, max_row=head + n), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=2, min_row=head + 1, max_row=head + n))
    s = ch.series[0]
    s.graphicalProperties.line.solidFill = color
    if sec["as"] == BAR:
        s.graphicalProperties.solidFill = color
    else:
        s.graphicalProperties.line.width = 28575                          # 2.25 pt
        s.marker.symbol, s.marker.size = "circle", 6
        s.marker.graphicalProperties.solidFill = color
        s.marker.graphicalProperties.line.solidFill = "FFFFFF"
        s.smooth = False
    if n <= 12:
        s.dLbls = DataLabelList()
        s.dLbls.showVal = True
        for flag in ("showSerName", "showCatName", "showLegendKey", "showPercent"):
            setattr(s.dLbls, flag, False)
        if sec["as"] == LINE:
            s.dLbls.position = "t"
    ch.title = f"{sec['title']} · {series_label(sec)}"
    ch.legend = None
    ch.y_axis.title = "Transactions"
    ch.y_axis.number_format = "#,##0"
    ch.x_axis.delete = ch.y_axis.delete = False                          # Excel hides the axes otherwise
    ch.y_axis.majorGridlines.spPr = None
    ch.width, ch.height = 16, 8
    return ch


def custom_excel(content: dict) -> bytes:
    """The custom report as an Excel file, laid out like the PDF:
    Overview - title band, key figures as cards, the line per currency, a linked list of the sections and every
               chosen chart (two per row, real Excel charts);
    then one sheet per section - its table, or the numbers behind its chart with the chart beside them, and its
               invoice list (with filters). Orientation and density chosen in the app are kept for printing."""
    from openpyxl import Workbook
    fs = 9 if content["density"] == "Compact" else 10
    sheet_of = {sec["key"]: sec["title"][:31] for sec in content["sections"]}
    wb = Workbook()
    ov = wb.active
    ov.title = "Overview"
    _xl_page(ov, content)
    for c in range(2, 18):
        ov.column_dimensions[get_column_letter(c)].width = 11
    _xl_band(ov, content["title"], f"{content['period']}  ·  {content['type']}"
             + (f"  ·  {content['stamp']}" if content["stamp"] else ""), 17, fs)
    row = 5

    # ---- key figures: one card each (label over a large value, green bar on the left), like the PDF
    if content["kpis"]:
        side = Side(style="thin", color=XL_RULE)
        for i, (label, value) in enumerate(content["kpis"]):
            c0 = 2 + i * 4
            for r in (row, row + 1):
                ov.merge_cells(start_row=r, start_column=c0, end_row=r, end_column=c0 + 2)
                for c in range(c0, c0 + 3):
                    ov.cell(r, c).border = Border(top=side if r == row else None, bottom=side if r > row else None,
                                                  left=Side(style="thick", color=XL_GREEN) if c == c0 else None,
                                                  right=side if c == c0 + 2 else None)
            ov.cell(row, c0, label.upper()).font = Font(size=fs - 1, bold=True, color=XL_MUTED)
            ov.cell(row + 1, c0, value).font = Font(size=fs + 8, bold=True, color=XL_TEAL)
            for r in (row, row + 1):
                ov.cell(r, c0).alignment = Alignment(indent=1, vertical="center")
        ov.row_dimensions[row].height, ov.row_dimensions[row + 1].height = 18, 30
        row += 3
    lines = content["currency_lines"] + ([content["exceptions_line"]] if content["exceptions_line"] else [])
    for line in lines:
        for c in range(2, 14):
            ov.cell(row, c).fill = _xl_fill(XL_TOTAL_FILL)
        ov.cell(row, 2, line).font = Font(size=fs, color=XL_INK)
        row += 1
    row += 1 if lines else 0

    # ---- the sections of this report, each a link to its sheet
    if content["sections"]:
        ov.cell(row, 2, "Sections").font = Font(size=fs + 2, bold=True, color=XL_INK)
        row += 1
        for sec in content["sections"]:
            shown = sec["as"] if sec["as"] != TABLE else "Table"
            cell = ov.cell(row, 2, sec["title"])
            cell.hyperlink, cell.font = f"#'{sheet_of[sec['key']]}'!A1", Font(size=fs, color="0563C1",
                                                                             underline="single")
            ov.cell(row, 5, shown).font = Font(size=fs, color=XL_MUTED)
            if sec["list"] is not None:
                ov.cell(row, 7, f"{len(sec['list']):,} invoices").font = Font(size=fs, color=XL_MUTED)
            row += 1
        row += 1

    # ---- one sheet per section; the charts are also placed on the Overview, two per row
    charts = []
    for sec in content["sections"]:
        s = wb.create_sheet(sheet_of[sec["key"]])
        _xl_page(s, content)
        width = max(8, len(sec["list"].columns) + 1 if sec["list"] is not None else 0,
                    len(sec["table"].columns) + 1 if sec["table"] is not None else 0)
        _xl_band(s, sec["title"], series_label(sec) if sec["series"] else content["period"], width + 1, fs,
                 back="Overview")
        r = 5
        if sec["table"] is not None:
            r = _xl_table(s, r, 2, sec["table"], fs, title="Summary", total_row=True, warning=sec["key"] == "sla")
        if sec["series"]:
            data = pd.DataFrame(sec["series"], columns=[sec["group"], "Transactions"])
            _xl_table(s, r, 2, data, fs, title="Numbers behind the chart")
            head = r + 1
            s.add_chart(_xl_chart(sec, s, head, len(data)), f"E{r}")
            charts.append(_xl_chart(sec, s, head, len(data)))
            r = max(r + len(data) + 3, r + 18)
        if sec["list"] is not None:
            title = f"{sec['list_title']} ({len(sec['list']):,})"
            if len(sec["list"]):
                _xl_table(s, r, 2, sec["list"], fs, title=title, filters=True)
                s.freeze_panes = s.cell(r + 2, 1)
            else:
                s.cell(r, 2, title).font = Font(size=fs + 1, bold=True, color=XL_INK)
                s.cell(r + 1, 2, "None.").font = Font(size=fs, color=XL_MUTED)
    if charts:
        ov.cell(row, 2, "Charts").font = Font(size=fs + 2, bold=True, color=XL_INK)
        row += 1
        for i, ch in enumerate(charts):
            ov.add_chart(ch, f"{'B' if i % 2 == 0 else 'J'}{row + (i // 2) * 17}")
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _nice_step(top: float) -> float:
    """A round step for a chart axis with about 4 lines (1, 2, 5, 10, 20, ...)."""
    raw = max(top / 4, 1)
    mag = 10 ** int(len(str(int(raw))) - 1)
    return next(m * mag for m in (1, 2, 5, 10) if m * mag >= raw)


def _pdf_chart(kind: str, points: list, width: float, height: float):
    """A one-series bar or line chart (IremboPay teal, light grid, no legend: the section title names it)."""
    from reportlab.graphics.charts.barcharts import VerticalBarChart
    from reportlab.graphics.charts.linecharts import HorizontalLineChart
    from reportlab.graphics.shapes import Drawing
    from reportlab.graphics.widgets.markers import makeMarker
    from reportlab.lib import colors

    teal, ink, grid = colors.HexColor(CHART_TEAL), colors.HexColor("#5B6B70"), colors.HexColor("#E1E7E9")
    labels, values = [p[0] for p in points] or ["–"], [p[1] for p in points] or [0]
    many = len(labels) > 12
    width = min(width, 120 + 70 * len(labels))           # few bars: a narrower chart, not very wide bars
    d = Drawing(width, height)
    d.hAlign = "LEFT"
    ch = VerticalBarChart() if kind == BAR else HorizontalLineChart()
    ch.x, ch.y = 36, 30 if many else 20
    ch.width, ch.height = width - 46, height - ch.y - 8
    ch.data = [values]
    step = _nice_step(max(values))
    ch.valueAxis.valueMin, ch.valueAxis.valueStep = 0, step
    ch.valueAxis.valueMax = step * (int(max(values) / step) + 1)
    ch.valueAxis.visibleGrid, ch.valueAxis.gridStrokeColor, ch.valueAxis.gridStrokeWidth = True, grid, 0.5
    ch.valueAxis.strokeColor, ch.valueAxis.visibleTicks = grid, False
    ch.valueAxis.labels.fontSize, ch.valueAxis.labels.fillColor = 7, ink
    ch.valueAxis.labels.fontName = ch.categoryAxis.labels.fontName = "Helvetica"
    ch.valueAxis.labelTextFormat = lambda v: f"{int(v):,}"
    step_l = 1 if ch.width / len(labels) >= 14 else 2                 # narrow chart: every second label
    ch.categoryAxis.categoryNames = [l if i % step_l == 0 else "" for i, l in enumerate(labels)]
    ch.categoryAxis.strokeColor, ch.categoryAxis.visibleTicks = grid, False
    ch.categoryAxis.labels.fontSize, ch.categoryAxis.labels.fillColor = 7, ink
    if many:
        ch.categoryAxis.labels.angle, ch.categoryAxis.labels.boxAnchor, ch.categoryAxis.labels.dy = 90, "e", -2
    if kind == BAR:
        ch.bars[0].fillColor, ch.bars[0].strokeColor = teal, None
        ch.groupSpacing = 2 if many else 8
    else:
        ch.lines[0].strokeColor, ch.lines[0].strokeWidth = teal, 2
        ch.lines[0].symbol = makeMarker("FilledCircle", size=5, fillColor=teal, strokeColor=colors.white,
                                        strokeWidth=1)
    d.add(ch)
    return d


def custom_pdf(content: dict) -> bytes:
    """The custom report as a PDF: title band, overview, then each section (table, chart and invoice list)."""
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.platypus import KeepTogether, LongTable, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
    from xml.sax.saxutils import escape

    teal, green = colors.HexColor("#123239"), colors.HexColor("#99E378")
    title_fill, grid, muted = colors.HexColor("#D0DEE3"), colors.HexColor("#9AA8AD"), colors.HexColor("#5B6B70")
    size = landscape(A4) if content["orientation"] == "Landscape" else A4
    page_w = size[0]
    margin = 14 * mm
    usable = page_w - 2 * margin
    compact = content["density"] == "Compact"
    fs, pad, gap = (7, 1.5, 4 * mm) if compact else (8.5, 3, 7 * mm)
    h1 = ParagraphStyle("h1", fontName="Helvetica-Bold", fontSize=16, textColor=green, leading=20)
    sub = ParagraphStyle("sub", fontName="Helvetica", fontSize=9.5, textColor=colors.HexColor("#D6E4E6"), leading=12)
    h2 = ParagraphStyle("h2", fontName="Helvetica-Bold", fontSize=12 if compact else 13, textColor=teal, leading=16,
                        spaceAfter=3)
    h3 = ParagraphStyle("h3", fontName="Helvetica-Bold", fontSize=fs + 1, textColor=teal, leading=fs + 4,
                        spaceBefore=4, spaceAfter=2)
    body = ParagraphStyle("body", fontName="Helvetica", fontSize=fs + 0.5, leading=fs + 4)
    small = ParagraphStyle("small", fontName="Helvetica", fontSize=fs, textColor=muted, leading=fs + 3)
    cell = ParagraphStyle("cell", fontName="Helvetica", fontSize=fs, leading=fs + 2)

    def table(df: pd.DataFrame, width: float, total_row=False):
        """A table: bold header, light grid, numbers on the right; long text wraps."""
        cols = list(df.columns)
        rows = [cols] + [[Paragraph(escape(cell_text(c, v)), cell) if c in ("Merchant", "Comment", "Decision")
                          else cell_text(c, v) for c, v in zip(cols, rec)] for rec in df.itertuples(index=False)]
        weights = [max([len(str(c))] + [len(cell_text(c, v)) for v in df[c].head(100)]) for c in cols]
        weights = [min(max(w, 6), 40) for w in weights]
        widths = [width * w / sum(weights) for w in weights]
        numeric = [j for j, c in enumerate(cols) if pd.api.types.is_numeric_dtype(df[c]) or c == "% of processed"]
        t = LongTable(rows, colWidths=widths, repeatRows=1, hAlign="LEFT")
        style = [("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"), ("BACKGROUND", (0, 0), (-1, 0), title_fill),
                 ("FONTSIZE", (0, 0), (-1, -1), fs), ("GRID", (0, 0), (-1, -1), 0.4, grid),
                 ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                 ("TOPPADDING", (0, 0), (-1, -1), pad), ("BOTTOMPADDING", (0, 0), (-1, -1), pad)]
        style += [("ALIGN", (j, 1), (j, -1), "RIGHT") for j in numeric]
        if total_row:
            style.append(("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"))
        t.setStyle(TableStyle(style))
        return t

    # ---- title band + overview
    band = Table([[Paragraph(escape(content["title"]), h1)],
                  [Paragraph(escape(" · ".join(x for x in (content["period"], content["type"], content["stamp"]) if x)),
                             sub)]], colWidths=[usable])
    band.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), teal), ("LEFTPADDING", (0, 0), (-1, -1), 10),
                              ("TOPPADDING", (0, 0), (0, 0), 8), ("BOTTOMPADDING", (0, -1), (-1, -1), 8),
                              ("LINEBEFORE", (0, 0), (0, -1), 5, green)]))
    story = [band]
    if content["kpis"]:
        n = len(content["kpis"])
        k = Table([[v for _, v in content["kpis"]], [label for label, _ in content["kpis"]]], colWidths=[usable / n] * n)
        k.setStyle(TableStyle([("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                               ("FONTSIZE", (0, 0), (-1, 0), 13 if compact else 16), ("TEXTCOLOR", (0, 0), (-1, 0), teal),
                               ("FONTSIZE", (0, 1), (-1, 1), fs), ("TEXTCOLOR", (0, 1), (-1, 1), muted),
                               ("ALIGN", (0, 0), (-1, -1), "CENTER"), ("BOX", (0, 0), (-1, -1), 0.6, grid),
                               ("LINEBEFORE", (1, 0), (-1, -1), 0.6, grid),
                               ("TOPPADDING", (0, 0), (-1, 0), 4 if compact else 7),
                               ("BOTTOMPADDING", (0, 1), (-1, 1), 4 if compact else 7)]))
        story += [Spacer(1, gap), k]
    lines = content["currency_lines"] + ([content["exceptions_line"]] if content["exceptions_line"] else [])
    if lines:
        story += [Spacer(1, gap / 2)] + [Paragraph(escape(line), body) for line in lines]

    def invoice_list(sec, named=False):
        """The invoices of a section, under its table, or under its chart card (named: the section is said)."""
        if sec["list"] is None:
            return []
        title = (f"{sec['title']} · " if named else "") + f"{sec['list_title']} ({len(sec['list']):,})"
        return [Paragraph(escape(title), h3),
                table(sec["list"], usable) if len(sec["list"]) else Paragraph("None.", small)]

    def chart_card(sec, width):
        """One chart in its own card: title, what is counted, the chart, the values (for print)."""
        inner = width - 16
        card = Table([[[Paragraph(escape(sec["title"]), h3), Paragraph(escape(series_label(sec)), small),
                        _pdf_chart(sec["as"], sec["series"], inner, (42 if compact else 55) * mm),
                        Paragraph(escape(" · ".join(f"{name} {v:,}" for name, v in sec["series"])), small)]]],
                     colWidths=[width], hAlign="LEFT")
        card.setStyle(TableStyle([("BOX", (0, 0), (-1, -1), 0.6, grid), ("VALIGN", (0, 0), (-1, -1), "TOP"),
                                  ("LEFTPADDING", (0, 0), (-1, -1), 8), ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                                  ("TOPPADDING", (0, 0), (-1, -1), 6), ("BOTTOMPADDING", (0, 0), (-1, -1), 6)]))
        return card

    # ---- the sections: charts side by side (two per row, each in a card), tables and lists full width
    space = 5 * mm                                      # between two chart cards
    for kind, item in layout_rows(content["sections"]):
        if kind == "section":
            sec = item
            part = [Spacer(1, gap), Paragraph(escape(sec["title"]), h2)]
            if sec["table"] is not None:
                part.append(table(sec["table"], min(usable, 150 * mm), total_row=True))
            story.append(KeepTogether(part))
            story += invoice_list(sec)
            continue
        half = (usable - space) / 2
        cards = [chart_card(sec, half if len(item) == 2 else usable) for sec in item]
        row = Table([cards], colWidths=[half + space, half] if len(cards) == 2 else [usable], hAlign="LEFT")
        row.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0),
                                 ("RIGHTPADDING", (0, 0), (0, -1), space if len(cards) == 2 else 0),
                                 ("RIGHTPADDING", (1, 0), (-1, -1), 0),
                                 ("TOPPADDING", (0, 0), (-1, -1), 0), ("BOTTOMPADDING", (0, 0), (-1, -1), 0)]))
        story += [Spacer(1, gap), row]
        for sec in item:
            story += invoice_list(sec, named=True)

    def footer(canvas, doc):
        canvas.saveState()
        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(muted)
        canvas.drawRightString(page_w - margin, 8 * mm, f"Page {doc.page}")
        canvas.restoreState()

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=size, leftMargin=margin, rightMargin=margin, topMargin=margin,
                            bottomMargin=margin, title=f"{content['title']} {content['period']}")
    doc.build(story, onFirstPage=footer, onLaterPages=footer)
    return buf.getvalue()
