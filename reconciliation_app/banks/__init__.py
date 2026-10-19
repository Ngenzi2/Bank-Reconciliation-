"""
BANKS
=====
The banks the app can reconcile. One file per bank holds what is specific to that bank; the common rules are in
engine.py. The app checks every bank file at start-up and says exactly what is missing.

    im.py        I&M Bank   (the template: the simplest bank, every optional setting left at its default)
    ecobank.py   Ecobank    (an example of every optional setting)
    boa.py       Bank of Africa Rwanda (works like I&M: credits only, Value Date vs paid_at; statement checked
                 against its own totals; matching by invoice, or by Ref no when the description has none)

To add a bank
    1. copy im.py to a new file (e.g. bk.py)
    2. change its values and functions (the list REQUIRED below says what each one is)
    3. if the bank works differently from I&M, add the settings it needs from the list OPTIONAL below
    4. add it to ACTIVE below; with more than one bank, step 1 of the app shows a "Bank" choice

CONTENTS
    1. Active banks
    2. What every bank file must provide (REQUIRED)
    3. Optional settings (OPTIONAL): leave them out to work like I&M
    4. Checking the bank files at start-up
"""
from . import boa, ecobank, im


# =====================================================================================================================
# 1. ACTIVE BANKS
# =====================================================================================================================
ACTIVE = [im, ecobank, boa]     # the first one is the default


# =====================================================================================================================
# 2. WHAT EVERY BANK FILE MUST PROVIDE
# =====================================================================================================================
REQUIRED = {
    "CODE": "short code used in report file names, e.g. 'IM'",
    "NAME": "short name, e.g. 'I&M'",
    "FULL_NAME": "full name, e.g. 'I&M Bank'",
    "PROVIDER": "payment_provider of this bank in the IremboPay payment file",
    "CURRENCIES": "list of currencies, one statement per currency",
    "FILE_TYPES": "statement file types accepted, e.g. ['xls', 'xlsx']",
    "STATEMENT_SHEET": "name of the statement sheet in the report",
    "COLUMNS": "statement column names for posted_date, payment_date, description, reference, amount",
    "MONDAY_AFTER_WEEKEND": "True when a period ending on a weekend also needs Monday's statement",
    "DAY_BEFORE_CURRENCIES": "currencies for which the statement of the day before is recommended",
    "COMPLETED_STATUSES": "IremboPay statuses that count as paid",
    "read_statement": "read one statement file -> (rows, other_credits, info)",
    "clean_statement": "typed columns + invoice_number",
    "channel": "bank channel of a row, shown in the review",
    "sla": "(days difference, SLA status) per row",
    "received": "True when the notification was received by Irembo",
    "pending": "True when the invoice is still pending (NEW)",
    "duplicate": "True when the invoice was paid with another channel",
}
COLUMN_KEYS = ["posted_date", "payment_date", "description", "reference", "amount"]   # posted_date may be None


# =====================================================================================================================
# 3. OPTIONAL SETTINGS: a bank leaves these out to work like I&M (see ecobank.py for an example of each)
# =====================================================================================================================
OPTIONAL = {
    # ---- identity
    "PROVIDERS": "payment_provider values that count as this bank (default: [PROVIDER]); use when IremboPay gives "
                 "the same bank more than one code (e.g. BPR: 'BPR' and 'BPR_RW')",
    # ---- files and checks (steps 2 and 3)
    "ONE_FILE_PER_CURRENCY": "False when one statement holds every currency (default True)",
    "MISSING_DAYS_BLOCK": "False to allow Run with a warning when days have no statement (default True)",
    "ACCEPT_EARLIER_DAYS": "days before the period whose statements are accepted (default 0: only the day before)",
    "REVERSE_CHECK": "True to list, in step 3, IremboPay payments of this bank on no statement (default False)",
    "PAYMENT_FILE_TYPES": "payment file types, e.g. ['csv', 'xlsx'] (default ['csv'])",
    "PAYMENT_OPTIONAL_COLUMNS": "payment file columns that may be missing, e.g. ['created_at']",
    # ---- SLA
    "SLA_DAYS_SETTING": "False when the SLA rule is fixed (the SLA days setting is hidden)",
    "SLA_DIFF_COLUMN": "name of the SLA difference column (default 'sla_days_difference')",
    # ---- statement rows and the report
    "DUPLICATE_KEY": "columns that make two statement rows the same export",
    "REVIEW_KEY_WITHOUT_REFERENCE": "True when the review recognises a row by currency | date | invoice only (the "
                                    "bank reference is not in the IremboPay files; default False)",
    "DETAIL_COLUMNS": "extra statement columns shown in the detail sheets",
    "DATETIME_COLUMNS": "statement columns with a time, formatted with hh:mm:ss",
    "TOTALS_CHECK_SKIP": "summary tables left out of the 'weeks add up to the month' check",
    # ---- the bank's own rules (functions)
    "one_row_per_invoice": "function(rows, payments): keep one statement row per invoice",
    "in_period": "function(rows, attempts, start, end): which rows belong to the period",
    "tables": "function(rows, attempts, a, b): rows of the 5 tables (default: by payment date)",
    "extra_sheets": "function(rows, attempts, start, end, whole): {sheet name: table} added to the report",
}


# =====================================================================================================================
# 4. CHECKING THE BANK FILES AT START-UP
# =====================================================================================================================
def check(bank):
    """Raise ValueError naming everything a bank file is missing."""
    file = bank.__name__.replace(".", "/") + ".py"
    missing = [f"{name} ({why})" for name, why in REQUIRED.items() if not hasattr(bank, name)]
    if hasattr(bank, "COLUMNS"):
        missing += [f"COLUMNS['{k}']" for k in COLUMN_KEYS if k not in bank.COLUMNS]
    if missing:
        raise ValueError(f"The bank file {file} is missing: " + "; ".join(missing))


def load() -> dict:
    """{code: bank module} for the active banks, after checking each one."""
    out = {}
    for bank in ACTIVE:
        check(bank)
        if bank.CODE in out:
            raise ValueError(f"Two bank files use the same CODE '{bank.CODE}'.")
        out[bank.CODE] = bank
    return out
