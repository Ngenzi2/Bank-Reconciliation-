"""Bank of Kigali: made-up statements in the BK layout (no real invoice or customer data; the 5 account numbers are the
Irembo collection accounts of banks/bk.py). The rules: the BK notebook, with the 4 fixes decided on 03/11/2026 and the
credit types agreed on 04/11/2026 (Transactions matched and checked; card payments and bank transfers counted, never
reviewed; debits not counted)."""
import io

import pandas as pd
import pytest
from openpyxl import Workbook

import engine as e
from banks import bk

HEAD = ["RowNum", "RECTYPE", "Currency", "AccountNumber", "BookingDate", "ValueDate", "ProcessingDateTime",
        "TransactionType", "TheirReference", "Source", "OurReference", "DebitAcct", "CreditAcct", "CreditAmt",
        "DebitAmt", "Narration", "ApiId", "Balance"]
RWF, USD = "100081756273", "100081756467"
YACU = "New BK Yacu BILL PAYMENT IREMBO"


def statement(rows, account=RWF, currency="RWF", head=HEAD, day="20260803"):
    """rows: (ValueDate, Source, OurReference, CreditAmt, DebitAmt, Narration) -> CSV bytes with the H1 / F1 rows."""
    out = [",".join(head), ",".join(["1", "H1BKI   ", currency, account, day, day] + [""] * 11 + ["1000"])]
    for i, (vd, source, ref, credit, debit, narr) in enumerate(rows, 2):
        values = {"RowNum": str(i), "RECTYPE": "D1BKIGRW", "Currency": currency, "AccountNumber": account,
                  "BookingDate": vd, "ValueDate": vd, "ProcessingDateTime": vd[2:] + "101500", "TransactionType": "",
                  "TheirReference": "", "Source": source, "OurReference": ref, "DebitAcct": "", "CreditAcct": "",
                  "CreditAmt": str(credit), "DebitAmt": str(debit), "Narration": narr, "ApiId": "", "Balance": ""}
        out.append(",".join(values[_name(h)] for h in head))
    out.append(",".join([str(len(rows) + 2), "F1BKI", currency, account, day, day] + [""] * 11 + ["2000"]))
    return "\n".join(out).encode()


def _name(h):
    return {"RECORDNUM": "RowNum", "ACCOUNT_CURRENCY": "Currency"}.get(h) or next(c for c in HEAD if c.upper() == h.upper())


def payments(rows):
    """rows: (invoice, amount, currency, paid_at, provider, provider_reference, status[, channel])."""
    cols = ["invoice_number", "transaction_id", "amount", "currency", "payment_status", "description", "expiry_at",
            "created_at", "settled_at", "paid_at", "name", "transaction_reference", "provider_reference",
            "payment_channel", "payment_provider", "transaction_status"]
    out = [",".join(cols)]
    for inv, amount, cur, paid, provider, pref, status, *chan in rows:
        out.append(",".join([inv, "T" + inv, str(amount), cur, status, "", "", "2026-08-01 08:00:00", "",
                             paid if status != "NEW" else "", "RSSB", "", pref, (chan or ["BK_AGENT"])[0], provider,
                             "SUCCESSFUL"]))
    return "\n".join(out).encode()


def read(content, name=f"IREMBO_{RWF}_STMT_20260803_202608040612_1.csv"):
    return e.read_bank_file(name, content, "BK")


ROWS = [  # ValueDate, Source, OurReference, CreditAmt, DebitAmt, Narration
    ("20260803", YACU, "FTCM1", 5000, 0, "BK YACU BILL PAYMENT IREMBO :  / 886000000001 Irembo bill"),   # by reference
    ("20260803", YACU, "FTCM2", 2000, 0, "BK YACU BILL PAYMENT IREMBO :  / 886000000002 Irembo bill"),   # by narration
    ("20260803", "Branch Teller", "TT3", 700, 0, "Cash Deposit : 886000000003/PAYMENT"),                # paid next day
    ("20260803", YACU, "FTCM4", 300, 0, "BK YACU BILL PAYMENT IREMBO :  / 886000000004 Irembo bill"),    # still NEW
    ("20260803", YACU, "FTCM5", 400, 0, "BK YACU BILL PAYMENT IREMBO :  / 886000000005 Irembo bill"),    # other provider
    ("20260803", YACU, "FTCM6", 900, 0, "BK YACU BILL PAYMENT IREMBO :  / 886000000006 Irembo bill"),    # transfer (fix 4)
    ("20260803", YACU, "FTCM7", 100, 0, "BK YACU BILL PAYMENT IREMBO :  / 886000000007 Irembo bill"),    # not found
    ("20260803", YACU, "FTCM8", 1000, 0, "BK YACU BILL PAYMENT IREMBO :  / 886000000008 Irembo bill"),   # credited twice:
    ("20260803", YACU, "FTCM9", 1000, 0, "BK YACU BILL PAYMENT IREMBO :  / 886000000008 Irembo bill"),   # IremboPay = FTCM9
    ("20260803", "E- commerce purchase", "FT10", 8000, 0, "E- commerce purchase : EC2608031"),          # card payment
    ("20260803", "Card Purchase Advices", "FT11", 6000, 0, "Purchase-settlement advices : 4213xx"),     # card payment
    ("20260803", "IB Account transfer", "FT12", 50000, 0, "IB Account transfer : IREMBO LTD / Refund"),  # bank transfer
    ("20260803", "", "FT13", 0, -20, "Account Transfer Charges : AC-1"),                                 # a debit
    ("20260803", "AGENT REFUND", "FT14", 0, -500, "AGENT REFUND : IREMB1 / FT14 886000000001"),          # a refund
    ("20260810", YACU, "FTCM15", 5000, 0, "BK YACU BILL PAYMENT IREMBO :  / 886000000015 Irembo bill"),  # next week
]
PAYS = [("886000000001", 5000, "RWF", "2026-08-03 10:00:00", "BK", "FTCM1", "SETTLED"),
        ("886000000002", 2000, "RWF", "2026-08-03 11:00:00", "BK", "", "PAID"),
        ("886000000003", 700, "RWF", "2026-08-04 09:00:00", "BK", "TT3", "SETTLED"),
        ("886000000004", 300, "RWF", "", "BK", "", "NEW"),
        ("886000000005", 400, "RWF", "2026-08-03 12:00:00", "MTN", "MP5", "SETTLED"),
        ("886000000006", 900, "RWF", "2026-08-03 13:00:00", "IREMBO", "", "SETTLED", "BANK_TRANSFER_BK"),
        ("886000000008", 1000, "RWF", "2026-08-03 14:00:00", "BK", "FTCM9", "SETTLED")]
WEEK = e.make_period("week", "2026-08-03")


def run(tmp_path, rows=ROWS, pays=PAYS, require_review=False, extra_files=()):
    files = [read(statement(rows)), *extra_files, e.read_payment_file("p.csv", payments(pays), "BK")]
    return e.run(tmp_path / "base", WEEK, files, require_review=require_review, bank="BK", report_dir=tmp_path / "r")


def counts(res, title):
    t = res.blocks[0][2][title]
    return dict(zip(t["Currency"].fillna("All"), t["Number of Transactions"]))


# ---- the bank is active ---------------------------------------------------------------------------------------------
def test_bk_is_an_active_bank():
    assert "BK" in e.BANKS and e.BANKS["BK"] is bk
    assert set(bk.EXTRA_SHEETS) <= e.REPORT_SHEETS


# ---- reading --------------------------------------------------------------------------------------------------------
def test_daily_file_keeps_only_the_movements_and_covers_its_day():
    f = read(statement(ROWS))
    assert f.rows == len(ROWS)                                   # H1 (opening) and F1 (closing) are not movements
    assert f.currency == "RWF" and f.bank == "BK" and f.balances["account"] == RWF
    assert (f.first, f.last) == (pd.Timestamp("2026-08-03"), pd.Timestamp("2026-08-10"))   # the dates inside
    assert f.other_credits == 3                                  # 2 card payments + 1 bank transfer


@pytest.mark.parametrize("name", ["statement.csv",                                  # no date in the name
                                  f"IREMBO_{RWF}_STMT_20260901_202609020612_1.csv",   # a wrong date in the name
                                  f"IREMBO_ACCOUNT_{RWF}_MONTHLY_STMT_AUGUST_2026_202609031036.csv"])
def test_the_days_covered_come_from_the_data_never_the_name(name):
    f = read(statement(ROWS, day="20260801"), name=name)              # H1 dated 01/08, movements 03/08 to 10/08
    assert (f.first, f.last) == (pd.Timestamp("2026-08-01"), pd.Timestamp("2026-08-10"))
    assert f.rows == len(ROWS)


def test_a_file_with_only_its_h1_and_f1_rows_covers_their_day():
    f = read(statement([]), name="empty day.csv")                    # a day without any movement
    assert f.rows == 0 and f.first == f.last == pd.Timestamp("2026-08-03")


def test_capital_names_and_another_column_order_are_read():
    head = ["RECORDNUM", "RECTYPE", "ACCOUNT_CURRENCY", "ACCOUNTNUMBER", "BOOKINGDATE", "VALUEDATE",
            "PROCESSINGDATETIME", "TRANSACTIONTYPE", "SOURCE", "THEIRREFERENCE", "OURREFERENCE", "DEBITACCT",
            "CREDITACCT", "CREDITAMT", "DEBITAMT", "NARRATION", "APIID", "BALANCE"]
    f = read(statement(ROWS, head=head))
    data = bk.clean_statement(f.data)
    assert f.rows == len(ROWS) and data.loc[0, "Source"] == YACU and data.loc[0, "OurReference"] == "FTCM1"


def test_xlsx_statement_is_read():
    wb = Workbook()
    for line in statement(ROWS).decode().splitlines():
        wb.active.append(line.split(","))
    buf = io.BytesIO()
    wb.save(buf)
    f = read(buf.getvalue(), name=f"IREMBO_{RWF}_STMT_20260803_202608040612_1.xlsx")
    assert f.rows == len(ROWS) and f.currency == "RWF"


@pytest.mark.parametrize("account, currency, words", [
    ("100010639478", "RWF", "not one of the 5 Irembo collection accounts"),     # Irembo's operating account
    ("1.00E+11", "RWF", "opened and saved in Excel"),                           # an Excel-damaged file
    (USD, "RWF", "is a USD account"),
])
def test_a_wrong_file_is_refused_with_a_plain_message(account, currency, words):
    with pytest.raises(e.ReconError, match=words):
        read(statement(ROWS[:2], account=account, currency=currency))


def test_a_file_that_is_not_a_bk_statement_is_refused():
    with pytest.raises(e.ReconError, match="not a BK statement"):
        read(b"invoice_number,amount\n1,2\n", name="other.csv")


# ---- the row types and the invoice --------------------------------------------------------------------------------
def test_row_types_follow_the_notebook_with_the_swift_charges_fixed():
    rows = pd.DataFrame({"Source": ["Card Purchase", "Card Purchase Advices", "E- commerce purchase",
                                    "Incoming Trsf frm local banks", "", "", "AGENT REFUND", YACU],
                         "Narration": ["Purchase-settlement : 4713", "Purchase-settlement advices : 42",
                                       "E- commerce purchase : EC1", "From X", "SWIFT Payment from ABC",
                                       "SWIFT Payment - Charges", "AGENT REFUND : IREMB1", "/ 886000000001"]})
    assert bk.row_types(rows).tolist() == ["POS TXN", "Card payment", "Card payment", "Bank transfer", "Bank transfer",
                                           "Bank transfer Charges", "Refund", "Transaction"]


def test_credits_that_are_not_bill_payments_get_their_own_type():
    rows = pd.DataFrame({"Source": ["VISA SETTLEMENT", "In-House Cheques", "Outward bulk Transfer On-Us", "", "",
                                    "Branch Teller", "Branch Teller"],
                         "Narration": ["VISA SETTLEMENT : NH1 / 20260703", "Cheque Deposit : APTC",
                                       "Outward bulk Transfer On-Us : J1 / RETURN", "EKASH incoming transfer : B R",
                                       "Cash deposit-BK YACU : A B / KWISHYURA AMANDE",
                                       "Cash Deposit : 886000000003/PAYMENT", "Cash Deposit : NCGR LTD /  / NCGR LTD"]})
    assert bk.row_types(rows).tolist() == ["Card payment", "Bank transfer", "Bank transfer", "Bank transfer",
                                           "Cash deposit", "Transaction",        # a Branch Teller payment stays one
                                           "Cash deposit"]                       # 09/11/2026: a name, no invoice


def test_invoice_is_cut_from_the_narration_by_source():
    rows = pd.DataFrame({"Source": [YACU, "Branch Teller", "MTN Irembo payments", YACU, "Card Purchase Advices"],
                         "Narration": ["BK YACU :  / 886000000001 Irembo bill", "Cash Deposit : 886000000002/PAY",
                                       "MTN : 886 000 000 003", "BK YACU BILL PAYMENT IREMBO :",
                                       "Purchase / 886000000005"]})
    inv, rule = bk.narration_invoice(rows)
    assert inv.tolist() == ["886000000001", "886000000002", "886000000003", "", ""]
    assert rule[0] == "Narration (after the last /)" and rule[2] == "Narration (after the first :)"


def test_reference_first_then_narration_and_a_shared_reference_is_not_used():
    s = bk.clean_statement(read(statement(ROWS[:2] + [ROWS[9]])).data)
    p = pd.DataFrame({"invoice_number": ["886000000099", "886000000002", "886000000003"],
                      "provider_reference": ["FTCM1", "FT10", "FT10"]})
    out = bk.find_invoices(s, p)
    assert out.loc[0, "invoice_number"] == "886000000099"                 # step 1 wins over the narration
    assert out.loc[0, bk.FOUND_BY] == "OurReference = provider_reference"
    assert out.loc[1, "invoice_number"] == "886000000002" and out.loc[1, bk.FOUND_BY].startswith("Narration")
    assert pd.isna(out.loc[2, "invoice_number"])                           # a card payment: never matched


# ---- the figures ----------------------------------------------------------------------------------------------------
def test_the_figures_follow_the_notebook(tmp_path):
    res = run(tmp_path)
    assert counts(res, "Overall processed payments") == {"RWF": 11, "USD": 0, "EUR": 0, "GBP": 0, "All": 11}
    # every credit of the week (8 bill payments + 2 card payments + 1 transfer), the invoice credited twice once;
    # the debit and the refund are not counted, nor the credit of the next week
    assert counts(res, "Overall notifications received by Irembo")["RWF"] == 5   # 1, 2, 3, 6 (fix 4), 8
    assert counts(res, "SLA Violated")["RWF"] == 1                               # 3: paid the next day
    assert counts(res, "Pending notifications")["RWF"] == 1                      # 4: NEW
    assert counts(res, "Duplicate Payment")["RWF"] == 1                          # 5: paid with MTN
    assert e.report_figures(res.blocks)["processed"] == 11


def test_only_the_bill_payments_go_to_review(tmp_path):
    res = run(tmp_path, require_review=True)
    assert res.stopped_for_review and res.not_found == 1                  # 7: a bill payment not found
    reviewed = res.review
    assert set(reviewed["invoice_number"]) == {"886000000007", "886000000008"}     # + the second credit of 8
    assert not reviewed["Narration"].str.contains("E- commerce|Purchase-settlement|IB Account").any()


def test_a_bill_payment_without_invoice_number_goes_straight_to_pending(tmp_path):
    empty = ("20260803", YACU, "FTCM16", 3000, 0, "BK YACU BILL PAYMENT IREMBO :")     # nothing after the colon
    res = run(tmp_path, rows=ROWS + [empty], require_review=True)
    assert "FTCM16" not in set(res.review["OurReference"])                # never in Exception Resolution
    res = run(tmp_path, rows=ROWS + [empty])
    assert counts(res, "Overall processed payments")["RWF"] == 12          # counted in Processed
    assert counts(res, "Overall notifications received by Irembo")["RWF"] == 5   # not received
    assert counts(res, "Pending notifications")["RWF"] == 2                # 4 (NEW) + this one
    assert "FTCM16" in set(res.sheets["Pending notification"]["OurReference"])
    assert "FTCM16" not in set(res.sheets["Not found (BK)"]["OurReference"])
    assert res.sheets["Matched details"].set_index("OurReference").at["FTCM16", "Reason"] == bk.NO_INVOICE


def test_the_credit_irembopay_recorded_is_the_one_counted(tmp_path):
    d = run(tmp_path).sheets["Matched details"]
    eight = d[d["invoice_number"].eq("886000000008")].set_index("OurReference")
    assert not eight.at["FTCM9", "is_extra_credit"] and eight.at["FTCM8", "is_extra_credit"]
    assert eight.at["FTCM8", "Counted credit"] == "FTCM9" and eight.at["FTCM9", "Counted credit"] == ""
    assert list(d.columns[d.columns.get_loc("Result"):][:4]) == ["Result", "Counted credit", "review_status",
                                                               "review_comment"]


def test_payment_details_say_why_a_credit_is_not_checked(tmp_path):
    d = run(tmp_path).sheets["Matched details"].set_index("OurReference")
    assert d.at["FT10", "Reason"] == "Not an Irembo bill payment (Card payment)"
    assert d.at["FT12", "Reason"] == "Not an Irembo bill payment (Bank transfer)"
    assert d.at["FT10", "sla_status"] == "Not applicable" and d.at["FTCM5", "sla_status"] == "Not processed"
    assert d.at["FTCM7", "Reason"] == "Not found in IremboPay"


# ---- the report -----------------------------------------------------------------------------------------------------
def test_the_report_has_the_bk_sheets(tmp_path):
    res = run(tmp_path)
    sheets = pd.ExcelFile(res.report_file).sheet_names
    assert all(s in sheets for s in bk.EXTRA_SHEETS)
    late = res.sheets["SLA violated (BK)"]
    assert late["invoice_number"].tolist() == ["886000000003"]
    assert late["Time diff"].iloc[0] == "09hrs 01days" and late["Charge applicable"].iloc[0] == "Yes"
    assert res.sheets["Bank transfer (BK)"]["OurReference"].tolist() == ["FT12"]
    assert sorted(res.sheets["Other credits (BK)"]["OurReference"]) == ["FT10", "FT11", "FT12"]
    assert res.sheets["Not found (BK)"]["OurReference"].tolist() == ["FTCM7"]
    debits = res.sheets["Debits (BK)"].set_index("OurReference")
    assert sorted(debits.index) == ["FT13", "FT14"]
    assert debits.at["FT13", bk.TYPE] == "Debit" and debits.at["FT14", bk.TYPE] == "Refund"


def test_payment_details_holds_the_irembopay_rows_and_matched_details_every_credit(tmp_path):
    res = run(tmp_path)
    sheets = pd.ExcelFile(res.report_file).sheet_names
    assert "Matched details" in sheets and "Payment details" in sheets
    pay = res.sheets["Payment details"]
    assert list(pay.columns) == list(e.read_payment_file("p.csv", payments(PAYS), "BK").data.columns)   # its columns
    assert "886000000005" in set(pay["invoice_number"])                  # paid in the week (MTN), on the statement
    assert "CreditAmt" not in pay and len(res.sheets["Matched details"]) == 12      # every credit of the week
    assert e.details_sheet("BK", sheets) == "Matched details" and e.details_sheet("IM") == "Payment details"
    assert e.details_sheet("BK", ["Payment details"]) == "Payment details"         # a BK report made before


def test_matched_details_says_in_one_word_what_happened_to_each_credit(tmp_path):
    empty = ("20260803", YACU, "FTCM16", 3000, 0, "BK YACU BILL PAYMENT IREMBO :")
    d = run(tmp_path, rows=ROWS + [empty]).sheets["Matched details"].set_index("OurReference")
    assert d.columns[d.columns.get_loc("CreditAmt") + 1] == "Result"                 # next to the amount
    assert {ref: d.at[ref, "Result"] for ref in ("FTCM1", "FTCM8", "FTCM7", "FTCM16", "FT10")} == {
        "FTCM1": "Matched", "FTCM8": "Second credit", "FTCM7": "Not found", "FTCM16": "No invoice number",
        "FT10": "Not a bill payment"}


def test_the_statement_sheet_shows_where_each_invoice_comes_from(tmp_path):
    s = run(tmp_path).sheets["BK statement"].set_index("OurReference")
    assert list(s.columns[-4:]) == ["invoice_number", bk.FOUND_BY, "provider_reference", "Source file"]
    assert (s.at["FTCM1", bk.FOUND_BY], s.at["FTCM1", "provider_reference"]) == (
        "OurReference = provider_reference", "FTCM1")                     # linked: the two references are the same
    assert s.at["FTCM2", bk.FOUND_BY] == "Narration (after the last /)"


def test_the_same_row_in_a_daily_and_a_monthly_file_counts_once(tmp_path):
    monthly = read(statement(ROWS), name=f"IREMBO_ACCOUNT_{RWF}_MONTHLY_STMT_AUGUST_2026_202609031036.csv")
    res = run(tmp_path, extra_files=[monthly])
    assert counts(res, "Overall processed payments")["All"] == 11


def test_payment_per_merchant_has_a_band_per_currency_and_reads_back(tmp_path):
    usd = [("20260803", YACU, "FTCM20", 10, 0, "BK YACU BILL PAYMENT IREMBO :  / 886000000020 Irembo bill")]
    files = [read(statement(ROWS)), read(statement(usd, account=USD, currency="USD"),
                                         name=f"IREMBO_{USD}_STMT_20260803_202608040612_1.csv"),
             e.read_payment_file("p.csv", payments(PAYS + [("886000000020", 10, "USD", "2026-08-03 10:00:00", "BK",
                                                              "FTCM20", "SETTLED")]), "BK")]
    res = e.run(tmp_path / "base", WEEK, files, require_review=False, bank="BK", report_dir=tmp_path / "r")
    ws = __import__("openpyxl").load_workbook(res.report_file)["Summary-Weekly"]
    top = next(r for r in range(1, ws.max_row + 1) for c in (6, 7) if ws.cell(r, c).value == e.MERCHANT_TITLE)
    col = 6 if ws.cell(top, 6).value == e.MERCHANT_TITLE else 7
    below = [ws.cell(r, col).value for r in range(top + 2, top + 9)]
    assert below[0] == "RWF" and ws.cell(top + 2, col).font.b and ws.cell(top + 2, col + 1).value is None
    assert "USD" in below and below[below.index("USD") - 1] == "All"
    assert e.read_merchant_counts(res.report_file, "Summary-Weekly") == e.merchant_counts(
        res.merchants[res.blocks[0][1]])                                   # the bands are skipped when read back


# ---- the files of step 3 --------------------------------------------------------------------------------------------
def test_every_collection_account_needs_a_file():
    files = [read(statement(ROWS)), e.read_payment_file("p.csv", payments(PAYS), "BK")]
    chk = e.check_period(e.make_period("day", "2026-08-03"), files, "BK", today="2026-11-04")
    assert any(RWF in t and ok for ok, t in chk.lines)
    assert sum("no file added" in b for b in chk.blocking) == 4          # the 4 other accounts
