"""Bank of Africa (BOA): reading and checking the statement, categories, matching with IremboPay, the period, the
SLA, the report, the app. Made-up files with the real structure (boa_fixtures.py), plus the real samples when
BANK_RECON_SAMPLES points to them (skipped otherwise)."""
import os
from pathlib import Path

import pandas as pd
import pytest

import engine as e
from banks import boa
from boa_fixtures import payments, statement
from conftest import APP_DIR, ROOT

SAMPLES = Path(os.environ["BANK_RECON_SAMPLES"]) if os.environ.get("BANK_RECON_SAMPLES") else None

SEP = e.make_period("month", "2026-09-01")

# the made-up September statement, oldest first: (Value Date, Ref no, Description, Debit, Credit)
ROWS = [
    ("30/08/2026", "S00000", "ePay irembo 880000000000", None, 1000),        # oldest day: may be cut by "Last N"
    ("01/09/2026", "S00001", "ePay irembo 880000000001", None, 5000),        # matched (internet banking)
    ("02/09/2026", "S00002", "irembo Payment for 880000000002", None, 12000),  # matched (USSD), paid_at 3 days late
    ("03/09/2026", "S00003", "YOUR CASH DEPOSIT 880000000003", None, 7500),  # IremboPay still NEW: pending
    ("04/09/2026", "S00004", "irembo Invoice Permit", None, 100000),         # no invoice: matched by Ref no
    ("05/09/2026", "S00005", "ePay irembo 880000000005", None, 3000),        # IremboPay says another channel
    ("06/09/2026", "S00006", "ePay irembo 880000000006", None, 4000),        # not in IremboPay
    ("07/09/2026", "S00007", "ePay irembo 880000000007", None, 2000),        # IremboPay amount 2,500
    ("08/09/2026", "S00008", "ePay irembo 880000000008", None, 9000),        # credited twice: IremboPay has S00008
    ("08/09/2026", "S00009", "ePay irembo 880000000008", None, 9000),
    ("09/09/2026", "S00010", "ePay irembo 880000000010", None, 600),         # reversed below
    ("10/09/2026", "S00010", "Cancel ePay irembo 880000000010", -600, None),
    ("11/09/2026", "S00011", "IBank monthly Sep 2026", -1000, None),
    ("12/09/2026", "S00012", "REFUND/IBank monthly Aug 2026", None, 1000),
    ("13/09/2026", "S00013", "Cancel ePay irembo 880000000099", -50, None),  # a debit naming Irembo: never a payment
    ("14/09/2026", "S00014", "VOC1/26 TRANSFER FAVOUR TEST LTD LOCKBOX", -5000, None),
    ("15/09/2026", "S00015", "SALARY PAYMENT", None, 300),
    ("16/09/2026", "S00016", "irembo Invoice Permit", None, 100000),         # Ref shared by 2 payments: ambiguous
    ("17/09/2026", "S00017", "ePay irembo 880000000017", None, 70),          # IremboPay status FAILED
    ("18/09/2026", "S00018", "ePay irembo 012345678901", None, 100),         # leading zero kept
    ("19/09/2026", "S00019", "irembo Invoice Permit", None, 5000),           # same amount as ...001: NOT matched
    ("30/09/2026", "S00030", "ePay irembo 880000000030", None, 1500),        # IremboPay currency USD
    ("01/10/2026", "S00031", "ePay irembo 880000000031", None, 800),         # October: not in September
]
PAYS = [
    dict(invoice="880000000001", amount=5000, paid_at="2026-09-01 10:00:00", ref="S00001"),
    dict(invoice="880000000002", amount=12000, paid_at="2026-09-05 09:00:00", ref="S00002", channel="BOA_USSD"),
    dict(invoice="880000000003", amount=7500, paid_at="2026-09-03 09:00:00", ref="S00003", status="NEW",
         channel="BOA_BRANCH"),
    dict(invoice="880000000004", amount=100000, paid_at="2026-09-04 09:00:00", ref="S00004",
         description="Invoice Permit"),
    dict(invoice="880000000005", amount=3000, paid_at="2026-09-05 09:00:00", ref="M1", provider="MTN_MOMO"),
    dict(invoice="880000000007", amount=2500, paid_at="2026-09-07 09:00:00", ref="S00007"),
    dict(invoice="880000000008", amount=9000, paid_at="2026-09-06 09:00:00", ref="S00008"),   # Sat -> Mon credit
    dict(invoice="880000000010", amount=600, paid_at="2026-09-09 09:00:00", ref="S00010"),
    dict(invoice="880000000161", amount=100000, paid_at="2026-09-16 09:00:00", ref="S00016"),
    dict(invoice="880000000162", amount=100000, paid_at="2026-09-16 09:30:00", ref="S00016"),
    dict(invoice="880000000017", amount=70, paid_at="2026-09-17 09:00:00", ref="S00017", status="FAILED"),
    dict(invoice="012345678901", amount=100, paid_at="2026-09-18 09:00:00", ref="S00018"),
    dict(invoice="880000000030", amount=1500, paid_at="2026-09-30 09:00:00", ref="S00030", currency="USD"),
    dict(invoice="880000000031", amount=800, paid_at="2026-10-01 09:00:00", ref="S00031"),
]


def read(content, name="Account_history_test.xlsx"):
    return e.read_bank_file(name, content, "BOA")


def run_september(tmp_path, rows=ROWS, pays=PAYS):
    files = [read(statement(rows)), e.read_payment_file("Payment details.csv", payments(pays), "BOA")]
    res = e.run(tmp_path / "base", SEP, files, require_review=False, bank="BOA", report_dir=tmp_path / "r")
    return res, res.sheets["Payment details"].set_index("Ref no", drop=False)


# ---- reading the statement ------------------------------------------------------------------------------------------
def test_statement_is_read_and_checked_against_its_own_totals():
    f = read(statement(ROWS))
    assert (f.currency, f.bank) == ("RWF", "BOA")
    assert (f.first, f.last) == (pd.Timestamp("2026-08-31"), pd.Timestamp("2026-10-01"))   # oldest day left out
    # only the Irembo payment credits are kept (like I&M); the other credits are counted: refund, salary, reversed
    assert f.rows == 16 and f.other_credits == 3
    assert "Debit" not in f.data.columns and (f.data["Credit"].astype(float) > 0).all()
    assert {"S00010", "S00012", "S00015"}.isdisjoint(f.data["Ref no"])         # reversed, refund, other credit
    assert {"Ledger Date", "Value Date", "Branch", "Ref no", "Credit", "Balance"} <= set(f.data.columns)


def test_header_found_anywhere_dates_as_cells_repeated_header_and_extra_rows():
    plain = read(statement(ROWS)).data
    other = read(statement(ROWS, blank_lines_before_header=4, date_cells=True, repeat_header_at=5,
                           extra_rows_after=[["printed by test"], ["Page 1 of 1"]])).data
    a, b = boa.clean_statement(plain), boa.clean_statement(other)
    for col in ("Value Date", "Ref no", "Description", "Credit", "Balance", "invoice_number"):
        assert a[col].equals(b[col]), col


@pytest.mark.parametrize("change, message", [
    (dict(count=99), "statement says 99 transactions"),
    (dict(total_credit=1), "Total Credit"),
    (dict(total_debit=-1), "Total Debit"),
    (dict(closing=1), "not the Closing"),
])
def test_a_statement_that_does_not_add_up_is_refused(change, message):
    with pytest.raises(e.ReconError, match=message):
        read(statement(ROWS, **change))


def test_a_broken_balance_is_refused(tmp_path):
    import io
    from openpyxl import load_workbook
    wb = load_workbook(io.BytesIO(statement(ROWS)))
    ws = wb.active
    row = next(r for r in ws.iter_rows() if r[0].value == 3)
    row[8].value = row[8].value + 1                             # one balance edited by hand
    buf = io.BytesIO()
    wb.save(buf)
    with pytest.raises(e.ReconError, match="Balance column does not follow"):
        read(buf.getvalue())


@pytest.mark.parametrize("content, message", [
    (statement(ROWS, currency="USD"), "USD account"),
    (statement(ROWS, drop_column="Value Date"), "not a Bank of Africa account history"),
    (statement([("31/02/2026", "X1", "ePay irembo 880000000001", None, 5)]), "cannot be read as transactions"),
    (b"not an excel file", "cannot be opened"),
], ids=["usd-account", "missing-value-date", "invalid-date", "not-excel"])
def test_wrong_or_malformed_files_are_refused(content, message):
    with pytest.raises(e.ReconError, match=message):
        read(content)


def test_an_im_statement_is_not_taken_for_boa():
    from openpyxl import Workbook
    import io
    wb = Workbook()
    wb.active.append(["Transaction Date", "Value Date", "Description", "Tran Id", "Credit", "Balance (RWF)"])
    buf = io.BytesIO()
    wb.save(buf)
    with pytest.raises(e.ReconError, match="not a Bank of Africa"):
        read(buf.getvalue())


# ---- credits only ----------------------------------------------------------------------------------------------------
def test_only_irembo_credits_are_kept_and_debits_never_are():
    s = boa.clean_statement(read(statement(ROWS)).data).set_index("Ref no")
    assert {"S00001", "S00002", "S00003", "S00004", "S00018"} <= set(s.index)  # IB, USSD, cash deposit, permit
    assert {"S00011", "S00013", "S00014"}.isdisjoint(s.index)                  # debits (even naming Irembo)
    assert "S00010" not in s.index                                             # credit cancelled in the file
    assert s.loc["S00018", "invoice_number"] == "012345678901"                 # leading zero kept
    assert pd.isna(s.loc["S00004", "invoice_number"])                          # permit: no invoice in the text


# ---- matching and the report -------------------------------------------------------------------------------------
def test_matching_september(tmp_path):
    res, d = run_september(tmp_path)
    assert res.bank_rows == 13                     # Irembo payments of September (reversed, October, August out)
    assert res.not_found == 3 and set(d[~d.found_in_irembopay]["Ref no"]) == {"S00006", "S00016", "S00019"}
    assert d.loc["S00004", "invoice_number"] == "880000000004"                 # permit matched by Ref no
    assert d.loc["S00004", "Ref check"].startswith("matched by Ref no")
    assert "ambiguous" in d.loc["S00016", "Ref check"]                         # never chosen arbitrarily
    assert pd.isna(d.loc["S00019", "invoice_number"])                         # never matched on amount alone
    assert d.loc["S00018", "found_in_irembopay"]                               # leading-zero invoice matched
    # the double credit: both rows on the statement, same invoice; the earlier one (S00008) is a re-send and is
    # excluded, so only S00009 is left (still flagged: IremboPay recorded the credit against S00008's reference)
    assert "S00008" not in d.index
    assert d.loc["S00009", "Ref check"] == "IremboPay recorded another credit: Ref S00008"
    assert not bool(d.loc["S00009", "is_extra_credit"]) and int(d.loc["S00009", "times_on_statement"]) == 1
    assert "Amount difference" in d.loc["S00007", "Reason"]
    assert "Currency difference" in d.loc["S00030", "Reason"]
    assert "Not settled (FAILED)" in d.loc["S00017", "Reason"]                 # an unknown status is not success
    assert "Other channel (MTN_MOMO)" in d.loc["S00005", "Reason"]
    t = e.counts_of(res.blocks[-1][2])
    assert t["Overall processed payments"]["Total"] == 13
    assert t["Overall notifications received by Irembo"]["Total"] == 7         # 001 002 004 007 009 018 030
    assert t["Pending notifications"]["Total"] == 1                            # 003 (NEW)
    assert t["Duplicate Payment"]["Total"] == 1                                # 005 (another channel)
    # SLA as for I&M: paid_at day vs Value Date, 0 days allowed -> 002 and the surviving row of the double credit
    assert t["SLA Violated"]["Total"] == 2
    assert d.loc["S00001", "sla_status"] == "On time" and d.loc["S00009", "sla_status"] == "SLA violated"
    assert d.loc["S00001", "Branch"] == "HEAD OFFICE"


def test_report_holds_credits_only(tmp_path):
    res, _ = run_september(tmp_path)
    assert set(res.sheets) == {"Pending notification", "Duplicate payment other channel", "BOA statement",
                               "Payment details"}                               # the same sheets as I&M
    assert (res.sheets["BOA statement"]["Credit"] > 0).all()


def test_period_uses_the_value_date_and_files_cover_the_month():
    files = [read(statement(ROWS)), e.read_payment_file("p.csv", payments(PAYS), "BOA")]
    chk = e.check_period(SEP, files, "BOA", today="2026-10-11")
    assert chk.blocking == [] and any(t.startswith("Bank RWF: 01/09 to 30/09") for ok, t in chk.lines if ok)
    aug = e.check_period(e.make_period("day", "2026-08-30"), files, "BOA", today="2026-10-11")
    assert any("missing" in t for ok, t in aug.lines if not ok)                # the oldest day is not covered


def test_wrong_bank_payment_file_is_refused():
    other = e.read_payment_file("p.csv", payments([dict(PAYS[0], provider="I_M")]), "BOA")
    assert "has no Bank of Africa payments" in e.accept_file(SEP, other, [], "BOA")


def test_report_files_are_written(tmp_path):
    res, _ = run_september(tmp_path)
    from openpyxl import load_workbook
    sheets = load_workbook(res.report_file, read_only=True).sheetnames
    assert sheets[:2] == ["Summary-Weekly", "Summary-Monthly"]
    assert {"BOA statement", "Payment details"} <= set(sheets) and "Excluded bank rows" not in sheets
    assert e.report_pdf(res.report_file)[:4] == b"%PDF"
    assert e.read_report(res.report_file)["bank"] == "Bank of Africa"


# ---- the app ---------------------------------------------------------------------------------------------------------
def test_app_offers_boa_and_checks_its_files():
    from streamlit.testing.v1 import AppTest
    at = AppTest.from_file(str(APP_DIR / "app.py"), default_timeout=120)
    at.session_state["bank_code"] = "BOA"
    at.session_state["step"] = 2
    at.session_state["report_label"] = "Month"
    at.session_state["picked"] = SEP.start.date()
    files = [read(statement(ROWS)), e.read_payment_file("p.csv", payments(PAYS), "BOA")]
    at.session_state["files"] = {f.digest: f for f in files}
    at.session_state["goto"] = "reconciliation"             # the app opens on the Dashboard
    at.run()
    assert not at.exception, [x.value for x in at.exception]
    text = " ".join(str(m.value) for m in at.markdown)
    assert "Bank of Africa Rwanda Reconciliation" in text and "Bank RWF: 01/09 to 30/09 covered" in text


# ---- the real samples: BANK_RECON_SAMPLES / "Bank of Africe" / Account_history_2026-10-08.xlsx + Payment details.csv
REAL = (SAMPLES / "Bank of Africe") if SAMPLES else None


@pytest.mark.skipif(not REAL or not (REAL / "Account_history_2026-10-08.xlsx").exists(),
                    reason="BANK_RECON_SAMPLES not set, or missing the BOA sample")
def test_real_boa_samples(tmp_path):
    st = read((REAL / "Account_history_2026-10-08.xlsx").read_bytes(), "Account_history_2026-10-08.xlsx")
    pay = e.read_payment_file("Payment details.csv", (REAL / "Payment details.csv").read_bytes(), "BOA")
    # 429 rows checked against the totals; 370 Irembo credits kept; other credits: 2 refunds + 2 reversed credits
    assert (st.rows, st.other_credits, st.currency, pay.rows) == (370, 4, "RWF", 366)
    rows = matched = not_found = 0
    for m in range(2, 10):                                      # February to September 2026
        period = e.make_period("month", f"2026-{m:02d}-01")
        res = e.run(tmp_path / "b", period, [st, pay], require_review=False, bank="BOA", report_dir=tmp_path / "r")
        rows, matched, not_found = rows + res.bank_rows, matched + res.matched, not_found + res.not_found
    assert (rows, matched, not_found) == (345, 345, 0)


# ---- the other banks are unchanged (real samples: BANK_RECON_SAMPLES / "I&M Bank" / "2026-08" / "statements") ------
DL = SAMPLES


@pytest.mark.skipif(not DL or not (DL / "I&M Bank" / "2026-08" / "statements").exists(),
                    reason="BANK_RECON_SAMPLES not set, or missing the I&M sample")
def test_im_unchanged(tmp_path):
    stm = sorted((DL / "I&M Bank" / "2026-08" / "statements").glob("*.xls"))
    pay = next(DL.glob("_SELECT_pi_invoice_number*.csv"))
    files = [e.read_bank_file(p.name, p.read_bytes(), "IM") for p in stm] + \
        [e.read_payment_file(pay.name, pay.read_bytes(), "IM")]
    res = e.run(tmp_path, e.make_period("month", "2026-08-01"), files, require_review=False, bank="IM",
                report_dir=tmp_path / "r")
    assert (res.bank_rows, res.not_found) == (5419, 10)
