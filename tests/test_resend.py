"""Common engine rules that do not depend on one bank:
- a re-send of an invoice already credited on an earlier day (read via a "reference" file, up to
  engine.RESEND_LOOKBACK_DAYS before the period) is excluded from Processed.
- when an invoice is in several IremboPay export files, the newest file wins (not the most "successful" row)."""
import csv
import io
from dataclasses import replace

import pandas as pd

import engine as e
from boa_fixtures import statement

SEP = e.make_period("month", "2026-09-01")

PAY_COLUMNS = ["invoice_number", "transaction_id", "amount", "currency", "payment_status", "description", "expiry_at",
               "created_at", "settled_at", "paid_at", "name", "transaction_reference", "provider_reference",
               "payment_channel", "payment_provider", "transaction_status"]


def payments_csv(rows):
    """rows: [dict(invoice, amount, paid_at, created_at, status, provider, transaction_status, ref)]."""
    out = io.StringIO()
    w = csv.writer(out)
    w.writerow(PAY_COLUMNS)
    for i, r in enumerate(rows):
        w.writerow([r["invoice"], f"T{i:06d}", f"{r['amount']:.2f}", r.get("currency", "RWF"),
                    r.get("status", "PAID"), "", "", r.get("created_at", r["paid_at"]), "", r["paid_at"],
                    "Test Merchant", "", r.get("ref", ""), r.get("channel", "BOA_IB"),
                    r.get("provider", "BANK_OF_AFRICA"), r.get("transaction_status", "SUCCESSFUL")])
    return out.getvalue().encode()


def read_bank(rows, name="Account_history.xlsx"):
    return e.read_bank_file(name, statement(rows), "BOA")


def test_a_resend_within_the_lookback_window_is_excluded_from_processed(tmp_path):
    # the invoice was already credited in August (read as a reference file: not required, not counted on its own)
    earlier = replace(read_bank(
        [("20/08/2026", "R00050", "ePay irembo 880000000050", None, 2000)]), reference=True)
    # September: the same invoice credited again (a re-send), plus a genuinely new invoice
    current = read_bank([
        ("02/09/2026", "R00051", "ePay irembo 880000000050", None, 2000),
        ("03/09/2026", "R00052", "ePay irembo 880000000052", None, 3000),
    ])
    # IremboPay's paid_at is inside September for both, so only the statement history tells the two cases apart
    pay = e.read_payment_file("Payment details.csv", payments_csv([
        dict(invoice="880000000050", amount=2000, paid_at="2026-09-02 09:00:00", ref="R00050"),
        dict(invoice="880000000052", amount=3000, paid_at="2026-09-03 09:00:00", ref="R00052"),
    ]), "BOA")

    res = e.run(tmp_path / "base", SEP, [earlier, current, pay], require_review=False, bank="BOA",
                report_dir=tmp_path / "r")
    assert res.bank_rows == 1                                     # only the new invoice is Processed
    assert res.by_currency["RWF"]["rows"] == 1

    # without the reference file, the engine has no way to know about the earlier credit: both are Processed
    res_no_ref = e.run(tmp_path / "base2", SEP, [current, pay], require_review=False, bank="BOA",
                       report_dir=tmp_path / "r2")
    assert res_no_ref.bank_rows == 2


def test_a_resend_is_still_excluded_even_when_irembopay_already_shows_it_paid_earlier(tmp_path):
    # no statement at all from before the period, but IremboPay already shows the invoice paid in August; a second,
    # genuinely new invoice keeps the period from being entirely empty
    current = read_bank([
        ("02/09/2026", "R00060", "ePay irembo 880000000060", None, 1000),
        ("03/09/2026", "R00061", "ePay irembo 880000000061", None, 1500),
    ])
    pay = e.read_payment_file("Payment details.csv", payments_csv([
        dict(invoice="880000000060", amount=1000, paid_at="2026-08-25 09:00:00", ref="R00060"),
        dict(invoice="880000000061", amount=1500, paid_at="2026-09-03 09:00:00", ref="R00061"),
    ]), "BOA")
    res = e.run(tmp_path / "base", SEP, [current, pay], require_review=False, bank="BOA", report_dir=tmp_path / "r")
    assert res.bank_rows == 1
    assert res.sheets["Payment details"]["invoice_number"].astype(str).tolist() == ["880000000061"]


def test_a_resend_counts_once_even_with_a_different_bank_reference(tmp_path):
    # a failed first send and a successful resend usually get different bank references; the invoice still
    # counts once in Processed, whatever the two references are (same day, within the period itself)
    current = read_bank([
        ("02/09/2026", "R00080", "ePay irembo 880000000080", None, 2000),   # the failed attempt, its own reference
        ("02/09/2026", "R00081", "ePay irembo 880000000080", None, 2000),   # the resend, a different reference
    ])
    pay = e.read_payment_file("Payment details.csv", payments_csv([
        dict(invoice="880000000080", amount=2000, paid_at="2026-09-02 09:05:00", ref="R00081"),
    ]), "BOA")
    res = e.run(tmp_path / "base", SEP, [current, pay], require_review=False, bank="BOA", report_dir=tmp_path / "r")
    assert res.bank_rows == 1
    assert res.sheets["Payment details"]["invoice_number"].astype(str).tolist() == ["880000000080"]


def test_check_period_ignores_reference_files_for_coverage_and_never_blocks_on_them(tmp_path):
    earlier = replace(read_bank([("20/08/2026", "R00050", "ePay irembo 880000000050", None, 2000)]), reference=True)
    assert e.accept_file(SEP, earlier, [], "BOA") is None          # never rejected for being outside the period
    chk = e.check_period(SEP, [earlier], "BOA", today=pd.Timestamp("2026-10-01"))
    assert chk.blocking                                             # the September statement is still missing
    assert not any("20/08" in text for _, text in chk.lines)        # the reference file is not part of the checklist


def test_newest_export_file_wins_an_invoice_not_the_most_successful_row(tmp_path):
    current = read_bank([("02/09/2026", "R00070", "ePay irembo 880000000070", None, 1000)])
    older_file = e.read_payment_file("export_1.csv", payments_csv([
        dict(invoice="880000000070", amount=1000, paid_at="2026-09-02 09:00:00", created_at="2026-09-05 10:00:00",
             ref="R00070", provider="BANK_OF_AFRICA", transaction_status="SUCCESSFUL"),
    ]), "BOA")
    newer_file = e.read_payment_file("export_2.csv", payments_csv([
        dict(invoice="880000000070", amount=1000, paid_at="2026-09-02 09:00:00", created_at="2026-09-01 08:00:00",
             ref="M1", provider="MTN_MOMO", transaction_status="SUCCESSFUL"),
    ]), "BOA")
    older_file = replace(older_file, order=1)
    newer_file = replace(newer_file, order=2)
    # the older file's row has the later created_at and is just as "successful": the old rule would have picked it
    res = e.run(tmp_path / "base", SEP, [current, older_file, newer_file], require_review=False, bank="BOA",
               report_dir=tmp_path / "r")
    details = res.sheets["Payment details"].set_index(res.sheets["Payment details"]["invoice_number"].astype(str))
    assert details.loc["880000000070", "payment_provider"] == "MTN_MOMO"
