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
    # September: the same invoice credited again, same bank reference (a true re-send), plus a genuinely new invoice
    current = read_bank([
        ("02/09/2026", "R00050", "ePay irembo 880000000050", None, 2000),
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


def test_a_second_credit_with_a_different_reference_is_a_bank_duplicate(tmp_path):
    # two genuine SUCCESS credits for the same invoice, different bank references, same day: an extra credit, not
    # a re-send - both count in Processed, the second one goes to Duplicate Payment as "Bank duplicate"
    current = read_bank([
        ("02/09/2026", "R00080", "ePay irembo 880000000080", None, 2000),   # the first credit, its own reference
        ("02/09/2026", "R00081", "ePay irembo 880000000080", None, 2000),   # a second credit, a different reference
    ])
    pay = e.read_payment_file("Payment details.csv", payments_csv([
        dict(invoice="880000000080", amount=2000, paid_at="2026-09-02 09:05:00", ref="R00080"),
    ]), "BOA")
    res = e.run(tmp_path / "base", SEP, [current, pay], require_review=False, bank="BOA", report_dir=tmp_path / "r")
    assert res.bank_rows == 2
    d = res.sheets["Payment details"].set_index("Ref no", drop=False)
    # IremboPay recorded R00080 (the payment's own ref=): that credit is the invoice's own notification, whatever
    # order the statement lists them in; R00081 is the extra credit
    assert not bool(d.loc["R00080", "is_extra_credit"])
    assert bool(d.loc["R00081", "is_extra_credit"]) and d.loc["R00081", "Reason"] == "Bank duplicate"
    t = e.counts_of(res.blocks[-1][2])
    assert t["Overall processed payments"]["Total"] == 2
    assert t["Overall notifications received by Irembo"]["Total"] == 1          # the other credit is Notified
    assert t["Duplicate Payment"]["Total"] == 1


def test_a_bank_duplicate_is_never_counted_in_sla(tmp_path):
    # the extra credit itself would be "SLA violated" (paid 3 days after its own Value Date) if it were counted at
    # all; it must not be, since it is never Notified (SLA is counted among Notified payments, from the first
    # credit, only)
    current = read_bank([
        ("02/09/2026", "R00085", "ePay irembo 880000000085", None, 1000),   # the first credit: on time
        ("05/09/2026", "R00086", "ePay irembo 880000000085", None, 1000),   # the extra credit: 3 days later
    ])
    pay = e.read_payment_file("Payment details.csv", payments_csv([
        dict(invoice="880000000085", amount=1000, paid_at="2026-09-02 09:05:00", ref="R00085"),
    ]), "BOA")
    res = e.run(tmp_path / "base", SEP, [current, pay], require_review=False, bank="BOA", report_dir=tmp_path / "r")
    d = res.sheets["Payment details"].set_index("Ref no", drop=False)
    assert bool(d.loc["R00086", "is_extra_credit"]) and d.loc["R00086", "sla_status"] == "SLA violated"
    t = e.counts_of(res.blocks[-1][2])
    assert t["Overall processed payments"]["Total"] == 2
    assert t["SLA Violated"]["Total"] == 0           # the extra credit's violation is not counted
    assert t["Duplicate Payment"]["Total"] == 1


def test_a_second_credit_with_the_same_reference_is_still_a_resend(tmp_path):
    # the same bank reference seen twice (e.g. the export listing the same credit on two consecutive days) is
    # still a plain re-send, ignored - not a bank duplicate, even though it is not the very first occurrence
    current = read_bank([
        ("02/09/2026", "R00090", "ePay irembo 880000000090", None, 1000),
        ("03/09/2026", "R00090", "ePay irembo 880000000090", None, 1000),   # the same reference again
    ])
    pay = e.read_payment_file("Payment details.csv", payments_csv([
        dict(invoice="880000000090", amount=1000, paid_at="2026-09-02 09:05:00", ref="R00090"),
    ]), "BOA")
    res = e.run(tmp_path / "base", SEP, [current, pay], require_review=False, bank="BOA", report_dir=tmp_path / "r")
    assert res.bank_rows == 1
    assert res.sheets["Payment details"]["invoice_number"].astype(str).tolist() == ["880000000090"]


def test_irembopay_recorded_the_later_credit(tmp_path):
    # two genuine SUCCESS credits, different references; IremboPay's own provider_reference is the SECOND one
    # (R00111) - that credit is the payment even though it came later, not the earliest one
    current = read_bank([
        ("02/09/2026", "R00110", "ePay irembo 880000000110", None, 1000),    # first on the statement
        ("04/09/2026", "R00111", "ePay irembo 880000000110", None, 1000),    # later, but IremboPay's own ref
    ])
    pay = e.read_payment_file("Payment details.csv", payments_csv([
        dict(invoice="880000000110", amount=1000, paid_at="2026-09-04 09:05:00", ref="R00111"),
    ]), "BOA")
    res = e.run(tmp_path / "base", SEP, [current, pay], require_review=False, bank="BOA", report_dir=tmp_path / "r")
    d = res.sheets["Payment details"].set_index("Ref no", drop=False)
    assert not bool(d.loc["R00111", "is_extra_credit"])
    assert bool(d.loc["R00110", "is_extra_credit"]) and d.loc["R00110", "Reason"] == "Bank duplicate"
    t = e.counts_of(res.blocks[-1][2])
    assert t["Overall processed payments"]["Total"] == 2
    assert t["Overall notifications received by Irembo"]["Total"] == 1
    assert t["Duplicate Payment"]["Total"] == 1


def test_irembopay_holds_neither_reference_falls_back_to_the_first_credit(tmp_path):
    # two genuine SUCCESS credits on the statement, but IremboPay's provider_reference matches neither of them
    # (e.g. it was paid through yet another channel and later matched by invoice number only): the first credit
    # by date is the payment, same as before this rule existed
    current = read_bank([
        ("02/09/2026", "R00120", "ePay irembo 880000000120", None, 1000),
        ("04/09/2026", "R00121", "ePay irembo 880000000120", None, 1000),
    ])
    pay = e.read_payment_file("Payment details.csv", payments_csv([
        dict(invoice="880000000120", amount=1000, paid_at="2026-09-02 09:05:00", ref="SOME_OTHER_REF"),
    ]), "BOA")
    res = e.run(tmp_path / "base", SEP, [current, pay], require_review=False, bank="BOA", report_dir=tmp_path / "r")
    d = res.sheets["Payment details"].set_index("Ref no", drop=False)
    assert not bool(d.loc["R00120", "is_extra_credit"])        # the earliest credit, since neither ref matches
    assert bool(d.loc["R00121", "is_extra_credit"]) and d.loc["R00121", "Reason"] == "Bank duplicate"
    t = e.counts_of(res.blocks[-1][2])
    assert t["Overall processed payments"]["Total"] == 2
    assert t["Duplicate Payment"]["Total"] == 1


def test_a_resend_of_the_recorded_payment_is_still_ignored(tmp_path):
    # three credits: an earlier one IremboPay does not recognise, the recorded payment, and a re-send of that
    # recorded payment (same reference) - only the recorded one and the unrelated extra credit count
    current = read_bank([
        ("01/09/2026", "R00130", "ePay irembo 880000000130", None, 1000),    # not IremboPay's recorded ref
        ("03/09/2026", "R00131", "ePay irembo 880000000130", None, 1000),    # the recorded payment
        ("05/09/2026", "R00131", "ePay irembo 880000000130", None, 1000),    # a re-send of the recorded payment
    ])
    pay = e.read_payment_file("Payment details.csv", payments_csv([
        dict(invoice="880000000130", amount=1000, paid_at="2026-09-03 09:05:00", ref="R00131"),
    ]), "BOA")
    res = e.run(tmp_path / "base", SEP, [current, pay], require_review=False, bank="BOA", report_dir=tmp_path / "r")
    assert res.bank_rows == 2                                   # the re-send (second R00131) is not counted at all
    d = res.sheets["Payment details"].set_index("Ref no", drop=False)
    assert set(d.index) == {"R00130", "R00131"}
    assert not bool(d.loc["R00131", "is_extra_credit"])
    assert bool(d.loc["R00130", "is_extra_credit"]) and d.loc["R00130", "Reason"] == "Bank duplicate"


def test_a_bank_duplicate_belongs_to_the_week_of_the_second_credit(tmp_path):
    # the first credit is in week 1, the second (different reference) crosses into week 2: it must be counted
    # (Processed and Duplicate) in week 2, not week 1
    current = read_bank([
        ("01/09/2026", "R00100", "ePay irembo 880000000100", None, 1000),    # Tuesday, week 1
        ("08/09/2026", "R00101", "ePay irembo 880000000100", None, 1000),    # the following Tuesday, week 2
    ])
    pay = e.read_payment_file("Payment details.csv", payments_csv([
        dict(invoice="880000000100", amount=1000, paid_at="2026-09-01 09:05:00", ref="R00100"),
    ]), "BOA")
    res = e.run(tmp_path / "base", SEP, [current, pay], require_review=False, bank="BOA", report_dir=tmp_path / "r")
    assert res.bank_rows == 2
    week1 = next(t for sheet, heading, t in res.blocks if sheet == "Summary-Weekly" and heading.startswith("Week 1"))
    week2 = next(t for sheet, heading, t in res.blocks if sheet == "Summary-Weekly" and heading.startswith("Week 2"))

    def total(block, title):
        t = block[title]
        return int(t.loc[t["Currency"].isna(), "Number of Transactions"].iloc[0])

    assert total(week1, "Overall processed payments") == 1 and total(week1, "Duplicate Payment") == 0
    assert total(week2, "Overall processed payments") == 1 and total(week2, "Duplicate Payment") == 1


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
