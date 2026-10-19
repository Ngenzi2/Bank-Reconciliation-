"""Check-screen warnings (dianah branch, in plain words): a statement currency that does not match IremboPay for
the same invoice, and an IremboPay export taken before the newest statement."""
from dataclasses import replace

import pandas as pd

import engine as e
from boa_fixtures import payments, statement

SEP = e.make_period("month", "2026-09-01")
ROWS = [("05/09/2026", "S00001", "ePay irembo 880000000001", None, 5000)]


def read(content, name="Account_history.xlsx"):
    return e.read_bank_file(name, content, "BOA")


def test_currency_mismatch_is_reported_by_file():
    stm = read(statement(ROWS))
    pay = e.read_payment_file("p.csv", payments(
        [dict(invoice="880000000001", amount=5000, paid_at="2026-09-05 09:00:00", ref="S00001", currency="USD")]
    ), "BOA")
    chk = e.check_period(SEP, [stm, pay], "BOA", today="2026-10-01")
    assert any("don't match IremboPay's currency" in w and "Account_history.xlsx" in w for w in chk.warnings)


def test_matching_currency_has_no_warning():
    stm = read(statement(ROWS))
    pay = e.read_payment_file("p.csv", payments(
        [dict(invoice="880000000001", amount=5000, paid_at="2026-09-05 09:00:00", ref="S00001")]
    ), "BOA")
    chk = e.check_period(SEP, [stm, pay], "BOA", today="2026-10-01")
    assert not any("currency" in w for w in chk.warnings)


def test_export_older_than_the_newest_statement_is_a_warning():
    stm = read(statement(ROWS))
    pay = e.read_payment_file("p.csv", payments(
        [dict(invoice="880000000001", amount=5000, paid_at="2026-09-05 09:00:00", ref="S00001")]
    ), "BOA")
    old_pay = replace(pay, taken_at=pd.Timestamp("2026-09-02 08:00:00"))   # taken before the 05/09 statement
    chk = e.check_period(SEP, [stm, old_pay], "BOA", today="2026-10-01")
    assert any("taken" in w and "before the newest statement" in w for w in chk.warnings)

    fresh_pay = replace(pay, taken_at=pd.Timestamp("2026-09-06 08:00:00"))
    chk2 = e.check_period(SEP, [stm, fresh_pay], "BOA", today="2026-10-01")
    assert not any("before the newest statement" in w for w in chk2.warnings)
