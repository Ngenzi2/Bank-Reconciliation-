"""'VS previous period': N/A instead of a number when the IremboPay export does not go back far enough to cover
the period before (the same check the reconciliation notebooks make before showing a week-over-week or
month-over-month change) - approved 2026-10-21. When the export does reach back that far, a missing previous
report is still an error, exactly as before: the comparison would be trustworthy, it just hasn't been saved."""
import pytest

import engine as e
from bpr_fixtures import payments, statement

WEEK = e.make_period("week", "2026-08-03")   # first week of August 2026; the week before is 27/07 - 02/08/2026


def run(tmp_path, earliest_export_date):
    """One bank row (03/08/2026, matched, inside the week) + an unrelated IremboPay row only there to set how
    far back the export goes (earliest_export_date), so it never affects what is matched or resent."""
    bank = [("03/08/2026", "R1", "886000000001", None, "2026-08-03 09:00:00", None, 1000)]
    statements = [e.read_bank_file("statement.xlsx", statement(bank), "BPR")]
    pay_rows = [dict(invoice="886000000001", amount=1000, paid_at="2026-08-03 09:05:00", ref="R1"),
               dict(invoice="886999999999", amount=1, paid_at=f"{earliest_export_date} 00:00:00", ref="R0")]
    pay = [e.read_payment_file("payments.csv", payments(pay_rows), "BPR")]
    return e.run(tmp_path / "base", WEEK, statements + pay, compare=True, require_review=False, bank="BPR",
                report_dir=tmp_path / "r")


def test_na_when_the_export_does_not_reach_the_previous_week(tmp_path):
    # the export's earliest paid_at (01/08) is after the previous week's start (27/07): too thin to trust
    res = run(tmp_path, "2026-08-01")
    assert res.previous_counts is None
    assert "not compared" in res.compared_with
    main = next(t for sheet, _, t in res.blocks if sheet == "Summary-Weekly")
    col = main["Overall processed payments"]["VS prev Week (%)"]
    assert (col == "N/A").all() and not col.isna().any()


def test_still_needs_a_previous_report_when_the_export_does_reach_back(tmp_path):
    # the export's earliest paid_at (20/07) is before the previous week even starts: the data is there, so a
    # missing previous report is still a real problem, not silently N/A
    with pytest.raises(e.ReconError, match="No BPR report found"):
        run(tmp_path, "2026-07-20")
