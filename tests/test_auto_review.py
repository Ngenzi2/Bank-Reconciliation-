"""Automatic classification of rows not found in IremboPay (dianah branch, gap #4/#6): no manual review for
normal cases."""
import numpy as np
import pandas as pd

import engine as e
from services import auto_review


def make_review(rows):
    """rows: [(invoice_number or None, reviewed_already)]."""
    base = {c: "" for c in e.EDIT_COLS}
    out = []
    for invoice, reviewed in rows:
        r = dict(base)
        r["invoice_number"] = invoice
        r["Reviewed (Yes/No)"] = reviewed
        out.append(r)
    df = pd.DataFrame(out)
    df["invoice_number"] = df["invoice_number"].replace({None: np.nan})
    return df


def test_each_checkout_answer_gets_its_own_classification():
    review = make_review([
        ("1", ""), ("2", ""), ("3", ""), ("4", ""), ("5", ""), (None, ""),
    ])
    answers = {"1": "Paid", "2": "Expired", "3": "Open, not paid yet", "4": "Not an IremboPay invoice"}
    # "5" has no answer at all (the checkout could not be asked)
    out, needs_a_look = auto_review.classify(review, answers)
    by_invoice = out.set_index(out["invoice_number"].astype(object), drop=False)

    assert list(out["Reviewed (Yes/No)"]) == ["Yes"] * 6                       # every row decided

    row = by_invoice.loc["1"]
    assert (row["Report table"], row["payment_status"]) == ("Paid with other channel", "")

    row = by_invoice.loc["2"]
    assert (row["Report table"], row["payment_status"]) == ("Pending notification", "EXPIRED")

    row = by_invoice.loc["3"]
    assert (row["Report table"], row["payment_status"]) == ("Pending notification", "NEW")

    row = by_invoice.loc["4"]                                                  # NOT FOUND counts as EXPIRED
    assert (row["Report table"], row["payment_status"]) == ("Pending notification", "EXPIRED")

    row = by_invoice.loc["5"]                                                  # could not check: still Pending
    assert (row["Report table"], row["payment_status"]) == ("Pending notification", e.COULD_NOT_CHECK_STATUS)
    assert needs_a_look == ["5"]

    no_invoice = out[out["invoice_number"].isna()].iloc[0]
    assert (no_invoice["Report table"], no_invoice["payment_status"]) == ("Pending notification", e.NO_INVOICE_STATUS)


def test_a_row_already_decided_is_left_as_it_is():
    review = make_review([("1", "")])
    review.loc[0, ["Reviewed (Yes/No)", "Report table", "payment_status"]] = ["Yes", "Notification received", "PAID"]
    out, needs_a_look = auto_review.classify(review, {"1": "Expired"})
    assert out.loc[0, "Report table"] == "Notification received"               # untouched, not re-decided
    assert needs_a_look == []


def test_processed_equals_notified_plus_pending_plus_duplicate_after_auto_classification(tmp_path):
    from boa_fixtures import payments, statement
    rows = [
        ("01/09/2026", "S1", "ePay irembo 880000000001", None, 1000),     # matched, notified
        ("02/09/2026", "S2", "ePay irembo 880000000002", None, 2000),     # checkout: Paid -> Duplicate
        ("03/09/2026", "S3", "ePay irembo 880000000003", None, 3000),     # checkout: Expired -> Pending
        ("04/09/2026", "S4", "ePay irembo 880000000004", None, 4000),     # checkout: NEW -> Pending
        ("05/09/2026", "S5", "ePay irembo 880000000005", None, 5000),     # checkout: NOT FOUND -> Pending (EXPIRED)
        ("06/09/2026", "S6", "ePay irembo 880000000006", None, 6000),     # could not check -> Pending
        ("07/09/2026", "S7", "irembo Invoice Permit", None, 7000),        # no invoice number -> Pending
    ]
    pays = [dict(invoice="880000000001", amount=1000, paid_at="2026-09-01 09:00:00", ref="S1")]
    files = [e.read_bank_file("s.xlsx", statement(rows), "BOA"),
            e.read_payment_file("p.csv", payments(pays), "BOA")]
    period = e.make_period("month", "2026-09-01")
    first = e.run(tmp_path / "base", period, files, require_review=True, bank="BOA", report_dir=tmp_path / "r1")
    answers = {"880000000002": "Paid", "880000000003": "Expired", "880000000004": "Open, not paid yet",
              "880000000005": "Not an IremboPay invoice"}        # 880000000006: no answer (could not check)
    filled, needs_a_look = auto_review.classify(first.review, answers)
    e.save_decisions(tmp_path / "base", period, filled, "BOA")
    res = e.run(tmp_path / "base", period, files, require_review=True, bank="BOA", report_dir=tmp_path / "r2")
    assert not res.stopped_for_review
    assert needs_a_look == ["880000000006"]

    monthly = next(t for sheet, _, t in res.blocks if sheet == "Summary-Monthly")

    def total(title):
        t = monthly[title]
        return int(t.loc[t["Currency"].isna(), "Number of Transactions"].iloc[0])

    processed, notified = total("Overall processed payments"), total("Overall notifications received by Irembo")
    pending, duplicate = total("Pending notifications"), total("Duplicate Payment")
    assert processed == 7
    assert notified == 1
    assert pending == 5                              # EXPIRED, NEW, NOT FOUND, could not check, no invoice number
    assert duplicate == 1
    assert processed == notified + pending + duplicate


def test_apply_review_accepts_the_two_automatic_statuses(tmp_path):
    from boa_fixtures import payments, statement
    rows = [("05/09/2026", "S1", "ePay irembo 880000000001", None, 1000),
            ("06/09/2026", "S2", "ePay irembo 880000000002", None, 2000)]
    filler = [dict(invoice="880000000099", amount=1, paid_at="2026-09-01 09:00:00", ref="X")]
    files = [e.read_bank_file("s.xlsx", statement(rows), "BOA"),
            e.read_payment_file("p.csv", payments(filler), "BOA")]
    period = e.make_period("month", "2026-09-01")
    first = e.run(tmp_path / "base", period, files, require_review=True, bank="BOA", report_dir=tmp_path / "r1")
    review = first.review.copy()
    answers = {"880000000001": "Open, not paid yet"}     # "880000000002": could not check (no answer)
    filled, needs_a_look = auto_review.classify(review, answers)
    e.save_decisions(tmp_path / "base", period, filled, "BOA")
    res = e.run(tmp_path / "base", period, files, require_review=True, bank="BOA", report_dir=tmp_path / "r2")
    assert not res.stopped_for_review
    assert needs_a_look == ["880000000002"]
    pending = res.sheets["Pending notification"].set_index(res.sheets["Pending notification"]["invoice_number"])
    assert pending.loc["880000000001", "payment_status"] == "NEW"
    assert pending.loc["880000000002", "payment_status"] == e.COULD_NOT_CHECK_STATUS
