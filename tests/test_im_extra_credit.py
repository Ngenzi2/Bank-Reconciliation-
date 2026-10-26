"""I&M (20/10/2026): an invoice credited more than once counts once (the earliest credit); every other credit of
that invoice goes to the step-5 table "Invoice credited more than once"."""
import pandas as pd

from banks import im


def test_the_earliest_credit_counts_and_the_others_are_reviewed():
    rows = pd.DataFrame({
        "invoice_number": ["886281187773", "886281187773", "886000000001", "886000000002", None],
        "found_in_irembopay": [True, True, True, False, False],
        "Value Date": pd.to_datetime(["2026-10-09", "2026-10-09", "2026-10-09", "2026-10-09", "2026-10-09"]),
        "Tran Id": ["S2", "S1", "S3", "S4", "S5"],
    })
    reasons = im.review_reasons(rows)
    assert reasons.tolist() == ["Invoice credited again: the credit counted is S1", "", "", "", ""]
    assert im.SEPARATE_CREDITED_AGAIN
