"""Review rule (2026-10-20, replaces 2026-10-12): "expired" is not a payment status. An invoice the IremboPay check
found expired was never paid: it is put in "Pending notification" with payment_status NEW, and the expiry is a note
(review status "Pending notification (invoice expired)"). Every bank.
Uses the real August 2026 samples when they are on this computer (skipped otherwise)."""
import glob
import os
from pathlib import Path

import pandas as pd
import pytest

import engine as e
from banks import ecobank

D = Path(os.path.expanduser("~/Downloads"))
SAMPLES = {
    "ECO": (sorted(glob.glob(str(D / "data/August 2026 21-35-06-731/IREMBO PAY*.xlsx"))),
            D / "_SELECT_pi_invoice_number_pi_transaction_id_pi_amount_pi_currenc_202609291459.csv"),
    "IM": (sorted(glob.glob(str(D / "I&M Bank/2026-08/statements/*.xls"))),
           D / "_SELECT_pi_invoice_number_pi_transaction_id_pi_amount_pi_currenc_202610021501.csv"),
}


def test_pending_is_in_irembopay_without_paid_at():
    # 26/10/2026: paid = paid_at filled, whatever the status; pending = in IremboPay without paid_at
    rows = pd.DataFrame({"found_in_irembopay": [True, True, True, False],
                         "payment_status": ["EXPIRED", "NEW", "PAID", None],
                         "paid_at": pd.to_datetime([None, None, "2026-08-01", None]),
                         "payment_provider": ["ECOBANK"] * 3 + [None]})
    assert ecobank.pending(rows).tolist() == [True, True, False, False]
    assert ecobank.received(rows).tolist() == [False, False, True, False]


def run_with_decisions(tmp_path, bank, decide):
    stm, pay = SAMPLES[bank]
    if not stm or not pay.exists():
        pytest.skip("real samples not on this computer")
    files = [e.read_bank_file(Path(p).name, Path(p).read_bytes(), bank) for p in stm] + \
            [e.read_payment_file(pay.name, pay.read_bytes(), bank)]
    period = e.make_period("month", "2026-08-01")
    acc = []
    for f in files:
        if e.accept_file(period, f, acc, bank) is None:
            acc.append(f)
    base = tmp_path / "base"
    first = e.run(base, period, acc, require_review=True, bank=bank, report_dir=tmp_path / "r1")
    review = first.review.copy()
    for c in e.EDIT_COLS:
        review[c] = review[c].fillna("").astype(str)
    review = decide(review)
    e.save_decisions(base, period, review, bank)
    res = e.run(base, period, acc, require_review=True, bank=bank, report_dir=tmp_path / "r2")
    details = res.sheets["Payment details"]
    return review, res, details


def pending_total(res):
    main = next(b for b in res.blocks if b[0] == "Summary-Monthly")[2]["Pending notifications"]
    return int(main[main["Currency"].isna()]["Number of Transactions"].iloc[0])


def test_ecobank_expired_goes_to_pending_as_new_with_a_note(tmp_path):
    def decide(review):
        review["Reviewed (Yes/No)"] = "Yes"
        review["Report table"] = "Paid with other channel"
        once = review["invoice_number"].astype(str).map(review["invoice_number"].astype(str).value_counts()).eq(1)
        picks = review.index[once][:2]
        review.loc[picks[0], ["Report table", "payment_status"]] = ["Pending notification", "EXPIRED"]
        review.loc[picks[1], ["Report table", "payment_status"]] = ["Pending notification", "NEW"]
        review.attrs["picks"] = list(picks)
        return review
    review, res, details = run_with_decisions(tmp_path, "ECO", decide)
    assert not res.stopped_for_review
    inv = review.loc[review.attrs["picks"], "invoice_number"].astype(str)
    by_inv = details.set_index(details["invoice_number"].astype(str))
    assert by_inv["payment_status"][inv.iloc[0]] == "NEW" and by_inv["payment_status"][inv.iloc[1]] == "NEW"
    assert by_inv["review_status"][inv.iloc[0]] == "Pending notification (invoice expired)"
    assert by_inv["review_status"][inv.iloc[1]] == "Pending notification"
    assert pending_total(res) >= 2                                   # both are counted as pending


def test_im_expired_goes_to_pending_as_new_with_a_note(tmp_path):
    """I&M (and Bank of Africa, BPR) as Ecobank: an expired invoice is pending, NEW, with the note."""
    def decide(review):
        review["Reviewed (Yes/No)"] = "Yes"
        review["Report table"] = "Paid with other channel"
        # two invoices that are on the statement once (not an extra credit of an invoice credited twice)
        once = review["invoice_number"].astype(str).map(review["invoice_number"].astype(str).value_counts()).eq(1)
        picks = review.index[once & ~review["Reason"].fillna("").str.startswith("Invoice credited again")][:2]
        review.loc[picks[0], ["Report table", "payment_status"]] = ["Pending notification", "EXPIRED"]
        review.loc[picks[1], ["Report table", "payment_status"]] = ["Pending notification", ""]
        review.attrs["picks"] = list(picks)
        return review
    review, res, details = run_with_decisions(tmp_path, "IM", decide)
    by_inv = details.set_index(details["invoice_number"].astype(str))
    inv = review.loc[review.attrs["picks"], "invoice_number"].astype(str)
    assert by_inv["payment_status"][inv.iloc[0]] == "NEW" and by_inv["payment_status"][inv.iloc[1]] == "NEW"
    assert by_inv["review_status"][inv.iloc[0]] == "Pending notification (invoice expired)"
    assert pending_total(res) == 2


def test_every_bank_paid_means_paid_at():
    from banks import boa, bpr, im
    for bank in (im, boa, ecobank, bpr):
        rows = pd.DataFrame({"found_in_irembopay": [True, True, True], "payment_status": ["FAILED", "NEW", "PAID"],
                             "paid_at": pd.to_datetime(["2026-08-01", None, "2026-08-02"]),
                             "payment_provider": [bank.PROVIDER, bank.PROVIDER, "OTHER"]})
        assert bank.PENDING_STATUSES == {"NEW"}
        assert bank.received(rows).tolist() == [True, False, False]           # FAILED with a paid_at: paid
        assert bank.pending(rows).tolist() == [False, True, False]
        assert bank.duplicate(rows).tolist() == [False, False, True]


def test_ecobank_other_channel_decision_counts_once_in_the_reconciled_month(tmp_path):
    """A row the review places in 'Paid with other channel' has no paid_at (not in the payment file): it is dated by
    its first Ecobank attempt INSIDE the reconciled month (11 August invoices were first tried on 31 July and
    retried in August: they count in August), counted once, and lands in the Duplicate Payment table. The weeks
    add up to the month (no invoice in two weeks). Every row gets its decision, also the 4 with no bank reference."""
    def decide(review):
        review["Reviewed (Yes/No)"] = "Yes"
        review["Report table"] = "Paid with other channel"
        return review
    review, res, details = run_with_decisions(tmp_path, "ECO", decide)
    assert not res.stopped_for_review and res.not_reviewed == 0
    month = next(b for b in res.blocks if b[0] == "Summary-Monthly")[2]["Duplicate Payment"]
    total = int(month[month["Currency"].isna()]["Number of Transactions"].iloc[0])
    # 26/10/2026: only the requests Ecobank answered SUCCESS are processed, each invoice once, by Request Date
    det = res.sheets["Payment details"]
    decided = det["review_status"].eq("Paid with other channel")
    counted = det[decided & det["Cba Rsp Code"].astype(str).eq("000") & ~det["is_extra_credit"].astype(bool)]
    assert len(counted) == counted["invoice_number"].nunique()
    # 26/10/2026: an extra debit put in "Paid with other channel" counts too (a double payment)
    extra = det[decided & det["is_extra_credit"].astype(bool)]
    assert total == len(counted) + len(extra)
    weeks = [b[2]["Duplicate Payment"] for b in res.blocks if b[0] == "Summary-Weekly"]
    assert sum(int(w[w["Currency"].isna()]["Number of Transactions"].iloc[0]) for w in weeks) == total
    sheet = res.sheets["Duplicate payment other channel"]
    assert len(sheet) == total
    assert set(sheet["invoice_number"]) == set(counted["invoice_number"]) | set(extra["invoice_number"])
    assert "Debited more than once" not in res.sheets                    # removed: each invoice is recorded once


def test_ecobank_review_ignores_the_cba_reference():
    """Ecobank: an invoice is recognised by currency | Request Date | invoice - the Cba Ref No is not in the IremboPay
    files and can be missing, so it never stops a decision from matching. I&M still uses its bank reference."""
    rows = pd.DataFrame({"Currency": ["RWF", "RWF"], "Request Date": pd.to_datetime(["2026-08-10", "2026-08-19"]),
                         "Cba Ref No": [None, "C1"], "invoice_number": ["886222926783", "886231956838"]})
    assert e.key_of(rows, "ECO").tolist() == ["RWF|2026-08-10||886222926783", "RWF|2026-08-19||886231956838"]
    im = pd.DataFrame({"Currency": ["RWF"], e.get_bank("IM").COLUMNS["payment_date"]: pd.to_datetime(["2026-08-10"]),
                       e.get_bank("IM").COLUMNS["reference"]: ["FT123"], "invoice_number": ["886222926783"]})
    assert e.key_of(im, "IM").tolist() == ["RWF|2026-08-10|FT123|886222926783"]


def test_ecobank_decisions_saved_with_a_cba_reference_still_match(tmp_path):
    """Decisions saved before this change carry the Cba Ref No in their key: they are still found."""
    def decide(review):
        review["Reviewed (Yes/No)"] = "Yes"
        review["Report table"] = "Paid with other channel"
        return review
    review, res, details = run_with_decisions(tmp_path, "ECO", decide)
    store_file = e.store_path(tmp_path / "base")
    store = pd.read_csv(store_file, dtype=str).fillna("")
    parts = store["key"].str.split("|", n=3, expand=True)          # an Ecobank invoice may itself contain "|"
    store["key"] = parts[0] + "|" + parts[1] + "|OLDREF" + store.index.astype(str) + "|" + parts[3]   # the old format
    store.to_csv(store_file, index=False)
    period = e.make_period("month", "2026-08-01")
    stm, pay = SAMPLES["ECO"]
    files = [e.read_bank_file(Path(p).name, Path(p).read_bytes(), "ECO") for p in stm] + \
            [e.read_payment_file(pay.name, pay.read_bytes(), "ECO")]
    acc = []
    for f in files:
        if e.accept_file(period, f, acc, "ECO") is None:
            acc.append(f)
    again = e.run(tmp_path / "base", period, acc, require_review=True, bank="ECO", report_dir=tmp_path / "r3")
    assert again.not_reviewed == 0 and not again.stopped_for_review


def test_an_excluded_row_stays_in_the_statement_with_a_comment(tmp_path):
    """26/10/2026: "Exclude" no longer removes the row: it stays in the statement and Payment details with its
    status and comment, and is counted in no table (the figures are those of before)."""
    def decide(review):
        review["Reviewed (Yes/No)"] = "Yes"
        review["Report table"] = "Exclude"
        review["Comment"] = "checked with the bank"
        return review
    review, res, details = run_with_decisions(tmp_path, "IM", decide)
    excluded = details[details["review_status"].eq(e.EXCLUDED)]
    assert len(excluded) == len(review) > 0
    assert excluded["review_comment"].eq("checked with the bank").all()
    stmt = res.sheets[e.get_bank("IM").STATEMENT_SHEET]
    assert len(stmt) == len(details)                                   # nothing removed from the statement
    month = next(b for b in res.blocks if b[0] == "Summary-Monthly")[2]
    total = lambda t: int(month[t][month[t]["Currency"].isna()]["Number of Transactions"].iloc[0])
    kept = details[~details["review_status"].eq(e.EXCLUDED) & ~details["is_extra_credit"].astype(bool)]
    assert total("Overall processed payments") == len(kept)           # the excluded rows are in no table
    for sheet in ("Pending notification", "Duplicate payment other channel"):
        assert not res.sheets[sheet]["review_status"].eq(e.EXCLUDED).any()
