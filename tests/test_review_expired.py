"""Review rule (approved 2026-10-12): an invoice the IremboPay check found EXPIRED is put in "Pending notification"
and keeps EXPIRED as its status in the report - Ecobank first, then I&M and Bank of Africa too.
Uses the real August 2026 samples when BANK_RECON_SAMPLES points to them (skipped otherwise)."""
import glob
import os
from pathlib import Path

import pandas as pd
import pytest

import engine as e
from banks import ecobank

_ROOT = os.environ.get("BANK_RECON_SAMPLES")
D = Path(_ROOT) if _ROOT else None
SAMPLES = {
    "ECO": (sorted(glob.glob(str(D / "data/August 2026 21-35-06-731/IREMBO PAY*.xlsx"))) if D else [],
            (D / "_SELECT_pi_invoice_number_pi_transaction_id_pi_amount_pi_currenc_202609291459.csv") if D else None),
    "IM": (sorted(glob.glob(str(D / "I&M Bank/2026-08/statements/*.xls"))) if D else [],
           (D / "_SELECT_pi_invoice_number_pi_transaction_id_pi_amount_pi_currenc_202610021501.csv") if D else None),
}


def test_ecobank_counts_expired_as_pending_not_received():
    rows = pd.DataFrame({"payment_status": ["EXPIRED", "NEW", "PAID"]})
    assert ecobank.pending(rows).tolist() == [True, True, False]


def run_with_decisions(tmp_path, bank, decide):
    stm, pay = SAMPLES[bank]
    if not stm or pay is None or not pay.exists():
        pytest.skip("BANK_RECON_SAMPLES not set, or missing the sample")
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


def test_ecobank_expired_goes_to_pending_and_keeps_expired(tmp_path):
    def decide(review):
        review["Reviewed (Yes/No)"] = "Yes"
        review["Report table"] = "Paid with other channel"
        review.loc[review.index[0], ["Report table", "payment_status"]] = ["Pending notification", "EXPIRED"]
        review.loc[review.index[1], ["Report table", "payment_status"]] = ["Pending notification", "NEW"]
        return review
    review, res, details = run_with_decisions(tmp_path, "ECO", decide)
    assert not res.stopped_for_review
    inv = review["invoice_number"].astype(str)
    status = details.set_index(details["invoice_number"].astype(str))["payment_status"]
    assert status[inv.iloc[0]] == "EXPIRED" and status[inv.iloc[1]] == "NEW"
    assert pending_total(res) >= 2                                   # both are counted as pending


def test_the_pending_list_shows_expired_as_new_with_a_comment(tmp_path):
    """The Pending notification sheet (the list the team reads) shows "NEW" for an expired invoice, with "Expired"
    in a Comment column - Payment details (the full technical record) still shows EXPIRED. The Pending count is
    the same either way (approved 2026-10-21)."""
    def decide(review):
        review["Reviewed (Yes/No)"] = "Yes"
        review["Report table"] = "Paid with other channel"
        review.loc[review.index[0], ["Report table", "payment_status"]] = ["Pending notification", "EXPIRED"]
        review.loc[review.index[1], ["Report table", "payment_status"]] = ["Pending notification", "NEW"]
        return review
    review, res, details = run_with_decisions(tmp_path, "ECO", decide)
    inv = review["invoice_number"].astype(str)
    pending = res.sheets["Pending notification"].set_index(
        res.sheets["Pending notification"]["invoice_number"].astype(str))
    assert pending.loc[inv.iloc[0], ["payment_status", "Comment"]].tolist() == ["NEW", "Expired"]
    assert pending.loc[inv.iloc[1], ["payment_status", "Comment"]].tolist() == ["NEW", ""]
    details_status = details.set_index(details["invoice_number"].astype(str))["payment_status"]
    assert details_status[inv.iloc[0]] == "EXPIRED"                       # Payment details is unchanged
    assert pending_total(res) == len(pending)                             # the count did not move


def test_im_expired_goes_to_pending_and_keeps_expired(tmp_path):
    """Since 2026-10-12 I&M (and Bank of Africa) follow the same rule as Ecobank: an EXPIRED invoice is pending and
    keeps EXPIRED in the report; a pending decision without a status is still NEW."""
    def decide(review):
        review["Reviewed (Yes/No)"] = "Yes"
        review["Report table"] = "Paid with other channel"
        review.loc[review.index[0], ["Report table", "payment_status"]] = ["Pending notification", "EXPIRED"]
        review.loc[review.index[1], ["Report table", "payment_status"]] = ["Pending notification", ""]
        return review
    review, res, details = run_with_decisions(tmp_path, "IM", decide)
    status = details.set_index(details["invoice_number"].astype(str))["payment_status"]
    inv = review["invoice_number"].astype(str)
    assert status[inv.iloc[0]] == "EXPIRED" and status[inv.iloc[1]] == "NEW"
    assert pending_total(res) == 2


def test_every_bank_counts_expired_as_pending():
    from banks import boa, im
    rows = pd.DataFrame({"payment_status": ["EXPIRED", "NEW", "PAID"]})
    for bank in (im, boa, ecobank):
        assert bank.pending(rows).tolist() == [True, True, False]


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
    assert total == len(review) == review["invoice_number"].nunique()
    weeks = [b[2]["Duplicate Payment"] for b in res.blocks if b[0] == "Summary-Weekly"]
    assert sum(int(w[w["Currency"].isna()]["Number of Transactions"].iloc[0]) for w in weeks) == total
    sheet = res.sheets["Duplicate payment other channel"]
    assert len(sheet) == sheet["invoice_number"].nunique() == total
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
