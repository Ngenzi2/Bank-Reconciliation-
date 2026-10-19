"""
AUTOMATIC CLASSIFICATION OF ROWS NOT FOUND IN IREMBOPAY (dianah branch: no manual review for normal cases)
==============================================================================================================
Before the dianah branch, every bank row not found in the IremboPay export needed a person to decide (step 5,
Exception Resolution). The rules decide these by themselves now:

    no invoice number                          -> Pending, status "no invoice number"
    IremboPay checkout says Paid                -> Duplicate (paid with another, unnamed channel: "OTHER")
    IremboPay checkout says Expired, or
      "Not an IremboPay invoice" (NOT FOUND)     -> Pending, status EXPIRED (approved 2026-10-12: NOT FOUND
                                                    counts as EXPIRED)
    IremboPay checkout says Open, not paid yet  -> Pending, status NEW
    the checkout could not be asked             -> Pending too (never blocks the report); its invoice goes in
                                                    needs_a_look, a short optional list to check again later

A row already decided (this period, or suggested from an earlier one: engine.review_rows) is left as it is.
"""
from __future__ import annotations

import pandas as pd

import engine

_CHECKOUT_TO_PENDING_STATUS = {"Expired": "EXPIRED", "Not an IremboPay invoice": "EXPIRED",
                               "Open, not paid yet": "NEW"}


def classify(review: pd.DataFrame, answers: dict) -> tuple:
    """review: a RunResult.review (or engine.review_rows()'s). answers: {invoice: irembopay_checkout.check()'s
    answer}, already looked up (integrations.irembopay_checkout.check_many). Returns (the review with every row
    decided, ready for engine.save_decisions(); the invoice numbers that need a look later)."""
    review = review.copy()
    needs_a_look = []
    todo = review.index[~review["Reviewed (Yes/No)"].map(engine.is_yes)]
    for i in todo:
        invoice = review.at[i, "invoice_number"]
        invoice = str(invoice).strip() if pd.notna(invoice) else ""
        if not invoice:
            review.loc[i, ["Reviewed (Yes/No)", "Report table", "payment_status"]] = (
                "Yes", "Pending notification", engine.NO_INVOICE_STATUS)
            continue
        answer = answers.get(invoice)
        if answer == "Paid":
            review.loc[i, ["Reviewed (Yes/No)", "Report table"]] = ("Yes", "Paid with other channel")
        elif answer in _CHECKOUT_TO_PENDING_STATUS:
            review.loc[i, ["Reviewed (Yes/No)", "Report table", "payment_status"]] = (
                "Yes", "Pending notification", _CHECKOUT_TO_PENDING_STATUS[answer])
        else:                                       # could not check: no internet, site down, an unexpected reply
            review.loc[i, ["Reviewed (Yes/No)", "Report table", "payment_status"]] = (
                "Yes", "Pending notification", engine.COULD_NOT_CHECK_STATUS)
            needs_a_look.append(invoice)
    return review, needs_a_look
