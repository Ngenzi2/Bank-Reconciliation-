"""Exception Resolution (17/11/2026): the Report table of each row the agreed rules can decide is filled in for the
reviewer to confirm; a choice already made is never replaced, and what the rules cannot decide stays empty."""
import pandas as pd

import engine as e


def review(rows):
    """rows: (invoice, reason, VALUE_DATE, Report table already chosen)."""
    out = pd.DataFrame(rows, columns=["invoice_number", "Reason", "VALUE_DATE", "Report table"])
    for c in e.EDIT_COLS:
        if c not in out:
            out[c] = ""
    return out


PAYS = pd.DataFrame({
    "invoice_number": ["886000000001", "886000000002", "886000000003", "886000000004", "886000000005"],
    "paid_at": pd.to_datetime(["2026-10-27 10:00", "2026-10-28 09:00", "2026-10-27 11:00", None, "2026-10-27 12:00"]),
    "payment_provider": ["BPR_RW", "BPR_RW", "MTN", None, "BPR_RW"],
    "transaction_status": ["SUCCESSFUL"] * 3 + [None, "SUCCESSFUL"],
    "created_at": pd.to_datetime(["2026-10-27"] * 5)})


def test_each_case_gets_the_table_the_rules_give():
    rows = review([
        ("886000000001", "No VALUE_DATE in the statement (BOOKING_DATE used for the period)", "2026-10-27", ""),  # same day
        ("886000000002", "No VALUE_DATE in the statement (BOOKING_DATE used for the period)", "2026-10-27", ""),  # next day
        ("886000000003", "Paid through MTN in IremboPay (not BPR)", "2026-10-27 10:00", ""),
        ("886000000004", "No VALUE_DATE in the statement (BOOKING_DATE used for the period)", "2026-10-27", ""),  # NEW
        ("886000000005", "Invoice credited again: IremboPay has no record of this credit; the credit counted is FT1",
         "2026-10-27 12:00", ""),
        ("886000000006", "DOC ID not in the IremboPay payment file", "2026-10-27 13:00", ""),              # checkout: Paid
        ("886000000007", "DOC ID not in the IremboPay payment file", "2026-10-27 13:00", ""),              # Expired
        ("886000000008", "DOC ID not in the IremboPay payment file", "2026-10-27 13:00", ""),              # unknown
        ("886000000009", "DOC ID not in the IremboPay payment file; cancelled by the debit FT77", "2026-10-27", ""),
        (None, "No DOC ID", "2026-10-27 14:00", ""),
        ("886000000010", "DOC ID not in the IremboPay payment file", "2026-10-27 13:00", "Failed notification"),  # chosen
    ])
    answers = {"886000000006": "Paid", "886000000007": "Expired", "886000000008": "Not an IremboPay invoice",
               "886000000010": "Paid"}
    out, filled = e.rule_decisions(rows, PAYS, "BPR", answers)
    assert out["Report table"].tolist() == [
        "Notification received", "Notification received", "Paid with other channel", "Pending notification", "",
        "Paid with other channel", "Pending notification", "", "Exclude", "Pending notification", "Failed notification"]
    assert out["SLA violated (Yes/No)"].eq("").all()                           # the SLA is never filled
    assert out["payment_status"].tolist()[3] == "NEW" and out["payment_status"].tolist()[6] == "NEW"
    assert out["Comment"].tolist()[6] == "Invoice expired (IremboPay checkout)"
    assert out["Comment"].tolist()[8] == "Cancelled by the debit FT77: not a payment"
    assert out["Reviewed (Yes/No)"].eq("").all()                               # the reviewer confirms
    assert sum(n for _, n, _ in filled) == 8


def test_nothing_to_decide_leaves_the_table_as_it_was():
    rows = review([("886000000008", "DOC ID not in the IremboPay payment file", "2026-10-27", "")])
    out, filled = e.rule_decisions(rows, PAYS.iloc[:0], "BPR", {})
    assert out["Report table"].tolist() == [""] and filled == []


def test_an_ecobank_credit_without_invoice_is_pending():
    rows = pd.DataFrame({"invoice_number": ["IREMBO|27-Oct-2026|000162"], "Reason": ["Not found in IremboPay"],
                         "Request Date": pd.to_datetime(["2026-10-27 12:52"]), "Report table": [""]})
    for c in e.EDIT_COLS:
        if c not in rows:
            rows[c] = ""
    out, _ = e.rule_decisions(rows, PAYS, "ECO", {"IREMBO|27-Oct-2026|000162": "Not an IremboPay invoice"})
    assert out["Report table"].tolist() == ["Pending notification"] and out["payment_status"].tolist() == ["NEW"]
