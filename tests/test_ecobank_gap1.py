"""Ecobank: Processed starts from the bank's own SUCCESS rows, windowed by Request Date - never by IremboPay's
paid_at (CLAUDE.md gap #1, fixed 2026-10-25). Processed = Notified + Pending + Duplicate always holds now, even
for an invoice IremboPay never matched at all."""
import io

import pandas as pd
import pytest
from openpyxl import Workbook

import engine as e
from banks import ecobank

WEEK = e.make_period("week", "2026-09-07")

PAY_COLUMNS = ["invoice_number", "transaction_id", "amount", "currency", "payment_status", "description", "expiry_at",
              "created_at", "settled_at", "paid_at", "name", "transaction_reference", "provider_reference",
              "payment_channel", "payment_provider", "transaction_status"]


def statement(rows):
    """rows: [(request_date 'YYYY-MM-DD HH:MM:SS' UTC, invoice, cba_rsp_msg, cba_ref, amount)]."""
    wb = Workbook()
    ws = wb.active
    ws.append(["Source Channel", "Biller Code", "Bill Ref No", "Ccy", "Amount", "Request Date", "Cba Rsp Code",
              "Cba Rsp Msg", "Biller Rsp Code", "Biller Rsp Msg", "Cba Ref No", "Customer Ref No"])
    for request, invoice, cba_msg, cba_ref, amount in rows:
        code = "000" if cba_msg == "SUCCESS" else "400"
        ws.append(["AGENCYBANKING", "", invoice, "RWF", amount, request, code, cba_msg, "000", "SUCCESS", cba_ref, ""])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def payments(rows):
    out = io.StringIO()
    import csv
    w = csv.writer(out)
    w.writerow(PAY_COLUMNS)
    for i, r in enumerate(rows):
        w.writerow([r["invoice"], f"T{i:06d}", f"{r['amount']:.2f}", "RWF", r.get("status", "PAID"), "", "",
                   r.get("paid_at", ""), "", r.get("paid_at", ""), "Test Merchant", "", r.get("ref", ""), "",
                   r.get("provider", "ECOBANK"), "SUCCESSFUL"])
    return out.getvalue().encode()


def test_a_bank_success_row_irembopay_never_matched_is_still_processed_and_pending(tmp_path):
    # 886200000001: Ecobank says SUCCESS, requested inside the week - but IremboPay has no record of it at all,
    # so the rules classify it from the checkout page (here, a saved decision stands in for that check: NEW)
    rows = [("2026-09-08 09:00:00", "886200000001", "SUCCESS", "R1", 1000),
           ("2026-09-09 09:00:00", "886200000002", "SUCCESS", "R2", 2000)]
    files = [e.read_bank_file("s.xlsx", statement(rows), "ECO"),
            e.read_payment_file("p.csv", payments([
                dict(invoice="886200000002", amount=2000, paid_at="2026-09-09 09:00:00", ref="R2")]), "ECO")]
    first = e.run(tmp_path / "base", WEEK, files, require_review=True, bank="ECO", report_dir=tmp_path / "r1")
    review = first.review.copy()
    review["Reviewed (Yes/No)"] = "Yes"
    review["Report table"] = "Pending notification"
    review["payment_status"] = "NEW"
    e.save_decisions(tmp_path / "base", WEEK, review, "ECO")
    res = e.run(tmp_path / "base", WEEK, files, require_review=True, bank="ECO", report_dir=tmp_path / "r2")
    assert not res.stopped_for_review
    t = e.counts_of(res.blocks[-1][2])
    processed = t["Overall processed payments"]["Total"]
    notified = t["Overall notifications received by Irembo"]["Total"]
    pending = t["Pending notifications"]["Total"]
    duplicate = t["Duplicate Payment"]["Total"]
    assert processed == 2                                       # both SUCCESS rows, matched or not
    assert pending == 1 and notified == 1 and duplicate == 0     # the unmatched one is Pending, not left out
    assert processed == notified + pending + duplicate


def test_a_request_outside_the_week_is_not_processed_even_if_irembopay_paid_at_falls_inside_it(tmp_path):
    # the bank requested this one in August; IremboPay's own paid_at happens to be in the chosen week - it must
    # NOT be counted this week (the bank statement is the source of truth, never IremboPay's own date)
    rows = [("2026-08-20 09:00:00", "886200000003", "SUCCESS", "R3", 1000)]
    files = [e.read_bank_file("s.xlsx", statement(rows), "ECO"),
            e.read_payment_file("p.csv", payments([
                dict(invoice="886200000003", amount=1000, paid_at="2026-09-08 09:00:00", ref="R3")]), "ECO")]
    with pytest.raises(e.NoTransactionsError):
        e.run(tmp_path / "base", WEEK, files, require_review=False, bank="ECO", report_dir=tmp_path / "r")
