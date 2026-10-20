"""Made-up BPR statements and IremboPay files, built with the same structure as the real exports (BOOKING_DATE,
BANK_REFERENCE, DESCRIPTION, a DOC-like column, VALUE_DATE, DEBIT, CREDIT, CLOSING_BALANCE). No real account or
invoice data."""
import csv
import io

from openpyxl import Workbook

PAY_COLUMNS = ["invoice_number", "transaction_id", "amount", "currency", "payment_status", "description", "expiry_at",
               "created_at", "settled_at", "paid_at", "name", "transaction_reference", "provider_reference",
               "payment_channel", "payment_provider", "transaction_status"]


def statement(rows, doc_header="DOC ID", with_value_date=True, balance=100000):
    """rows: [(BOOKING_DATE, BANK_REFERENCE, DESCRIPTION, doc id or None, VALUE_DATE or None, DEBIT, CREDIT)].
    Returns the .xlsx bytes (one sheet, header row first)."""
    header = ["BOOKING_DATE", "BANK_REFERENCE", "DESCRIPTION", doc_header]
    if with_value_date:
        header.append("VALUE_DATE")
    header += ["DEBIT", "CREDIT", "CLOSING_BALANCE"]
    wb = Workbook()
    ws = wb.active
    ws.append(header)
    bal = balance
    for booking, ref, desc, doc, value, debit, credit in rows:
        bal = bal + (credit or 0) + (debit or 0)
        row = [booking, ref, desc, doc]
        if with_value_date:
            row.append(value)
        row += [debit, credit, bal]
        ws.append(row)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def payments(rows):
    """rows: [dict(invoice, amount, paid_at, ref='', status='PAID', provider='BPR', currency='RWF')]."""
    out = io.StringIO()
    w = csv.writer(out)
    w.writerow(PAY_COLUMNS)
    for i, r in enumerate(rows):
        w.writerow([r["invoice"], f"T{i:06d}", f"{r['amount']:.2f}", r.get("currency", "RWF"),
                   r.get("status", "PAID"), r.get("description", ""), "", r["paid_at"], "", r["paid_at"],
                   "Test Merchant", "", r.get("ref", ""), r.get("channel", ""), r.get("provider", "BPR"),
                   "SUCCESSFUL"])
    return out.getvalue().encode()
