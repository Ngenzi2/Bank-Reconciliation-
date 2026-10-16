"""Made-up Bank of Africa statements and IremboPay files, built with the same structure as the real exports
(account details, header, numbered rows newest first, Totals row, footer). No real account or invoice data."""
import csv
import io
from datetime import datetime

from openpyxl import Workbook

PAY_COLUMNS = ["invoice_number", "transaction_id", "amount", "currency", "payment_status", "description", "expiry_at",
               "created_at", "settled_at", "paid_at", "name", "transaction_reference", "provider_reference",
               "payment_channel", "payment_provider", "transaction_status"]


def statement(rows, currency="RWF", opening=1000, count=None, total_credit=None, total_debit=None, closing=None,
              blank_lines_before_header=1, extra_rows_after=(), date_cells=False, repeat_header_at=None,
              drop_column=None):
    """rows: [(date 'dd/mm/yyyy', ref, description, debit or None, credit or None)] OLDEST FIRST; written newest
    first with row numbers and balances, like the real export. Returns the .xlsx bytes."""
    balances, bal = [], opening
    for _, _, _, d, c in rows:
        bal = bal + (c or 0) + (d or 0)
        balances.append(bal)
    credits = sum(c or 0 for *_, c in rows)
    debits = sum(d or 0 for *_, d, _ in rows)
    wb = Workbook()
    ws = wb.active
    ws.title = "Account history"
    for line in (["Bank of Africa Rwanda"], ["Account history"], ["Customer: TEST CUSTOMER - TEST LTD"],
                 ["Generated on 10 Oct 2026 09:00"], [], [f"Last {len(rows)} transaction (does not include today "
                                                          "transaction) for account 00000000000"],
                 ["Account Name", "TEST LTD"], ["Account type", "Current Account"], ["Account currency", currency],
                 ["Opening Bal", opening], ["Closing Bal", closing if closing is not None else bal],
                 ["Total Credit", total_credit if total_credit is not None else credits],
                 ["Total Debit", total_debit if total_debit is not None else -debits],
                 ["Transaction Count", count if count is not None else len(rows)]):
        ws.append(line)
    for _ in range(blank_lines_before_header):
        ws.append([])
    header = ["#", "Ledger Date", "Value Date", "Branch", "Ref no", "Description", "Debit", "Credit", "Balance"]
    if drop_column:
        header = [h if h != drop_column else "Something else" for h in header]
    ws.append(header)
    for n, i in enumerate(range(len(rows) - 1, -1, -1), start=1):
        day, ref, desc, d, c = rows[i]
        when = datetime.strptime(day, "%d/%m/%Y") if date_cells else day
        ws.append([n, when, when, "HEAD OFFICE", ref, desc, d, c, balances[i]])
        if repeat_header_at == n:
            ws.append(header)
    ws.append(["Totals (D / C)", None, None, None, None, None, debits, credits, None])
    ws.append([])
    ws.append(["Bank of Africa Rwanda - e-banking - Account history. Generated from your internet banking session."])
    for r in extra_rows_after:
        ws.append(r)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def payments(rows):
    """rows: [dict(invoice, amount, paid_at, ref, status='PAID', provider='BANK_OF_AFRICA', channel='BOA_IB',
    currency='RWF', description='')] -> IremboPay CSV bytes."""
    out = io.StringIO()
    w = csv.writer(out)
    w.writerow(PAY_COLUMNS)
    for i, r in enumerate(rows):
        w.writerow([r["invoice"], f"T{i:06d}", f"{r['amount']:.2f}", r.get("currency", "RWF"), r.get("status", "PAID"),
                    r.get("description", ""), "", r["paid_at"], "", r["paid_at"], "Test Merchant", "", r["ref"],
                    r.get("channel", "BOA_IB"), r.get("provider", "BANK_OF_AFRICA"), "SUCCESSFUL"])
    return out.getvalue().encode()
