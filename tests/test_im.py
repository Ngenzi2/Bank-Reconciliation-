"""I&M: reading the statement header (the common rules are in engine.py, tested elsewhere)."""
import io

import pandas as pd
from openpyxl import Workbook

from banks import im


def statement(date_from="05/08/2026", date_to="06/08/2026"):
    """A minimal I&M statement: the header dates, then one IremboPay credit."""
    wb = Workbook()
    ws = wb.active
    ws.append(["I&M Bank statement"])
    ws.append(["Date From(dd/MM/yyyy):", date_from])
    ws.append(["Date To(dd/MM/yyyy):", date_to])
    ws.append(["Transaction Date", "Value Date", "Description", "Tran Id", "Debit", "Credit", "Balance (RWF)"])
    ws.append(["05/08/2026", "05/08/2026", "MOBILE IREMBO PAYMENT 886300000010", "T00010", "", "4,000", "1,004,000"])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def test_header_dates_are_read_day_first():
    # 05/08/2026 is 5 August, not May 8 - the label itself says "(dd/MM/yyyy)"
    data, other_credits, info = im.read_statement(statement(), "statement.xlsx")
    assert info["date_from"] == pd.Timestamp("2026-08-05")
    assert info["date_to"] == pd.Timestamp("2026-08-06")


def test_header_dates_past_the_12th_are_unambiguous_either_way():
    # a sanity check: a day > 12 cannot be misread as a month, so this must keep working
    data, other_credits, info = im.read_statement(statement("20/08/2026", "21/08/2026"), "statement.xlsx")
    assert info["date_from"] == pd.Timestamp("2026-08-20")
    assert info["date_to"] == pd.Timestamp("2026-08-21")
