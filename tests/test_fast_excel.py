"""04/11/2026: a big detail sheet (more than FAST_SHEET_ROWS rows, e.g. a BK month) is written straight into the report
after openpyxl wrote its first rows. The file must be the same as the one openpyxl alone writes: every filled cell with
the same value, number format and type, the same filter, frozen header and column widths."""
import pandas as pd
from openpyxl import load_workbook

import engine as e


def table(n=60):
    return pd.DataFrame({
        "Currency": ["RWF", "USD", None] * (n // 3),
        "BookingDate": pd.to_datetime(["2026-08-03", None, "2026-08-04"] * (n // 3)),
        "CreditAmt": [5000.0, 12.5, float("nan")] * (n // 3),
        "Narration": ["Bill <IREMBO> & co", " leading space", ""] * (n // 3),
        "ProcessingDateTime": pd.to_datetime(["2026-08-03 10:15:00", "2026-08-03 23:59:59", None] * (n // 3)),
        "found_in_irembopay": [True, False, True] * (n // 3),
        "times_on_statement": [1, 2, 3] * (n // 3),
        "review_comment": [None, "ok", 7] * (n // 3),                  # a mixed column
        "invoice_number": ["886000000001", None, "886000000003"] * (n // 3),
    })


def write(path, sheets, limit, monkeypatch):
    monkeypatch.setattr(e, "FAST_SHEET_ROWS", limit)
    monkeypatch.setattr(e, "FAST_SAMPLE", 5)
    e._write_report(path, [], sheets, "BK")


def test_a_big_sheet_is_written_the_same_as_openpyxl_writes_it(tmp_path, monkeypatch):
    sheets = {"BK statement": table(), "Debits (BK)": table(9)}
    slow, fast = tmp_path / "slow.xlsx", tmp_path / "fast.xlsx"
    write(slow, sheets, 10**9, monkeypatch)                      # openpyxl only
    write(fast, sheets, 20, monkeypatch)                         # the 60-row sheet the fast way, the 9-row one not
    for name in sheets:
        pd.testing.assert_frame_equal(pd.read_excel(slow, sheet_name=name), pd.read_excel(fast, sheet_name=name))
    a, b = load_workbook(slow)["BK statement"], load_workbook(fast)["BK statement"]
    assert (a.max_row, a.auto_filter.ref, a.freeze_panes) == (b.max_row, b.auto_filter.ref, b.freeze_panes) == \
        (61, "A1:I61", "A2")
    assert all(a.column_dimensions[k].width == b.column_dimensions[k].width for k in a.column_dimensions)
    for ra, rb in zip(a.iter_rows(), b.iter_rows()):
        for ca, cb in zip(ra, rb):
            assert ca.value == cb.value
            if ca.value is not None:
                assert (ca.number_format, ca.font.b, ca.data_type) == (cb.number_format, cb.font.b, cb.data_type)
