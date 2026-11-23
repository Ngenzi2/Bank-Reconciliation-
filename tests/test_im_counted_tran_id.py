"""I&M (23/11/2026): an invoice credited twice - the second credit shows, in "Counted Tran Id", the Tran Id of the
credit counted for its invoice (the counted credit itself: empty); the column comes right after the bank's columns."""
import pandas as pd

import engine as e
from boa_fixtures import payments

WEEK = e.make_period("week", "2026-11-16")


def statement(rows):
    data = pd.DataFrame([{"Transaction Date": d, "Value Date": d, "Description": f"AG IREMBO PAYMENT {inv}",
                          "Tran Id": tran, "Credit": amount, "Currency": "RWF", "Source file": "im.xls"}
                         for d, inv, tran, amount in rows])
    return e.UploadedFile("im.xls", "bank", e.file_digest(str(rows).encode()), data, "RWF",
                          pd.Timestamp("2026-11-16"), pd.Timestamp("2026-11-23"), len(data), bank="IM")


def test_the_second_credit_shows_the_tran_id_counted(tmp_path):
    rows = [("19/11/2026", "886323167318", "0227813", "500.00"),        # counted
            ("19/11/2026", "886323167318", "0228662", "500.00"),        # the same invoice again
            ("19/11/2026", "886323000001", "0228700", "1,000.00")]
    pays = [dict(invoice="886323167318", amount=500, paid_at="2026-11-19 14:09:15", ref="P1", provider="I_M",
                 channel="I_M_APP_WEB"),
            dict(invoice="886323000001", amount=1000, paid_at="2026-11-19 15:00:00", ref="P2", provider="I_M",
                 channel="I_M_APP_WEB")]
    files = [statement(rows), e.read_payment_file("p.csv", payments(pays), "IM")]
    res = e.run(tmp_path / "base", WEEK, files, require_review=False, bank="IM", report_dir=tmp_path / "r")
    d = res.sheets["Matched details"]
    by_tran = d.set_index("Tran Id")
    assert by_tran.at["0228662", "Counted Tran Id"] == "0227813"
    assert by_tran.at["0227813", "Counted Tran Id"] == "" and by_tran.at["0228700", "Counted Tran Id"] == ""
    cols = list(d.columns)
    assert cols[cols.index("Credit") + 1] == "Counted Tran Id"
    assert "Counted credit" not in cols                                  # BK's name stays BK's
