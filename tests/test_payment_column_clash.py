"""A payment file sometimes carries the bank's own columns too (someone re-uploads a report's own Payment details
sheet by mistake, which already has the bank's column next to invoice_number): engine.run() must not let that
silently rename the statement's own column in the merge (<name>_x / <name>_y) - found 2026-10-28, a real Ecobank
run crashed with KeyError('Request Date') this way."""
import pandas as pd

import engine as e
from boa_fixtures import payments, statement
from test_boa import PAYS, ROWS, SEP, read


def test_a_payment_file_with_the_bank_s_own_column_does_not_crash_the_run(tmp_path):
    pay_df = pd.read_csv(pd.io.common.BytesIO(payments(PAYS)))
    pay_df["Value Date"] = "contamination"           # BOA's own statement column (payment_date), by mistake
    contaminated = pay_df.to_csv(index=False).encode()
    files = [read(statement(ROWS)), e.read_payment_file("p.csv", contaminated, "BOA")]
    assert "Value Date" in files[1].data.columns      # confirms the file really is contaminated
    res = e.run(tmp_path / "base", SEP, files, bank="BOA", report_dir=tmp_path / "r", require_review=False)
    assert res.bank_rows > 0
