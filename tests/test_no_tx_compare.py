"""A week compared with a previous week that had no transactions (its No-Transaction Report): the comparison is
made (never "no report found"), and the change reads "No change" (still none) or "+100.00%" (some now), not N/A."""
import pandas as pd
import pytest

import engine as e
from test_boa import PAYS, ROWS, payments, read, statement

WEEK = e.make_period("week", "2026-09-07")          # 07/09 - 13/09/2026: Bank of Africa credits in the fixtures
PREV = WEEK.previous()                               # 31/08 - 06/09/2026: a week reported with no transactions


def no_tx_report(tmp_path, bank="BOA", period=PREV):
    path = tmp_path / f"prev_{bank}.xlsx"
    path.write_bytes(e.no_transaction_excel(period, bank, pd.Timestamp("2026-09-07 09:00"), None))
    return path


def test_a_week_is_compared_with_a_no_transaction_week(tmp_path):
    files = [read(statement(ROWS)), e.read_payment_file("p.csv", payments(PAYS), "BOA")]
    res = e.run(tmp_path / "base", WEEK, files, compare=True, bank="BOA", report_dir=tmp_path / "r",
                previous_report=no_tx_report(tmp_path), require_review=False)
    assert "No-Transaction Report" in res.compared_with
    t = res.blocks[0][2]
    processed = t["Overall processed payments"]
    assert processed["Number of Transactions"].iloc[-1] > 0
    assert processed["VS prev Week (%)"].tolist()[-1] == "+100.00%"            # some now, none before
    assert "No change" in t["Pending notifications"]["VS prev Week (%)"].tolist()   # none now, none before
    assert "N/A" not in [v for tb in t.values() for v in tb["VS prev Week (%)"]]
    assert set(res.merchants[res.blocks[0][1]]["VS prev Week (%)"]) <= {"+100.00%", "No change"}
    assert e.read_report(res.report_file)["blocks"]                           # the report reads back


def test_a_no_transaction_report_can_be_uploaded_as_the_previous_one(tmp_path):
    target = e.save_uploaded_report(tmp_path, PREV, "BOA_prev.xlsx", no_tx_report(tmp_path).read_bytes(), "BOA",
                                    root=tmp_path / "store")
    assert target.exists()
    with pytest.raises(e.ReconError, match="not for Bank of Africa"):
        e.save_uploaded_report(tmp_path, PREV, "IM.xlsx", no_tx_report(tmp_path, "IM").read_bytes(), "BOA",
                               root=tmp_path / "store")
    with pytest.raises(e.ReconError, match="not for"):
        e.save_uploaded_report(tmp_path, WEEK, "BOA_prev.xlsx", no_tx_report(tmp_path).read_bytes(), "BOA",
                               root=tmp_path / "store")


def test_two_normal_weeks_keep_n_a_when_the_previous_count_was_0():
    tables = {"Pending notifications": pd.DataFrame({"Currency": ["RWF", None], "Number of Transactions": [3, 3]})}
    prev = {"Pending notifications": {"RWF": 0, "Total": 0}}
    assert e.add_vs_column(tables, prev, "Week")["Pending notifications"]["VS prev Week (%)"].tolist() == ["N/A", "N/A"]


def test_bank_of_africa_days_without_statement_warn_but_do_not_block():
    """Bank of Africa's export has no from/to dates: a day without transaction looks 'missing'. Run stays possible;
    the days are listed (shown in step 6 and in the Slack message)."""
    files = [read(statement(ROWS)), e.read_payment_file("p.csv", payments(PAYS), "BOA")]
    chk = e.check_period(e.make_period("week", "2026-10-12"), files, "BOA", today="2026-10-19")
    assert not any("missing" in b for b in chk.blocking) and chk.missing
    assert any("may simply have had no transaction" in w for w in chk.warnings)
    im = e.check_period(e.make_period("week", "2026-10-12"), [], "IM", today="2026-10-19")
    assert im.blocking                                                   # I&M stays strict


def test_the_slack_message_says_which_days_had_no_statement():
    from services import slack_message as M
    import test_slack as T
    p = M.build_blocks(e.get_bank("BOA"), WEEK, [("Summary-Weekly", T.HEAD, T.block())], {}, None,
                       pd.Timestamp("2026-10-19"), notes=["Bank RWF: 17/10 to 18/10"])
    text = M.preview_text(p)
    assert "⚠️ No statement for 17/10 to 18/10 (Bank RWF): no transaction those days, or the statement is missing" in text
