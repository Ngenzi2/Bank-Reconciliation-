"""I&M (23/11/2026): the weekend's payments are posted on Monday, so Monday's statement (header 23/11 to 23/11) also
covers Saturday and Sunday: no "21/11 to 22/11 missing" when the week's statements are all there."""
import pandas as pd

import engine as e


def statement(first, last, currency="RWF"):
    name = f"im_{currency}_{first}_{last}.xls"
    return e.UploadedFile(name, "bank", e.file_digest(name.encode()), pd.DataFrame(), currency,
                          pd.Timestamp(first), pd.Timestamp(last), 1, bank="IM")


def lines(chk):
    return [t for _, t in chk.lines if t.startswith("Bank RWF")]


def test_monday_covers_the_weekend_before_it():
    """The files of 16-22/11/2026 as downloaded: 16-18, 19-20 and Monday 23 (which holds 21 and 22)."""
    files = [statement("2026-11-16", "2026-11-18"), statement("2026-11-19", "2026-11-20"),
             statement("2026-11-23", "2026-11-23")]
    chk = e.check_period(e.make_period("week", "2026-11-16"), files, "IM", today="2026-11-24")
    assert lines(chk) == ["Bank RWF: 16/11 to 23/11 covered"]
    assert not chk.missing


def test_without_monday_the_weekend_is_still_missing():
    files = [statement("2026-11-16", "2026-11-18"), statement("2026-11-19", "2026-11-20")]
    chk = e.check_period(e.make_period("week", "2026-11-16"), files, "IM", today="2026-11-24")
    assert "Bank RWF: 21/11 to 23/11 missing" in lines(chk)[0]


def test_a_weekend_inside_a_month_is_covered_by_its_monday():
    """November 2026, statements of working days only: the weekends between them are not "missing"."""
    days = [d for d in pd.date_range("2026-11-01", "2026-11-30") if e.is_working_day(d)]
    files = [statement(f"{d:%Y-%m-%d}", f"{d:%Y-%m-%d}") for d in days]
    chk = e.check_period(e.make_period("month", "2026-11-01"), files, "IM", today="2026-12-02")
    assert lines(chk) == ["Bank RWF: 01/11 to 30/11 covered"]


def test_other_banks_keep_their_own_rule():
    """Ecobank has no Monday rule: a Monday statement does not cover the weekend before it."""
    files = [statement("2026-11-16", "2026-11-20"), statement("2026-11-23", "2026-11-23")]
    chk = e.check_period(e.make_period("week", "2026-11-16"), files, "ECO", today="2026-11-24")
    assert any("21/11 to 22/11 missing" in t for _, t in chk.lines)
