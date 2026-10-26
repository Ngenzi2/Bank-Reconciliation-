"""20/10/2026: a payment made on a weekend or a public holiday is posted on the next working day, so a period that
ends on one needs the statement of the next working day (I&M); the day before is no longer recommended."""
import pandas as pd
import pytest

import engine as e
from services import app_settings


@pytest.fixture(autouse=True)
def holidays_2026():
    e.set_holidays(app_settings.HOLIDAYS_2026)
    yield
    e.set_holidays([])


def test_week_ending_before_easter_needs_the_next_working_day():
    # 30/03 - 05/04 (Friday 03/04 = Good Friday); Monday 06/04 = Easter Monday and Tuesday 07/04 = Genocide Memorial Day -> Wednesday 08/04
    week = e.make_period("week", "2026-03-30")
    assert e.extra_monday(week, "IM") == pd.Timestamp("2026-04-08")
    assert [f"{d:%d/%m}" for d in e.closed_days_before(pd.Timestamp("2026-04-08"))] == ["03/04", "04/04", "05/04", "06/04",
                                                                                        "07/04"]


def test_a_day_on_a_holiday_needs_the_next_working_day():
    assert e.extra_monday(e.make_period("day", "2026-07-01"), "IM") == pd.Timestamp("2026-07-02")   # Independence Day
    assert e.extra_monday(e.make_period("day", "2026-07-02"), "IM") is None                         # a working day


def test_a_normal_weekend_still_needs_monday_and_other_banks_nothing():
    week = e.make_period("week", "2026-10-12")
    assert e.extra_monday(week, "IM") == pd.Timestamp("2026-10-19")
    assert e.extra_monday(week, "BOA") is None


def test_the_day_before_is_no_longer_recommended_for_im():
    from banks import im
    assert im.DAY_BEFORE_CURRENCIES == []


def test_holidays_are_saved_and_checked(tmp_path):
    path = tmp_path / "s.json"
    data = app_settings.load(path)
    assert any(h["name"] == "Easter Monday" for h in data["holidays"])
    data = app_settings.set_holidays([{"date": "2026-12-25", "name": "Christmas Day"}], "ed@irembo.com", path)
    assert app_settings.load(path)["holidays"] == [{"date": "2026-12-25", "name": "Christmas Day"}]
    assert data["history"][-1]["what"] == "Public holidays"
    with pytest.raises(app_settings.SettingsError):
        app_settings.set_holidays([{"date": "25/12/2026", "name": "x"}], "ed@irembo.com", path)
