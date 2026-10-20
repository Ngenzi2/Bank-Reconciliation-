"""Ecobank bugs found while checking the Slack message against the real samples (2026-10-20)."""
import pandas as pd

from banks import ecobank


def test_in_period_does_not_crash_when_no_statement_covers_the_period():
    """A period whose statement files have not been added yet (every attempt falls outside it): pandas crashes
    mapping an empty Series unless in_period() guards it. The invoices are then simply not inside the period."""
    attempts = pd.DataFrame({"invoice_number": ["886200000001", "886200000002"],
                             "Request Date": pd.to_datetime(["2026-10-05", "2026-10-06"])})
    rows = pd.DataFrame({"invoice_number": ["886200000001", "886200000002"],
                        "Request Date": pd.to_datetime(["2026-10-05", "2026-10-06"]),
                        "paid_at": pd.to_datetime([pd.NaT, pd.NaT])})
    start, end = pd.Timestamp("2026-10-12"), pd.Timestamp("2026-10-18")
    result = ecobank.in_period(rows, attempts, start, end)
    assert result.tolist() == [False, False]
