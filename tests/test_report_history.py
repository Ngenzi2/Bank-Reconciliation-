"""Where a report is saved on Drive: the notebook layout (default) and the app's own layout."""
import engine as e
from services import report_history as rh


def test_notebook_layout_uses_bank_name_and_weekly_or_monthly():
    week = e.make_period("week", "2026-08-03")
    month = e.make_period("month", "2026-08-01")
    assert rh.location_parts("IM", week, "notebook") == ("I&M", "Weekly")
    assert rh.location_parts("IM", month, "notebook") == ("I&M", "Monthly")


def test_app_layout_uses_bank_full_name_and_the_month_folder():
    month = e.make_period("month", "2026-08-01")
    assert rh.location_parts("IM", month, "app") == ("I&M Bank", "2026-08")


def test_drive_location_defaults_to_the_notebook_layout():
    week = e.make_period("week", "2026-08-03")
    assert rh.drive_location("BPR", week) == f"{rh.REPORTS_FOLDER}/BPR/Weekly"
    assert rh.drive_location("BPR", week, "app") == f"{rh.REPORTS_FOLDER}/BPR Bank/2026-08"
