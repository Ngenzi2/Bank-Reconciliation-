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


def test_a_week_that_crosses_a_month_boundary_belongs_to_its_thursdays_month():
    # 28 Sep - 4 Oct 2026: the Thursday is 1 October, so this week's month is October, not September
    week = e.make_period("week", "2026-09-28")
    assert week.month == "2026-10"
    assert rh.month_folder(week) == "2026-10"
    assert rh.location_parts("IM", week, "app") == ("I&M Bank", "2026-10")


def test_a_saved_weekly_report_is_filed_under_its_thursdays_month():
    # Report History / Dashboard group a saved report by Record.month - the same rule, read back from its name
    record = rh.Record("Google Drive", "IM", "week", e.make_period("week", "2026-09-28").start,
                       "Week · 28/09 - 04/10/2026", "IM_recon_summary_28-09-2026_to_04-10-2026.xlsx", "xlsx")
    assert record.month == "2026-10"
