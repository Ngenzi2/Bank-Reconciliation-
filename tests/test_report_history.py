"""Recognising the reconciliation notebooks' own saved reports (not this app's), wherever they are on Drive:
report_history.parse_notebook_name, drive_index falling back to it, and find_previous leaving them out (it needs
a report in this app's own layout)."""
import pandas as pd

import engine as e
from services import report_history as rh


def test_a_monthly_notebook_report_is_recognised_in_the_bank_s_monthly_folder():
    p = rh.parse_notebook_name("Ecobank_report_August_2026", "Ecobank / Monthly", "2026-09-02T10:00:00Z")
    assert p == {"code": "ECO", "kind": "month", "start": pd.Timestamp("2026-08-01"),
                "label": "01-08-2026_to_31-08-2026", "version": 1, "ext": "xlsx",
                "stem": "Ecobank_report_August_2026", "notebook": True}


def test_a_weekly_notebook_report_names_the_week_by_its_thursday():
    # 28 Sep - 4 Oct 2026 is CLAUDE.md's own example of "first week of October"
    p = rh.parse_notebook_name("BPR_report_First_week_of_October_2026", "BPR / Weekly", "2026-10-05T10:00:00Z")
    assert p["start"] == pd.Timestamp("2026-09-28")
    assert p["kind"] == "week" and p["code"] == "BPR"


def test_a_weekly_notebook_report_with_no_year_uses_the_file_s_own_modified_date():
    p = rh.parse_notebook_name("Ecobank_report_First_week_of_August", "Ecobank / Weekly", "2026-08-05T10:00:00Z")
    assert p["start"] == pd.Timestamp("2026-08-03")


def test_a_native_google_sheet_with_no_extension_is_still_recognised():
    p = rh.parse_notebook_name("Ecobank_report_August_2026", "Ecobank / Monthly", "2026-09-02T10:00:00Z")
    p_copy = rh.parse_notebook_name("Ecobank_report_August_2026 (1)", "Ecobank / Monthly", "2026-09-02T10:00:00Z")
    assert p_copy["stem"] == p["stem"]       # a Drive copy saved by hand: " (1)" dropped, same report


def test_a_file_not_in_a_weekly_or_monthly_folder_is_not_a_notebook_report():
    assert rh.parse_notebook_name("Ecobank_report_August_2026", "Ecobank", "2026-09-02T10:00:00Z") is None
    assert rh.parse_notebook_name("Ecobank_report_August_2026", "Ecobank / 2026-08", "2026-09-02T10:00:00Z") is None


def test_a_name_this_app_would_have_made_is_not_read_as_a_notebook_report():
    assert rh.parse_notebook_name("ECO_recon_summary_01-08-2026_to_31-08-2026.xlsx", "Ecobank / Monthly") is None


def test_an_unknown_bank_folder_is_not_recognised():
    assert rh.parse_notebook_name("Nobank_report_August_2026", "Nobank / Monthly", "2026-09-02T10:00:00Z") is None


class _FakeClient:
    """list_tree() returning made-up files; nothing is actually listed or downloaded."""
    def __init__(self, files):
        self._files = files

    def list_tree(self, root_id, depth):
        return self._files, sorted({f.folder for f in self._files if f.folder})


def test_drive_index_recognises_a_notebook_report_next_to_the_app_s_own():
    from integrations.google_drive import DriveFile
    files = [DriveFile("1", "ECO_recon_summary_01-09-2026_to_30-09-2026.xlsx", folder="Ecobank / 2026-09",
                       modified="2026-10-01T08:00:00Z"),
            DriveFile("2", "Ecobank_report_August_2026", folder="Ecobank / Monthly",
                      modified="2026-09-02T10:00:00Z", mime="application/vnd.google-apps.spreadsheet")]
    index = rh.drive_index(_FakeClient(files), "root")
    assert {r.notebook for r in index.records} == {False, True}
    nb = next(r for r in index.records if r.notebook)
    assert nb.bank == "ECO" and nb.kind == "month" and nb.mime == "application/vnd.google-apps.spreadsheet"


def test_find_previous_offers_a_notebook_report_when_it_is_the_only_one():
    from integrations.google_drive import DriveFile
    period = e.make_period("month", "2026-08-01")
    files = [DriveFile("2", "Ecobank_report_August_2026", folder="Ecobank / Monthly",
                       modified="2026-09-02T10:00:00Z", mime="application/vnd.google-apps.spreadsheet")]
    index = rh.drive_index(_FakeClient(files), "root")
    found = rh.find_previous(None, "root", "ECO", period, index=index)
    assert found is not None and found.id == "2" and found.mime == "application/vnd.google-apps.spreadsheet"


def test_find_previous_prefers_the_app_s_own_report_once_a_notebook_one_is_also_listed():
    from integrations.google_drive import DriveFile
    period = e.make_period("month", "2026-08-01")
    files = [DriveFile("1", "ECO_recon_summary_01-08-2026_to_31-08-2026.xlsx", folder="Ecobank / 2026-08",
                       modified="2026-09-01T08:00:00Z"),
            DriveFile("2", "Ecobank_report_August_2026", folder="Ecobank / Monthly",
                      modified="2026-09-02T10:00:00Z")]       # the notebook's own copy of the same month, saved later
    index = rh.drive_index(_FakeClient(files), "root")
    found = rh.find_previous(None, "root", "ECO", period, index=index)
    assert found is not None and found.id == "1"


def test_find_previous_leaves_out_a_different_month():
    from integrations.google_drive import DriveFile
    period = e.make_period("month", "2026-08-01")
    files = [DriveFile("2", "Ecobank_report_September_2026", folder="Ecobank / Monthly",
                       modified="2026-10-02T10:00:00Z")]
    index = rh.drive_index(_FakeClient(files), "root")
    assert rh.find_previous(None, "root", "ECO", period, index=index) is None


def _notebook_workbook(path, base=5):
    """A minimal Summary-Monthly sheet laid out like the notebooks' own (one column to the right of this app's),
    for BOA (one currency: simplest). Only 'Overall processed payments', SLA Violated and Duplicate Payment are
    on the Summary - Pending notifications is counted from its own sheet, as the notebooks do."""
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Summary-Monthly"
    ws.cell(base, 2, "Month: August 2026 (01/08/2026 - 31/08/2026)")
    ws.cell(base + 2, 2, "Overall processed payments")
    ws.cell(base + 2 + e.GS_STEP, 2, "SLA Violated")
    ws.cell(base + 2 + 2 * e.GS_STEP, 2, "Duplicate Payment")
    for top, n, amount in ((base + 2, 10, 100000), (base + 2 + e.GS_STEP, 1, 5000),
                           (base + 2 + 2 * e.GS_STEP, 2, 20000)):
        ws.cell(top + 1, 2, "Currency")
        ws.cell(top + 1, 3, "Number of Transactions")
        ws.cell(top + 1, 4, "Total Amount")
        ws.cell(top + 2, 2, "RWF"); ws.cell(top + 2, 3, n); ws.cell(top + 2, 4, amount)
        ws.cell(top + 3, 3, n); ws.cell(top + 3, 4, amount)                 # the Total row (no Currency)
    pending = wb.create_sheet("Pending notification")
    pending.append(["Ccy", "Amount"])
    pending.append(["RWF", 7000])
    wb.save(path)


def test_read_notebook_report_derives_notified_from_processed_minus_pending_and_duplicate(tmp_path):
    path = tmp_path / "BOA_report_August_2026.xlsx"
    _notebook_workbook(path)
    period = e.make_period("month", "2026-08-01")
    rep = e.read_notebook_report(path, "BOA", period)
    tables = dict(rep["blocks"][0][2].items())
    assert tables["Overall processed payments"]["Number of Transactions"].iloc[0] == 10
    assert tables["Pending notifications"]["Number of Transactions"].iloc[0] == 1
    assert tables["Duplicate Payment"]["Number of Transactions"].iloc[0] == 2
    notified = tables["Overall notifications received by Irembo"]
    assert notified["Number of Transactions"].iloc[0] == 7              # 10 processed - 1 pending - 2 duplicate
    assert notified["Total Amount"].iloc[0] == 73000                    # 100000 - 5000 - 20000


def test_read_notebook_report_also_reads_the_weeks_of_a_monthly_report(tmp_path):
    """27/10/2026 -> 2026-10-28: a monthly notebook report's Summary-Weekly sheet has one block per week (its own
    Pending notifications table too, found at whatever column it is in, not just the Pending notification sheet
    which only has the whole month)."""
    import openpyxl
    path = tmp_path / "BOA_report_August_2026.xlsx"
    _notebook_workbook(path)
    wb = openpyxl.load_workbook(path)
    ws = wb.create_sheet("Summary-Weekly")
    for base, (a, b), proc, sla, dup, pend in ((1, ("2026-08-01", "2026-08-07"), 6, 1, 1, 1),
                                               (26, ("2026-08-08", "2026-08-14"), 4, 0, 1, 1)):
        ws.cell(base, 2, f"Week ({a} to {b})")
        ws.cell(base + 2, 2, "Overall processed payments")
        ws.cell(base + 2 + e.GS_STEP, 2, "SLA Violated")
        ws.cell(base + 2 + 2 * e.GS_STEP, 2, "Duplicate Payment")
        ws.cell(base + 2 + e.GS_STEP, 8, "Pending notifications")           # the notebooks' own column, not derived
        for top, col, n, amount in ((base + 2, 2, proc, proc * 1000), (base + 2 + e.GS_STEP, 2, sla, sla * 1000),
                                    (base + 2 + 2 * e.GS_STEP, 2, dup, dup * 1000),
                                    (base + 2 + e.GS_STEP, 8, pend, pend * 1000)):
            ws.cell(top + 1, col, "Currency")
            ws.cell(top + 1, col + 1, "Number of Transactions")
            ws.cell(top + 1, col + 2, "Total Amount")
            ws.cell(top + 2, col, "RWF"); ws.cell(top + 2, col + 1, n); ws.cell(top + 2, col + 2, amount)
            ws.cell(top + 3, col + 1, n); ws.cell(top + 3, col + 2, amount)
    wb.save(path)
    period = e.make_period("month", "2026-08-01")
    rep = e.read_notebook_report(path, "BOA", period)
    assert [(sheet, heading) for sheet, heading, _ in rep["blocks"]] == [
        ("Summary-Monthly", "Month: August 2026 (01/08/2026 - 31/08/2026)"),
        ("Summary-Weekly", "Week: 01/08/2026 - 07/08/2026"), ("Summary-Weekly", "Week: 08/08/2026 - 14/08/2026")]
    w1 = rep["blocks"][1][2]
    assert w1["Overall processed payments"]["Number of Transactions"].iloc[0] == 6
    assert w1["Pending notifications"]["Number of Transactions"].iloc[0] == 1         # read directly, not derived
    assert w1["Overall notifications received by Irembo"]["Number of Transactions"].iloc[0] == 4   # 6 - 1 - 1


def test_read_report_counts_derives_the_two_tables_the_notebooks_never_write(tmp_path):
    path = tmp_path / "BOA_report_August_2026.xlsx"
    _notebook_workbook(path)
    counts, heading, not_reviewed = e.read_report_counts(path, "Summary-Monthly", "BOA")
    assert counts["Overall processed payments"] == {"RWF": 10, "Total": 10}
    assert counts["Pending notifications"] == {"RWF": 1, "Total": 1}              # from its own sheet
    assert counts["Overall notifications received by Irembo"] == {"RWF": 7, "Total": 7}   # 10 - 1 - 2
    assert not_reviewed == 0


def test_read_notebook_report_refuses_a_file_that_is_not_a_notebook_report(tmp_path):
    import openpyxl
    path = tmp_path / "ECO_report_August_2026.xlsx"
    wb = openpyxl.Workbook()
    wb.active.title = "Summary-Monthly"
    wb.save(path)
    period = e.make_period("month", "2026-08-01")
    try:
        e.read_notebook_report(path, "ECO", period)
        assert False, "should have refused a sheet with no summary table"
    except e.ReconError as err:
        assert "not a report the notebooks made" in str(err)


def test_a_month_is_compared_with_a_notebook_made_previous_month(tmp_path):
    """engine.run()'s own "Compare with previous period" (section H), not just the Dashboard: the previous report
    found by find_previous() can now be one of the reconciliation notebooks' own - this is the actual path that
    used to raise "this report does not have the expected tables" before read_report_counts could derive them."""
    from test_boa import PAYS, ROWS, SEP, payments, read, statement
    prev_path = tmp_path / "BOA_report_August_2026.xlsx"
    _notebook_workbook(prev_path)                        # built for BOA (one currency): matches test_boa's fixtures
    files = [read(statement(ROWS)), e.read_payment_file("Payment details.csv", payments(PAYS), "BOA")]
    res = e.run(tmp_path / "base", SEP, files, compare=True, bank="BOA", report_dir=tmp_path / "r",
               previous_report=prev_path, require_review=False)
    assert "report" in res.compared_with
    sheet, heading, t = next(b for b in res.blocks if b[0] == "Summary-Monthly")
    assert "N/A" not in [v for tb in t.values() for v in tb["VS prev Month (%)"]]
    assert set(res.merchants[heading]["VS prev Month (%)"]) == {"N/A"}   # no merchant table before


def test_read_drive_facts_reads_a_notebook_report_through_the_dashboard(tmp_path):
    from integrations.google_drive import DriveFile

    class Session:
        def file(self, kind, name):
            return tmp_path / name

        def remove(self, path):
            pass

    class Client:
        def download(self, file_id, mime="", http=None, as_xlsx=False):
            return (tmp_path / "source.xlsx").read_bytes()

    source = tmp_path / "source.xlsx"
    _notebook_workbook(source)
    period = e.make_period("month", "2026-08-01")
    record = rh.Record("Google Drive", "BOA", "month", period.start, "title", "BOA_report_August_2026.xlsx",
                       "xlsx", notebook=True, file_id="FID")
    facts = rh.read_drive_facts(Client(), record, Session())
    assert len(facts) == 1 and facts[0].kind == "month"
    assert facts[0].counts["processed"]["RWF"] == 10
    assert facts[0].counts["received"]["RWF"] == 7
