"""parse_name() with no file extension (a native Google Sheet, Drive's own doing when a report is saved as one -
approved 2026-10-26), and save_report()'s dedup recognising such a file as the report's own Excel file, not a
stranger."""
import engine as e
from integrations.google_drive import DriveFile
from services import report_history as rh

WEEK = e.make_period("week", "2026-09-07")


def test_a_name_with_no_extension_is_read_as_the_excel_one():
    p = rh.parse_name("ECO_recon_summary_07-09-2026_to_13-09-2026")
    assert p is not None
    assert p["ext"] == "xlsx"
    assert p["stem"] == "ECO_recon_summary_07-09-2026_to_13-09-2026"
    assert p["code"] == "ECO"


def test_a_name_with_the_extension_still_parses_as_before():
    p = rh.parse_name("ECO_recon_summary_07-09-2026_to_13-09-2026.xlsx")
    assert p["ext"] == "xlsx"
    assert p["stem"] == "ECO_recon_summary_07-09-2026_to_13-09-2026"


class FakeClient:
    def __init__(self):
        self.uploads = []    # [(name, replace_id, as_sheet)]

    def find_or_create_folder(self, parent_id, name):
        return f"{parent_id}/{name}"

    def thread_http(self):
        return None

    def forget_folders(self):
        pass

    def upload(self, parent_id, name, data, mime, replace_id=None, http=None, as_sheet=False):
        self.uploads.append((name, replace_id, as_sheet))
        return DriveFile(id=f"new-{name}", name=name, mime=mime)


def test_saving_again_finds_an_earlier_sheet_saved_report_by_stem(tmp_path):
    # the Excel of an earlier save is already a Google Sheet (no extension); saving again in "new" mode must still
    # recognise it as the same report's Excel file, not try to save a duplicate next to it
    stem = rh.report_stem("ECO", WEEK)
    existing = [DriveFile("OLD1", stem, mime="application/vnd.google-apps.spreadsheet"),
               DriveFile("OLD2", f"{stem}.pdf", mime="application/pdf")]
    client = FakeClient()
    results = rh.save_report(client, "root", "ECO", WEEK, {"xlsx": b"x", "pdf": b"p"}, mode="new", existing=existing)
    assert not any(r.ok for r in results)
    assert all("already exists" in r.error for r in results)
    assert client.uploads == []                                # nothing was actually sent to Drive


def test_find_previous_carries_the_mime_type_through():
    # so previous_file() (Compare with previous period) downloads a Sheet-saved report correctly instead of
    # trying get_media on it
    prev = WEEK.previous()
    label = rh.report_stem("ECO", prev).split(e.REPORT_SUFFIX)[1]
    record = rh.Record("Google Drive", "ECO", "week", prev.start, "title",
                       rh.report_stem("ECO", prev), "xlsx", label=label,
                       mime="application/vnd.google-apps.spreadsheet")
    index = rh.DriveIndex(records=[record], others=[], folders=[])
    found = rh.find_previous(None, "root", "ECO", prev, index=index)
    assert found is not None
    assert found.mime == "application/vnd.google-apps.spreadsheet"


def test_replacing_a_sheet_saved_report_passes_its_id_to_upload(tmp_path):
    stem = rh.report_stem("ECO", WEEK)
    existing = [DriveFile("OLD1", stem, mime="application/vnd.google-apps.spreadsheet"),
               DriveFile("OLD2", f"{stem}.pdf", mime="application/pdf")]
    client = FakeClient()
    results = rh.save_report(client, "root", "ECO", WEEK, {"xlsx": b"x", "pdf": b"p"}, mode="replace",
                             existing=existing)
    assert all(r.ok for r in results)
    assert (f"{stem}.xlsx", "OLD1", True) in client.uploads       # the xlsx replaces OLD1, as a Sheet
    assert (f"{stem}.pdf", "OLD2", False) in client.uploads       # the pdf replaces OLD2, never converted
