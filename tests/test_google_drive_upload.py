"""DriveClient.upload(): the usual checksum-confirmed upload, and as_sheet=True (a report's .xlsx stored as a
native Google Sheet instead - approved 2026-10-26)."""
import hashlib

import pytest

from integrations.google_drive import DriveClient, DriveError, SHEET_MIME, XLSX_MIME


class FakeRequest:
    def __init__(self, result):
        self.result = result

    def execute(self, http=None, num_retries=2):
        return self.result


class FakeFiles:
    def __init__(self, stored):
        self.stored = stored          # what Drive claims to have stored (a dict, as the API would return)
        self.calls = []

    def create(self, body, media_body, fields, **kw):
        self.calls.append(("create", body))
        return FakeRequest(self.stored)

    def update(self, fileId, media_body, fields, body=None, **kw):
        self.calls.append(("update", fileId, body))
        return FakeRequest(self.stored)


class FakeService:
    def __init__(self, stored):
        self._files = FakeFiles(stored)

    def files(self):
        return self._files


def client_with(stored):
    service = FakeService(stored)
    return DriveClient(service, mode="oauth"), service


def test_a_plain_upload_is_confirmed_by_size_and_checksum():
    data = b"some report bytes"
    stored = {"id": "F1", "name": "report.xlsx", "mimeType": XLSX_MIME, "size": str(len(data)),
             "md5Checksum": hashlib.md5(data).hexdigest()}
    client, service = client_with(stored)
    result = client.upload("parent", "report.xlsx", data, XLSX_MIME)
    assert result.id == "F1"
    assert service._files.calls[0] == ("create", {"name": "report.xlsx", "parents": ["parent"]})


def test_a_plain_upload_with_a_mismatched_checksum_is_an_error():
    data = b"some report bytes"
    stored = {"id": "F1", "name": "report.xlsx", "mimeType": XLSX_MIME, "size": "1", "md5Checksum": "wrong"}
    client, _ = client_with(stored)
    with pytest.raises(DriveError, match="different content"):
        client.upload("parent", "report.xlsx", data, XLSX_MIME)


def test_as_sheet_asks_drive_to_convert_and_is_confirmed_by_the_stored_mime_type():
    data = b"xlsx bytes, not checked byte for byte once converted"
    stored = {"id": "F2", "name": "report.xlsx", "mimeType": SHEET_MIME}   # no size/md5: a native Sheet has none
    client, service = client_with(stored)
    result = client.upload("parent", "report.xlsx", data, XLSX_MIME, as_sheet=True)
    assert result.id == "F2" and result.mime == SHEET_MIME
    assert service._files.calls[0] == ("create", {"mimeType": SHEET_MIME, "name": "report.xlsx", "parents": ["parent"]})


def test_as_sheet_still_fails_when_drive_did_not_convert_it():
    stored = {"id": "F3", "name": "report.xlsx", "mimeType": XLSX_MIME}    # Drive kept it as plain Excel
    client, _ = client_with(stored)
    with pytest.raises(DriveError, match="not store this as a Google Sheet"):
        client.upload("parent", "report.xlsx", b"data", XLSX_MIME, as_sheet=True)


def test_as_sheet_on_replace_sends_the_mime_type_in_the_update_body():
    stored = {"id": "F4", "name": "report.xlsx", "mimeType": SHEET_MIME}
    client, service = client_with(stored)
    client.upload("parent", "report.xlsx", b"data", XLSX_MIME, replace_id="OLD", as_sheet=True)
    assert service._files.calls[0] == ("update", "OLD", {"mimeType": SHEET_MIME})
