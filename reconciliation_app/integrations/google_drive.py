"""
GOOGLE DRIVE
============
Every Google Drive operation of the app goes through this file: signing in, listing a folder, downloading a file,
finding or creating a folder, uploading a file, checking the access rights. The rest of the app never calls the
Google API directly.

Signing in (config.py, auth):
    oauth            (default) as you: sign_in() opens Google's consent page in the browser of this computer once;
                     the token is saved outside the project (token_file, readable only by you) and renewed
                     automatically. Permissions asked: only OAUTH_SCOPES (see below). sign_out() removes it.
    service_account  the service-account key (credentials_file): reads shared folders; it cannot create files in a
                     My Drive folder (no storage of its own), so it needs a Shared Drive to save reports.

    connect(settings)       a client signed in the configured way (NotSignedIn when OAuth has no valid sign-in)
    check_access(...)       what the account may really do in a folder (asked to Drive, never assumed)
    upload_test(folder)     upload a small test file, verify it, read it back, move it to the bin
    list_files(...)         files of a folder and of its subfolders (up to 3 levels, e.g. bank / month)
    list_tree(...)          files and folders of a folder tree, one request per level (the folders of a level
                            are asked about together: about 0.5 s a request, so this is what keeps listing fast)
    download(file_id)       the content of a file, in memory (a Google Sheet as CSV)
    find_folder / find_or_create_folder   (a folder found once is remembered for the session: forget_folders())
    upload(...)             create a file, or replace one; confirmed by Drive's checksum of the stored content
    thread_http()           a connection of its own for an upload made at the same time as another

The Google libraries are loaded only when Drive is used, so the app works without them when Drive is not set up.
Errors are DriveError, in plain words, with Google's own reason added; the secret files are never shown.
"""
from __future__ import annotations

import hashlib
import io
import json
import os
import uuid
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

# OAuth (default): the narrowest permissions that cover the app's needs
#   drive.readonly  list and download the IremboPay exports you put on Drive (files the app did not create)
#   drive.file      create, verify and replace the reports (and their folders) that the app itself creates
OAUTH_SCOPES = ["https://www.googleapis.com/auth/drive.readonly", "https://www.googleapis.com/auth/drive.file"]
SERVICE_ACCOUNT_SCOPES = ["https://www.googleapis.com/auth/drive"]   # limited in practice to the folders shared
FOLDER_MIME = "application/vnd.google-apps.folder"
SHEET_MIME = "application/vnd.google-apps.spreadsheet"     # a Google Sheet: downloaded as CSV (or as_xlsx=True)
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
# a file is of a type by its name (.csv) or, when the name has no extension, by its kind on Drive
TYPE_MIMES = {"csv": {"text/csv", SHEET_MIME},
              "xlsx": {"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"},
              "xls": {"application/vnd.ms-excel"}}
FILE_FIELDS = "id, name, mimeType, size, modifiedTime, md5Checksum, parents"
SHARED = {"supportsAllDrives": True}
SIGN_IN_AGAIN = "Sign in again (side panel → Google Drive → Sign in with Google)."


class DriveError(Exception):
    """A Google Drive problem explained in plain words."""


class NotSignedIn(DriveError):
    """OAuth: there is no valid sign-in (never signed in, signed out, expired or revoked)."""


@dataclass
class DriveFile:
    id: str
    name: str
    mime: str = ""
    size: int | None = None
    modified: str = ""                       # RFC 3339, e.g. '2026-10-09T08:15:00.000Z'
    md5: str = ""
    folder: str = ""                         # path of the folder it was found in, e.g. 'I&M / 09/2026' ("" = top)

    @classmethod
    def from_api(cls, f: dict, folder: str = "") -> "DriveFile":
        return cls(f["id"], f.get("name", ""), f.get("mimeType", ""),
                   int(f["size"]) if f.get("size") is not None else None, f.get("modifiedTime", ""),
                   f.get("md5Checksum", ""), folder)


@dataclass
class AccessCheck:
    """The answer of check_access() / upload_test(): ok = all good; lines = what was found, for the user."""
    ok: bool
    lines: list = field(default_factory=list)          # (ok: bool, text)


def _quote(text: str) -> str:
    """A name inside a Drive query: quotes and backslashes escaped."""
    return text.replace("\\", "\\\\").replace("'", "\\'")


def _google_reason(err) -> str:
    """Google's own reason of an HTTP error, e.g. 'insufficientFilePermissions: ...' ('' when there is none)."""
    try:
        error = json.loads(err.content.decode("utf-8"))["error"]
        detail = (error.get("errors") or [{}])[0]
        reason = detail.get("reason") or error.get("status") or ""
        return f"{reason}: {error.get('message', '')}".strip(": ")
    except Exception:
        return ""


# ---- OAuth: signing in as yourself ----------------------------------------------------------------------------------
def _write_private(path: Path, text: str):
    """Write a file only this computer's user can read (folder 0700, file 0600), replacing it in one step."""
    path.parent.mkdir(mode=0o700, parents=True, exist_ok=True)
    tmp = path.with_name(f".{path.name}.{uuid.uuid4().hex[:6]}.tmp")
    fd = os.open(tmp, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600)
    with os.fdopen(fd, "w") as f:
        f.write(text)
    os.replace(tmp, path)
    os.chmod(path, 0o600)


def _granted(scope) -> set:
    return set(scope.split()) if isinstance(scope, str) else set(scope or [])


def load_oauth_credentials(settings):
    """Your saved sign-in, renewed when it has expired. NotSignedIn when there is none or Google refuses it."""
    token = settings.token_file
    if not token.is_file():
        raise NotSignedIn("Not signed in to Google Drive. " + SIGN_IN_AGAIN.replace("Sign in again", "Sign in"))
    try:
        from google.auth.exceptions import RefreshError, TransportError
        from google.auth.transport.requests import Request
        from google.oauth2.credentials import Credentials
    except ImportError:
        raise DriveError("The Google sign-in libraries are not installed. Run 'Setup (run once)' again.") from None
    try:
        info = json.loads(token.read_text())
        creds = Credentials.from_authorized_user_info(info, OAUTH_SCOPES)
    except Exception:                                   # never repeat the token's content
        raise NotSignedIn("The saved Google sign-in could not be read. " + SIGN_IN_AGAIN) from None
    missing = set(OAUTH_SCOPES) - _granted(info.get("scopes"))
    if missing:
        raise NotSignedIn("The saved Google sign-in does not include the permissions the app needs ("
                          + ", ".join(sorted(s.rsplit("/", 1)[-1] for s in missing)) + "). " + SIGN_IN_AGAIN)
    if not creds.valid:
        if not creds.refresh_token:
            raise NotSignedIn("The Google sign-in has expired. " + SIGN_IN_AGAIN)
        try:
            creds.refresh(Request())
        except RefreshError:
            raise NotSignedIn("Your Google sign-in has expired or was revoked. " + SIGN_IN_AGAIN) from None
        except TransportError:
            raise DriveError("Google could not be reached to renew the sign-in. Check the internet connection.") \
                from None
        _write_private(token, json.dumps({**json.loads(creds.to_json()), "scopes": sorted(_granted(info["scopes"]))}))
    return creds


def sign_in(settings, timeout: int = 300, flow=None) -> str:
    """Open Google's consent page in this computer's browser, wait for your answer (up to `timeout` seconds), save
    the token (readable only by you) and return the signed-in e-mail address. The token is saved only when every
    permission in OAUTH_SCOPES was granted. flow: an already prepared sign-in flow (normally None)."""
    if flow is None:
        client = settings.oauth_client_file
        if not client.is_file():
            raise DriveError(f"No OAuth client file at {client}.")
        try:
            from google_auth_oauthlib.flow import InstalledAppFlow
        except ImportError:
            raise DriveError("The Google sign-in libraries are not installed. Run 'Setup (run once)' again.") \
                from None
        os.environ.setdefault("OAUTHLIB_RELAX_TOKEN_SCOPE", "1")   # unticked permissions: checked below instead
        try:
            flow = InstalledAppFlow.from_client_secrets_file(str(client), OAUTH_SCOPES)
        except Exception:
            raise DriveError(f"The OAuth client file at {client} could not be used. Check that it is the "
                             "'Desktop app' client JSON downloaded from Google Cloud.") from None
    try:
        creds = flow.run_local_server(
            port=0, open_browser=True, timeout_seconds=timeout,
            authorization_prompt_message="Opening Google sign-in in your browser…",
            success_message="Signed in. You can close this tab and go back to the Bank Reconciliation app.")
    except Exception as err:
        if type(err).__name__ in ("WSGITimeoutError", "AttributeError"):     # no answer before the timeout
            raise DriveError(f"No answer from the Google sign-in within {max(timeout // 60, 1)} minute(s). If the "
                             "browser showed 'Access blocked' or 'admin_policy_enforced', Irembo's Google Workspace "
                             "policy blocks the app: ask IT to trust the OAuth client.") from None
        raise DriveError(f"The Google sign-in did not complete ({type(err).__name__}: {err}).") from None   # e.g. Cancel
    if creds is None:
        raise DriveError("The Google sign-in did not complete.")
    token_scope = getattr(getattr(flow, "oauth2session", None), "token", {}) or {}
    granted = _granted(token_scope.get("scope")) or _granted(getattr(creds, "granted_scopes", None)) \
        or _granted(creds.scopes)
    missing = set(OAUTH_SCOPES) - granted
    if missing:
        raise DriveError("Google did not grant every permission the app needs ("
                         + ", ".join(sorted(s.rsplit("/", 1)[-1] for s in missing))
                         + "). Sign in again and allow all of them; nothing was saved.")
    _write_private(settings.token_file, json.dumps({**json.loads(creds.to_json()), "scopes": sorted(granted)}))
    return DriveClient(_build(creds), mode="oauth").account_email()


def sign_out(settings, revoke=None) -> bool:
    """Forget the sign-in: ask Google to revoke it (best effort), then delete the token file. Only that file."""
    token = settings.token_file
    if not token.is_file():
        return False
    try:
        info = json.loads(token.read_text())
        value = info.get("refresh_token") or info.get("token")
        if value:
            if revoke is None:
                import urllib.parse
                import urllib.request
                data = urllib.parse.urlencode({"token": value}).encode()
                urllib.request.urlopen(urllib.request.Request("https://oauth2.googleapis.com/revoke", data=data),
                                       timeout=15)
            else:
                revoke(value)
    except Exception:
        pass                                             # offline: the local token is deleted anyway
    token.unlink(missing_ok=True)
    return True


def _build(creds):
    try:
        from googleapiclient.discovery import build
    except ImportError:
        raise DriveError("The Google Drive libraries are not installed. Run 'Setup (run once)' again.") from None
    try:
        # 27/10/2026: a request may hang 30 s at most (Google's default is 60 s, x3 with the retries): now and then
        # Google is slow to answer, and the app waited minutes for one folder listing
        try:
            import google_auth_httplib2
            import httplib2
            return build("drive", "v3", http=google_auth_httplib2.AuthorizedHttp(creds, http=httplib2.Http(timeout=30)),
                         cache_discovery=False)
        except ImportError:
            return build("drive", "v3", credentials=creds, cache_discovery=False)
    except Exception as err:
        raise DriveError(f"Google Drive could not be reached ({type(err).__name__}).") from None


class DriveClient:
    def __init__(self, service, mode: str = "service_account"):
        self.service = service                         # a googleapiclient Drive v3 service (or a test double)
        self.mode = mode                               # "oauth" or "service_account"
        self._folders = {}                             # {(parent id, name): folder id} found this session

    # ---- connecting ----------------------------------------------------------------------------------------------
    @classmethod
    def connect(cls, settings) -> "DriveClient":
        """Signed in the configured way (config.DriveSettings.auth): OAuth (default) or the service account."""
        if settings.is_oauth:
            return cls(_build(load_oauth_credentials(settings)), mode="oauth")
        path = settings.credentials_file
        if not path.is_file():
            raise DriveError(f"No credentials file at {path}.")
        try:
            from google.oauth2 import service_account
        except ImportError:
            raise DriveError("The Google Drive libraries are not installed. Run 'Setup (run once)' again.") from None
        try:
            creds = service_account.Credentials.from_service_account_file(str(path), scopes=SERVICE_ACCOUNT_SCOPES)
        except Exception:                               # never repeat the file's content in the message
            raise DriveError(f"The credentials file at {path} could not be used. Check that it is the "
                             "service-account JSON key downloaded from Google Cloud.") from None
        return cls(_build(creds), mode="service_account")

    @property
    def who(self) -> str:
        return "your Google account" if self.mode == "oauth" else "the service account"

    def _run(self, request, what: str, http=None):
        """Execute one API request; HTTP, sign-in and network errors become a DriveError saying what failed, with
        Google's own reason when there is one. http: a connection of its own (thread_http) for a parallel request."""
        try:
            return request.execute(http=http, num_retries=2) if http is not None else request.execute(num_retries=2)
        except Exception as err:
            if type(err).__name__ == "RefreshError":    # OAuth sign-in expired or revoked during the session
                raise NotSignedIn("Your Google sign-in has expired or was revoked. " + SIGN_IN_AGAIN) from None
            status = getattr(getattr(err, "resp", None), "status", None)
            if status in (401, 403):
                reason = f"{self.who} has no permission for this"
            elif status == 404:
                reason = ("not found, or not available to this app" if self.mode == "oauth"
                          else "not found, or not shared with the service account")
            elif status is not None:
                reason = f"Google Drive answered {status}"
            else:
                reason = f"no answer from Google Drive ({type(err).__name__}); check the internet connection"
            google = _google_reason(err)
            raise DriveError(f"Could not {what}: {reason}." + (f" [Google: {google}]" if google else "")) from None

    def account_email(self) -> str:
        """The e-mail address Drive is used with (your account, or the service account)."""
        about = self._run(self.service.about().get(fields="user(emailAddress)"), "read the signed-in account")
        return (about.get("user") or {}).get("emailAddress", "")

    # ---- access rights (asked to Drive, never assumed) ------------------------------------------------------------
    def check_access(self, folder_id: str, label: str, write: bool) -> AccessCheck:
        """What the account may do in a folder. write=True: it must be able to add and replace files; with the
        service account the folder must also be in a Shared Drive (it has no storage of its own). With OAuth the
        reports are created in your own My Drive; the real proof of writing is upload_test()."""
        lines = []
        try:
            info = self._run(self.service.files().get(
                fileId=folder_id, fields="id, name, mimeType, driveId, capabilities", **SHARED),
                f"open the {label}")
        except DriveError as err:
            return AccessCheck(False, [(False, str(err))])
        caps = info.get("capabilities", {})
        is_folder = info.get("mimeType") == FOLDER_MIME
        lines.append((is_folder, f"{label}: '{info.get('name', '?')}' " + ("is a folder" if is_folder
                                                                           else "is not a folder")))
        lines.append((bool(caps.get("canListChildren")), "can list its files" if caps.get("canListChildren")
                      else f"cannot list its files (share the folder with {self.who})"))
        ok = is_folder and bool(caps.get("canListChildren"))
        if write:
            if self.mode != "oauth":
                in_shared = bool(info.get("driveId"))
                lines.append((in_shared, "is in a Shared Drive" if in_shared else
                              "is NOT in a Shared Drive: a service account cannot save files in My Drive"))
                ok = ok and in_shared
            for cap, text in (("canAddChildren", "add new reports"), ("canEdit", "replace existing reports")):
                lines.append((bool(caps.get(cap)), f"can {text}" if caps.get(cap) else
                              f"cannot {text} (give {self.who} the Editor / Content manager role)"))
            ok = ok and bool(caps.get("canAddChildren")) and bool(caps.get("canEdit"))
            if self.mode == "oauth":
                lines.append((True, "the real proof of saving is the Upload test"))
        else:
            lines.append((bool(caps.get("canDownload", True)), "can download its files"
                          if caps.get("canDownload", True) else "cannot download its files"))
            ok = ok and bool(caps.get("canDownload", True))
        return AccessCheck(ok, lines)

    def upload_test(self, folder_id: str) -> AccessCheck:
        """Upload a small test file with a unique name into the folder, check Drive's size and checksum, read it back
        and move only that file to the bin. If anything does not match, the file is kept for investigation."""
        name = f"_bank-reconciliation-upload-test-{datetime.now():%Y%m%d-%H%M%S}-{uuid.uuid4().hex[:6]}.txt"
        data = (f"Bank Reconciliation upload test\n{name}\nThis file is moved to the bin by the app right after "
                "the test. If you see it, the clean-up did not complete and it can be deleted.\n").encode()
        try:
            stored = self.upload(folder_id, name, data, "text/plain")
        except DriveError as err:
            return AccessCheck(False, [(False, f"Upload refused: {err}")])
        lines = [(True, f"uploaded {name} ({len(data)} bytes); Drive confirmed the same size and MD5 checksum")]
        try:
            back = self.download(stored.id)
        except DriveError as err:
            return AccessCheck(False, lines + [(False, f"Read-back failed: {err}. The test file is kept: {name}")])
        if back != data:
            return AccessCheck(False, lines + [(False, f"Read-back differs from what was sent. The test file is kept "
                                                       f"for investigation: {name}")])
        lines.append((True, "read back: identical to what was sent"))
        try:
            done = self._run(self.service.files().update(fileId=stored.id, body={"trashed": True},
                                                         fields="id, trashed", **SHARED), "move the test file to the bin")
            ok = bool(done.get("trashed"))
        except DriveError as err:
            return AccessCheck(False, lines + [(False, f"{err} Delete it by hand: {name}")])
        lines.append((ok, "test file moved to the bin (only this file)" if ok else
                       f"Drive did not confirm the bin; delete it by hand: {name}"))
        return AccessCheck(ok, lines)

    # ---- listing --------------------------------------------------------------------------------------------------
    def list_children(self, folder_id: str, folders_only=False, name: str | None = None) -> list:
        """The files (or the folders) directly in a folder, every page; name = only that exact name."""
        q = [f"'{_quote(folder_id)}' in parents", "trashed = false"]
        if folders_only:
            q.append(f"mimeType = '{FOLDER_MIME}'")
        if name is not None:
            q.append(f"name = '{_quote(name)}'")
        out, token = [], None
        while True:
            page = self._run(self.service.files().list(
                q=" and ".join(q), fields=f"nextPageToken, files({FILE_FIELDS})", pageSize=200, pageToken=token,
                orderBy="modifiedTime desc", corpora="allDrives", includeItemsFromAllDrives=True, **SHARED),
                "list the folder")
            out += [DriveFile.from_api(f) for f in page.get("files", [])]
            token = page.get("nextPageToken")
            if not token:
                return out

    def list_tree(self, folder_id: str, depth: int = 3) -> tuple[list, list]:
        """(files, folder paths) of a folder and of its subfolders, `depth` levels down. Each file's `folder` is its
        path below the folder ('I&M / 09/2026'); folder paths go down to `depth` levels. One request per level:
        the folders of a level are asked about together (in groups of 40, so the query stays short)."""
        files, folders, level = [], [], [(folder_id, "")]
        for depth_now in range(depth + 1):
            below = []
            for i in range(0, len(level), 40):
                paths = dict(level[i:i + 40])
                for f, parent in self._children_of(list(paths)):
                    path = paths[parent]
                    if f.mime == FOLDER_MIME:
                        if depth_now < depth:
                            here = f"{path} / {f.name}" if path else f.name
                            folders.append(here)
                            below.append((f.id, here))
                    else:
                        files.append(DriveFile(**{**f.__dict__, "folder": path}))
            if not below:
                break
            level = below
        return files, folders

    def _children_of(self, parent_ids: list) -> list:
        """[(file or folder, its parent among parent_ids)] of every item directly in those folders, every page."""
        wanted = set(parent_ids)
        q = "(" + " or ".join(f"'{_quote(fid)}' in parents" for fid in parent_ids) + ") and trashed = false"
        out, token = [], None
        while True:
            page = self._run(self.service.files().list(
                q=q, fields=f"nextPageToken, files({FILE_FIELDS})", pageSize=1000, pageToken=token,
                orderBy="modifiedTime desc", corpora="allDrives", includeItemsFromAllDrives=True, **SHARED),
                "list the folder")
            for f in page.get("files", []):
                parent = next((p for p in f.get("parents") or [] if p in wanted), parent_ids[0])
                out.append((DriveFile.from_api(f), parent))
            token = page.get("nextPageToken")
            if not token:
                return out

    def list_files(self, folder_id: str, extensions=(), depth: int = 3) -> list:
        """Files of a folder and of its subfolders, up to `depth` levels (e.g. bank / month), newest first; each
        file's `folder` is its path below the folder ('I&M / 09/2026'). extensions: e.g. ['csv'] (empty = all)."""
        found, _ = self.list_tree(folder_id, depth)
        if extensions:
            ends = tuple(f".{e.lower()}" for e in extensions)
            mimes = set().union(*(TYPE_MIMES.get(e.lower(), set()) for e in extensions))
            found = [f for f in found if f.name.lower().endswith(ends) or f.mime in mimes]
        return sorted(found, key=lambda f: f.modified, reverse=True)

    # ---- reading --------------------------------------------------------------------------------------------------
    def download(self, file_id: str, mime: str = "", http=None, as_xlsx: bool = False) -> bytes:
        """The content of a file, in memory (nothing is written on the computer). A Google Sheet (mime) is
        downloaded as CSV (its first sheet) by default - a payments or statement file only ever needs the one
        table. as_xlsx: export it as a proper .xlsx instead, every sheet kept - a saved report (several sheets:
        Summary, Payment details, ...) needs this, never CSV. http: a connection of its own (thread_http), to
        download several files at the same time."""
        if mime == SHEET_MIME:
            export_mime = XLSX_MIME if as_xlsx else "text/csv"
            data = self._run(self.service.files().export_media(fileId=file_id, mimeType=export_mime),
                             "download the Google Sheet", http=http)
            return data if isinstance(data, bytes) else bytes(data)
        data = self._run(self.service.files().get_media(fileId=file_id, **SHARED), "download the file", http=http)
        return data if isinstance(data, bytes) else bytes(data)

    # ---- folders --------------------------------------------------------------------------------------------------
    def find_folder(self, parent_id: str, name: str) -> str | None:
        """The ID of the folder `name` inside the parent, or None. A folder found is remembered for the session
        (each look-up is a request to Drive); forget_folders() clears that memory."""
        if (parent_id, name) in self._folders:
            return self._folders[(parent_id, name)]
        hits = self.list_children(parent_id, folders_only=True, name=name)
        if hits:
            self._folders[(parent_id, name)] = hits[0].id
        return hits[0].id if hits else None

    def forget_folders(self):
        """Look the folders up on Drive again next time (after a failed upload: a folder may have been removed)."""
        self._folders.clear()

    def find_or_create_folder(self, parent_id: str, name: str) -> str:
        """The ID of the folder `name` inside the parent, created when it does not exist yet."""
        found = self.find_folder(parent_id, name)
        if found:
            return found
        made = self._run(self.service.files().create(
            body={"name": name, "mimeType": FOLDER_MIME, "parents": [parent_id]}, fields="id", **SHARED),
            f"create the folder '{name}'")
        if not made.get("id"):
            raise DriveError(f"Could not create the folder '{name}': Google Drive did not confirm it.")
        self._folders[(parent_id, name)] = made["id"]
        return made["id"]

    # ---- writing --------------------------------------------------------------------------------------------------
    def thread_http(self, timeout: int = 120):
        """A connection of its own, signed in like the client, for a request made at the same time as another (one
        connection cannot carry two requests at once); None when there is none to copy (a test double).
        timeout: seconds a request may hang before it fails (short for small files: a stuck one is tried again)."""
        creds = getattr(getattr(self.service, "_http", None), "credentials", None)
        if creds is None:
            return None
        try:
            import google_auth_httplib2
            import httplib2
        except ImportError:
            return None
        return google_auth_httplib2.AuthorizedHttp(creds, http=httplib2.Http(timeout=timeout))

    def upload(self, parent_id: str, name: str, data: bytes, mime: str, replace_id: str | None = None,
               http=None) -> DriveFile:
        """Create the file `name` in the folder, or replace the content of the file replace_id. Returns the stored
        file only when Drive confirms it: same size and same MD5 checksum as `data`; otherwise DriveError (the
        stored file, if any, is left as it is for investigation)."""
        try:
            from googleapiclient.http import MediaIoBaseUpload
        except ImportError:
            raise DriveError("The Google Drive libraries are not installed. Run 'Setup (run once)' again.") from None
        media = MediaIoBaseUpload(io.BytesIO(data), mimetype=mime, resumable=False)
        if replace_id:
            request = self.service.files().update(fileId=replace_id, media_body=media, fields=FILE_FIELDS, **SHARED)
        else:
            request = self.service.files().create(body={"name": name, "parents": [parent_id]}, media_body=media,
                                                  fields=FILE_FIELDS, **SHARED)
        stored = DriveFile.from_api(self._run(request, f"upload {name}", http) or {"id": ""})
        if not stored.id:
            raise DriveError(f"{name}: Google Drive did not confirm the upload.")
        if stored.size != len(data) or stored.md5 != hashlib.md5(data).hexdigest():
            raise DriveError(f"{name}: Google Drive stored a different content than the one sent "
                             "(size or checksum differ). Try again.")
        return stored
