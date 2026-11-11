"""
ERRORS IN PLAIN WORDS (02/11/2026)
==================================
Any error shown on a page is said in a few words the user can act on, never as a traceback: the app's own messages
as they are, the usual problems (file missing, open in Excel, damaged, no internet) in plain words, anything else as
"<what was being done> did not work" with what to do, and the technical name kept short for the developer.
"""
from __future__ import annotations

import zipfile

import engine
from integrations.google_drive import DriveError
from services.app_settings import SettingsError


def message(err, doing: str = "This step") -> str:
    """The message shown for `err`, raised while `doing` (e.g. "The reconciliation", "Reading this report")."""
    if isinstance(err, (engine.ReconError, DriveError, SettingsError)):
        return str(err)
    if isinstance(err, FileNotFoundError):
        return f"{doing} did not work: a temporary file was missing. Please try again."
    if isinstance(err, PermissionError):
        return f"{doing} did not work: a file is open in another program (e.g. Excel). Close it and try again."
    if isinstance(err, zipfile.BadZipFile) or type(err).__name__ == "InvalidFileException":
        return f"{doing} did not work: the file is not a valid Excel file (it may be damaged). Check the file."
    if isinstance(err, (ConnectionError, TimeoutError)) or type(err).__name__ in ("ServerNotFoundError", "SSLError"):
        return f"{doing} did not work: Google or the internet could not be reached. Check the connection and try again."
    if isinstance(err, MemoryError):
        return f"{doing} did not work: the files are too large for this computer's memory."
    text = str(err).strip()
    detail = text.splitlines()[0][:120] if text else ""
    return (f"{doing} did not work. Please try again; if it happens again, tell the developer "
            f"({type(err).__name__}{': ' + detail if detail else ''}).")
