"""
THE DASHBOARD'S FIGURES, REMEMBERED ON THIS COMPUTER BETWEEN SESSIONS
=====================================================================
Reading a report from Google Drive takes seconds (download + open); its figures never change while the file is the
same. So the figures of each report are kept in output/.dashboard_facts.pkl (output/ stays on this computer, like
the review decisions), keyed by the report's Drive file ID and its last-modified date: a new or replaced report is
read again from Drive, the others are not. The figures are always those of the report on Drive.

    load(path)              {key: [facts]} (empty when there is no file, it cannot be read, or VERSION changed)
    save(path, entries)     add / replace entries and drop those of reports no longer on Drive (keep = their keys)
    key(record)             "<file id>|<modified>"

Change VERSION when the way a report's figures are read changes (services/dashboard.py, engine.read_report): every
report is then read again once.
"""
from __future__ import annotations

import dataclasses
import pickle
from pathlib import Path

VERSION = 2                     # 2: weekly reports and their days, the rules of 26/10/2026


def key(record) -> str:
    return f"{record.file_id}|{record.modified}"


def load(path: Path) -> dict:
    try:
        with open(path, "rb") as fh:
            saved = pickle.load(fh)
        return saved["facts"] if saved.get("version") == VERSION else {}
    except Exception:                                   # no file yet, or unreadable: the reports are read from Drive
        return {}


def save(path: Path, entries: dict, keep=None) -> None:
    """entries: {key: [facts]}; the facts are stored without their Drive record (it is attached again on load).
    keep: the keys of the reports now on Drive (the others are dropped). A failure to write is not an error: the
    dashboard still works, it just reads from Drive next time."""
    data = load(path)
    for k, facts in entries.items():
        data[k] = [dataclasses.replace(f, record=None) for f in facts]
    if keep is not None:
        data = {k: v for k, v in data.items() if k in set(keep)}
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        tmp = path.with_suffix(".tmp")
        with open(tmp, "wb") as fh:
            pickle.dump({"version": VERSION, "facts": data}, fh)
        tmp.replace(path)
    except OSError:
        pass


def attach(facts: list, record) -> list:
    """The stored facts with the current Drive record of their report."""
    return [dataclasses.replace(f, record=record) for f in facts]
