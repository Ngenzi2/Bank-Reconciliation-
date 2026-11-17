"""
A SMALL SUMMARY FILE NEXT TO EACH REPORT ON GOOGLE DRIVE (11/11/2026)
=====================================================================
Reading a report for the Reports Dashboard means downloading its Excel and opening its large sheets (a BK month has
about 80,000 rows): seconds per report. So next to each Excel report on Drive there is a file of a few KB holding the
dashboard's figures of that report, read from that very Excel (report_history.facts_from_xlsx: the same reading as
the dashboard, so the figures are the same):

    <CODE>_recon_summary_<period>.summary.json      next to <CODE>_recon_summary_<period>.xlsx

The dashboard downloads only the summary. A summary is used only when it belongs to the Excel now on Drive (the
Excel's MD5 checksum is written in it) and was written with the same VERSION as facts_cache; otherwise the Excel is
read as before and its summary written again (catch-up: the reports saved before this file existed). Each report
has its own summary: two people saving at the same time never write into the same file. Writing a summary never
fails a save or the dashboard: without it, the Excel is read.

    name_for(xlsx_name)                 the summary's name
    to_bytes(facts, xlsx_md5)           the summary of a report's facts (dashboard.Fact)
    from_bytes(data, record)            its facts again, or None when it does not belong to the Excel `record`
    write(client, folder, ...)          upload it (a new file, or replacing the one already there); never raises

ONE FILE PER BANK (13/11/2026)
With many reports the dashboard still downloads one small file per report. So each bank also has one file holding the
figures of all its weekly and monthly reports, in the bank's folder (at the top of Reconciliation Reports when the
bank has no folder yet):

    <CODE>_dashboard_summary.json       {"version", "reports": {"<report's Drive ID>|<last modified>": [facts]}}

The dashboard reads these files first (one download per bank). A report that is not in its bank's file, or was
replaced since, is read from its own summary (above), else from its Excel, and added to the bank's file. Saving a
report adds it too. The file is read again just before it is written, so a report added meanwhile by someone else
is kept; when two people still write at the same moment, the report left out is simply added the next time the
dashboard is opened. When the file on Drive cannot be replaced (made by another account), a second one is written
next to it: the dashboard reads every file of the bank, the newest last.

    bank_name(code)                     the bank's file name
    bank_from_bytes(data)               {report key: [facts as saved]} ({} when unreadable or of another VERSION)
    facts_of(saved, record)             the facts of one report, as saved, with their report
    update_bank(client, folder, ...)    add reports to the bank's file; never raises
"""
from __future__ import annotations

import json

import pandas as pd

from services import facts_cache

SUFFIX = ".summary.json"
BANK_SUFFIX = "_dashboard_summary.json"
MIME = "application/json"


def name_for(xlsx_name: str) -> str:
    """'IM_recon_summary_01-09-2026_to_30-09-2026.xlsx' -> 'IM_recon_summary_01-09-2026_to_30-09-2026.summary.json'."""
    return xlsx_name.rsplit(".", 1)[0] + SUFFIX


def _plain(v):
    """numpy numbers as Python numbers (json writes those only)."""
    return v.item() if hasattr(v, "item") else str(v)


def _fact(f) -> dict:
    return {"bank": f.bank, "kind": f.kind, "start": f"{f.start:%Y-%m-%d}", "end": f"{f.end:%Y-%m-%d}",
            "counts": f.counts, "amounts": f.amounts,
            "merchants": None if f.merchants is None else [[m, c, n, a] for (m, c), (n, a) in f.merchants.items()],
            "weekend": f.weekend, "weekend_known": f.weekend_known, "source": f.source}


def to_bytes(facts: list, xlsx_md5: str) -> bytes:
    return json.dumps({"version": facts_cache.VERSION, "xlsx_md5": xlsx_md5, "facts": [_fact(f) for f in facts]},
                      default=_plain, ensure_ascii=False).encode("utf-8")


def from_bytes(data: bytes, record) -> list | None:
    """The facts of the summary, with `record` (the Excel on Drive) as their report; None when the summary is
    unreadable, of another VERSION, or written for another content of the Excel (replaced by hand since)."""
    try:
        saved = json.loads(data)
        if saved.get("version") != facts_cache.VERSION or not record.md5 or saved.get("xlsx_md5") != record.md5:
            return None
        return facts_of(saved["facts"], record)
    except Exception:
        return None


def facts_of(saved: list, record) -> list | None:
    """The facts as they were saved (_fact), with `record` as their report; None when they cannot be read."""
    from services.dashboard import Fact
    try:
        return [Fact(f["bank"], f["kind"], pd.Timestamp(f["start"]), pd.Timestamp(f["end"]), f["counts"], f["amounts"],
                     None if f["merchants"] is None else {(m, c): (n, a) for m, c, n, a in f["merchants"]},
                     f["weekend"], f["weekend_known"], record, f["source"]) for f in saved]
    except Exception:
        return None


def write(client, folder_id: str, xlsx_name: str, facts: list, xlsx_md5: str, replace_id: str | None = None,
          http=None):
    """Upload the summary of the Excel `xlsx_name` into its folder (replace_id: the summary already there). The
    stored file (DriveFile), or None when it could not be written (the dashboard then reads the Excel)."""
    if not folder_id:
        return None
    try:
        return client.upload(folder_id, name_for(xlsx_name), to_bytes(facts, xlsx_md5), MIME, replace_id=replace_id,
                             http=http)
    except Exception:
        return None


# ---- one file per bank (13/11/2026) ---------------------------------------------------------------------------------
def bank_name(code: str) -> str:
    """'ECO' -> 'ECO_dashboard_summary.json'."""
    return f"{code}{BANK_SUFFIX}"


def bank_from_bytes(data) -> dict:
    """{report key (facts_cache.key): [facts as saved]} of a bank's file; {} when it is unreadable or was written
    with another VERSION."""
    try:
        saved = json.loads(data)
        return dict(saved["reports"]) if saved.get("version") == facts_cache.VERSION else {}
    except Exception:
        return {}


def update_bank(client, folder_id: str, code: str, entries: dict, keep=None):
    """Add the reports {report key: [facts]} to the bank's file in the folder, on top of what is on Drive now (read
    again here). keep: the keys of the bank's reports now on Drive (the others are dropped). The stored file
    (DriveFile), or None when it could not be written (the dashboard then reads each report's own summary)."""
    if not folder_id or not entries:
        return None
    try:
        name = bank_name(code)
        files = sorted((f for f in client.list_children(folder_id, name=name) if f.name == name),
                       key=lambda f: f.modified)
        reports = {}
        for f in files:
            reports.update(bank_from_bytes(client.download(f.id)))
        reports.update({k: [_fact(f) for f in facts] for k, facts in entries.items()})
        if keep is not None:
            reports = {k: v for k, v in reports.items() if k in set(keep)}
        body = json.dumps({"version": facts_cache.VERSION, "reports": reports}, default=_plain,
                          ensure_ascii=False).encode("utf-8")
        try:
            return client.upload(folder_id, name, body, MIME, replace_id=files[-1].id if files else None)
        except Exception:
            if not files:
                raise
            return client.upload(folder_id, name, body, MIME)       # not ours to replace: a file of our own
    except Exception:
        return None
