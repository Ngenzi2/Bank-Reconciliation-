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
"""
from __future__ import annotations

import json

import pandas as pd

from services import facts_cache

SUFFIX = ".summary.json"
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
    from services.dashboard import Fact
    try:
        saved = json.loads(data)
        if saved.get("version") != facts_cache.VERSION or not record.md5 or saved.get("xlsx_md5") != record.md5:
            return None
        return [Fact(f["bank"], f["kind"], pd.Timestamp(f["start"]), pd.Timestamp(f["end"]), f["counts"], f["amounts"],
                     None if f["merchants"] is None else {(m, c): (n, a) for m, c, n, a in f["merchants"]},
                     f["weekend"], f["weekend_known"], record, f["source"]) for f in saved["facts"]]
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
