"""
ONE SUMMARY FILE PER BANK ON GOOGLE DRIVE (13/11/2026; 18/11/2026: the only summary)
==================================================================================
Reading a report for the Reports Dashboard means downloading its Excel and opening its large sheets (a BK month has
about 80,000 rows): seconds per report. So each bank has one file holding the dashboard's figures of all its weekly
and monthly reports, in the bank's folder (at the top of Reconciliation Reports when the bank has no folder yet):

    <CODE>_dashboard_summary.json       {"version", "reports": {"<report's Drive ID>|<last modified>": [facts]}}

Saving a report adds its figures to its bank's file; the dashboard reads the banks' files (one download per bank)
and adds any report missing from them (read once from its Excel). The file is read again just before it is
written, so a report added meanwhile by someone else is kept. A bank keeps one file: copies made before (when the
file could not be replaced) are merged into it and moved to the Drive bin.

Who keeps each bank's file (KEEPERS, 18/11/2026): the app may change only the files its own account made (Drive
permission "drive.file"), so two people writing the same bank's file made a copy each time. Only the keeper's app
writes a bank's file (with every copy's figures); anyone else's app writes nothing there and moves its own old
copies to the bin; a report it saves is added by the keeper's dashboard (read once from the Excel). A bank not in
KEEPERS: anyone's app writes it. 18/11/2026: no summary file per report
any more (<report>.summary.json, 11/11/2026): one more file for every report saved.

    bank_name(code)                     the bank's file name
    bank_from_bytes(data)               {report key: [facts as saved]} ({} when unreadable or of another VERSION)
    facts_of(saved, record)             the facts of one report, as saved, with their report
    update_bank(client, folder, ...)    add reports to the bank's file, one file kept; never raises
    trash_old(client, files)            the old per-report summaries to the bin (the dashboard tidies Drive)
"""
from __future__ import annotations

import json
from concurrent.futures import ThreadPoolExecutor

import pandas as pd

from services import facts_cache

OLD_SUFFIX = ".summary.json"       # the summary per report, 11-18/11/2026 (no longer written)
BANK_SUFFIX = "_dashboard_summary.json"
MIME = "application/json"
# 18/11/2026: the Google account whose app keeps each bank's summary file (the one who reconciles the bank)
KEEPERS = {"BPR": "d.teta@irembo.com", "ECO": "d.teta@irembo.com",
           "BK": "d.ngenzi@irembo.com", "BOA": "d.ngenzi@irembo.com", "IM": "d.ngenzi@irembo.com"}


def keeps(client, code: str) -> bool:
    """This app's account keeps the bank's file (or the bank has no keeper). Unknown account: False (nothing written)."""
    keeper = KEEPERS.get(code)
    if not keeper:
        return True
    try:
        return client.account_email().lower() == keeper.lower()
    except Exception:
        return False


def _plain(v):
    """numpy numbers as Python numbers (json writes those only)."""
    return v.item() if hasattr(v, "item") else str(v)


def _fact(f) -> dict:
    return {"bank": f.bank, "kind": f.kind, "start": f"{f.start:%Y-%m-%d}", "end": f"{f.end:%Y-%m-%d}",
            "counts": f.counts, "amounts": f.amounts,
            "merchants": None if f.merchants is None else [[m, c, n, a] for (m, c), (n, a) in f.merchants.items()],
            "weekend": f.weekend, "weekend_known": f.weekend_known, "source": f.source}


def facts_of(saved: list, record) -> list | None:
    """The facts as they were saved (_fact), with `record` as their report; None when they cannot be read."""
    from services.dashboard import Fact
    try:
        return [Fact(f["bank"], f["kind"], pd.Timestamp(f["start"]), pd.Timestamp(f["end"]), f["counts"], f["amounts"],
                     None if f["merchants"] is None else {(m, c): (n, a) for m, c, n, a in f["merchants"]},
                     f["weekend"], f["weekend_known"], record, f["source"]) for f in saved]
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


def _each(client, items: list, job) -> list:
    """job(item, http) for every item, 4 at a time, each on a Drive connection of its own (18/11/2026: Drive answers
    in about a second; one after the other when the client has no connection to give, e.g. in tests)."""
    if len(items) < 2 or getattr(client, "thread_http", lambda: None)() is None:
        return [job(i, None) for i in items]
    with ThreadPoolExecutor(min(4, len(items))) as pool:
        return list(pool.map(lambda i: job(i, client.thread_http()), items))


def _call(fn, *args, http=None):
    return fn(*args, http=http) if http is not None else fn(*args)


def update_bank(client, folder_id: str, code: str, entries: dict, keep=None, known: dict | None = None,
                moved: list | None = None):
    """Add the reports {report key: [facts]} to the bank's file in the folder, on top of what is on Drive now (read
    again here: every copy of it, merged). keep: the keys of the bank's reports now on Drive (the others are
    dropped). The stored file (DriveFile), or None when nothing was written (the dashboard then reads the Excel).
    18/11/2026: only the bank's keeper (KEEPERS) writes it, into its own copy (replaced, tried twice; a new one only
    when it has none); every other copy of its own goes to the bin. Anyone else's app writes nothing and moves its
    own copies to the bin (their figures are in the keeper's file, or come back from the Excel). With no entries,
    the copies are only tidied; nothing is done when there is a single file. known: {file ID: content} already
    downloaded (the dashboard's), not downloaded again; moved: the copies moved to the bin are added to it."""
    if not folder_id:
        return None
    try:
        name = bank_name(code)
        files = sorted((f for f in client.list_children(folder_id, name=name) if f.name == name),
                       key=lambda f: f.modified)
        if not entries and len(files) < 2:                           # nothing to add, nothing to tidy
            return None
        ours = [f for f in files if f.mine]
        if not keeps(client, code):                                  # not ours to keep: only our copies away
            (moved if moved is not None else []).extend(_to_bin(client, ours, name))
            return None
        known = {k: v for k, v in (known or {}).items() if isinstance(v, (bytes, bytearray))}   # not a failed one
        get = [f for f in files if known.get(f.id) is None]
        got = dict(zip([f.id for f in get], _each(client, get, lambda f, h: _call(client.download, f.id, http=h))))
        reports = {}
        for f in files:                                              # every copy, whoever made it, oldest first
            reports.update(bank_from_bytes(known.get(f.id) if known.get(f.id) is not None else got[f.id]))
        reports.update({k: [_fact(f) for f in facts] for k, facts in entries.items()})
        if keep is not None:
            reports = {k: v for k, v in reports.items() if k in set(keep)}
        body = json.dumps({"version": facts_cache.VERSION, "reports": reports}, default=_plain,
                          ensure_ascii=False).encode("utf-8")
        stored = None
        for _ in range(2 if ours else 1):                           # a slow Drive: once more before anything else
            try:
                stored = client.upload(folder_id, name, body, MIME, replace_id=ours[-1].id if ours else None)
                break
            except Exception:
                pass
        if stored is None:
            return None
        (moved if moved is not None else []).extend(_to_bin(client, [f for f in ours if f.id != stored.id], name))
        return stored
    except Exception:
        return None


def _to_bin(client, files: list, what: str) -> list:
    """Move these files to the Drive bin, 4 at a time; one Drive refuses is left (tried again next time). The names
    of the files moved."""
    def one(f, http):
        try:
            _call(client.trash, f.id, f"an old copy of {what}", http=http)
            return f.name
        except Exception:
            return None
    return [n for n in _each(client, files, one) if n]


def trash_old(client, files: list) -> list:
    """Move the old per-report summaries (<report>.summary.json, 11-18/11/2026) to the Drive bin: their figures are in
    the banks' files. Only files of that name are touched; a file Drive refuses to move (another account's) is left
    (its own account's app moves it). The names moved."""
    old = [f for f in files if f.name.endswith(OLD_SUFFIX) and not f.name.endswith(BANK_SUFFIX) and f.mine]
    return [n for f, n in zip(old, _to_bin_each(client, old)) if n]


def _to_bin_each(client, files: list) -> list:
    """_to_bin for files of different names: [name or None] in the same order."""
    def one(f, http):
        try:
            _call(client.trash, f.id, f.name, http=http)
            return f.name
        except Exception:
            return None
    return _each(client, files, one)
