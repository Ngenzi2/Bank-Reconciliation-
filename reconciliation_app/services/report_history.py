"""
SAVED REPORTS: GOOGLE DRIVE FIRST
=================================
Google Drive is the permanent storage of the standard reports (Excel + PDF). This file saves them there and finds
them again; it uses integrations/google_drive.py for every Drive call and engine.py for names and titles.

    On Drive:   <Reconciliation Reports>/<bank full name>/<YYYY-MM>/<CODE>_recon_summary_<first>_to_<last>[_vN].xlsx|pdf
                YYYY-MM = the month the period starts in; _vN = a new version saved next to an earlier one

    A report the reconciliation notebooks made (not this app) is recognised too, wherever it is: their own folders
    (<bank name>/Weekly or Monthly) and their own naming (parse_notebook_name), including a native Google Sheet
    with no file extension. It is read for the Reports Dashboard (engine.read_notebook_report); "Compare with
    previous period" (find_previous) does not offer one yet - it needs a report in this app's own layout.

    existing_reports()   the files already saved for a report (duplicate check before saving)
    save_report()        upload Excel + PDF (at the same time): as new files, replacing the existing ones, or as a
                         new version
    remember_saved()     add what was just saved to a listing already made (Drive is not listed again)
    drive_months()       the month folders of a bank          drive_records()  the reports of a bank and month
    local_records()      older reports still in output/ (listed only; never copied)
    find_previous()      the latest saved report of a period (for "Compare with previous period")
    read_drive_summary() the figures of a saved report (for the Reports Dashboard; no PDF is made)
"""
from __future__ import annotations

import calendar
import re
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass
from pathlib import Path

import pandas as pd

import engine
import pdf_report
from integrations.google_drive import DriveError, DriveFile

MIME = {"xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "pdf": "application/pdf"}
REPORTS_FOLDER = "Reconciliation Reports"           # the name shown to the user for the root folder
NAME_RE = re.compile(r"^(?P<code>[A-Z0-9]+)" + re.escape(engine.REPORT_SUFFIX) +
                     r"(?P<a>\d\d-\d\d-\d{4})_to_(?P<b>\d\d-\d\d-\d{4})(?: \(week (?P<w>\d+)\))?(?:_v(?P<v>\d+))?"
                     r"(?: \((?P<copy>\d+)\))?\.(?P<ext>xlsx|pdf)$", re.IGNORECASE)   # ' (1)': a copy saved by hand

# ---- the reconciliation notebooks' own naming (CLAUDE.md's reference notebooks), read alongside the app's own:
#      "<Bank>_report_<Month>_<Year>" (monthly) or "<Bank>_report_<Ordinal>_week_of_<Month>[_<Year>]" (weekly,
#      the year sometimes left out); the bank is matched by its Drive folder, not by this (sometimes inconsistent)
#      text. A native Google Sheet (no file extension) is as valid a name as a real .xlsx.
_NB_ORDINALS = {"first": 1, "second": 2, "third": 3, "fourth": 4, "fifth": 5}
_NB_MONTHS = {name.lower(): i for i, name in enumerate(calendar.month_name) if name}
_NB_WEEK_RE = re.compile(r"^(?P<ord>first|second|third|fourth|fifth)_week_of_(?P<month>[a-z]+)"
                         r"(?:_(?P<year>\d{4}))?$", re.IGNORECASE)
_NB_MONTH_RE = re.compile(r"^(?P<month>[a-z]+)_(?P<year>\d{4})$", re.IGNORECASE)


# ---- where a report goes ---------------------------------------------------------------------------------------------
def bank_folder(bank) -> str:
    """'I&M Bank', 'Ecobank': the bank's folder on Drive (its full name, from banks/)."""
    return engine.get_bank(bank).FULL_NAME


def month_folder(period: engine.Period) -> str:
    """'2026-09': the month of the period - for a week, the month it belongs to (most of its days: 28/09-04/10/2026
    is October Week 1, in 2026-10; 27/10/2026), else the month it starts in."""
    if period.report_type == "week":
        return f"{engine.week_number(period.start)[0]:%Y-%m}"
    return f"{period.start:%Y-%m}"


def drive_location(bank, period) -> str:
    return f"{REPORTS_FOLDER}/{bank_folder(bank)}/{month_folder(period)}"


def report_stem(bank, period) -> str:
    """'IM_recon_summary_01-09-2026_to_30-09-2026' (the same name as the standard report)."""
    return engine.report_path(Path("."), period, bank).stem


def parse_name(name: str):
    """'IM_recon_summary_01-09-2026_to_30-09-2026_v2.pdf' -> dict(code, kind, start, label, version, ext), or None."""
    m = NAME_RE.match(name)
    if not m:
        return None
    a, b = (pd.to_datetime(m[x], format="%d-%m-%Y") for x in ("a", "b"))
    days = (b - a).days                   # 0 / 6 = last day in the name; 1 / 7 = older names (the day after)
    kind = "day" if days in (0, 1) else ("week" if days in (6, 7) and a.dayofweek == 0 else "month")
    return {"code": m["code"].upper(), "kind": kind, "start": a, "label": f"{m['a']}_to_{m['b']}",
            "version": int(m["v"] or 1), "ext": m["ext"].lower(), "stem": name.rsplit(".", 1)[0]}


def _nth_week_of_month(ordinal: int, month: int, year: int) -> pd.Timestamp | None:
    """The Monday of the Nth week that belongs to <month> <year> (a week belongs to the month of its Thursday) -
    the inverse of the notebooks' own week naming, e.g. 'First week of August 2026'."""
    probe = pd.Timestamp(year=year, month=month, day=1) - pd.Timedelta(days=13)   # a week before, well clear
    monday = probe - pd.Timedelta(days=probe.weekday())
    weeks = []
    for _ in range(10):                                   # a month touches at most 6 weeks
        if engine.month_of("week", monday) == f"{year}-{month:02d}":
            weeks.append(monday)
        monday += pd.Timedelta(days=7)
    return weeks[ordinal - 1] if 0 < ordinal <= len(weeks) else None


def _notebook_year(modified, month: int) -> int:
    """A notebook file name with no year ('...of September.xlsx'): the year of the month nearest the file's own
    modified date (the notebooks only ever name the current or a recent past period)."""
    when = pd.to_datetime(modified, errors="coerce", utc=True)
    if pd.isna(when):
        return pd.Timestamp.now().year
    return when.year - 1 if month > when.month + 1 else when.year


def parse_notebook_name(name: str, folder: str, modified=None):
    """'BPR_report_First_week_of_August_2026' in folder 'BPR / Weekly' -> the same shape as parse_name(), plus
    "notebook": True - the reconciliation notebooks' own naming, not the app's own."""
    parts = [p.strip() for p in folder.split("/")]
    if len(parts) < 2 or parts[-1] not in ("Weekly", "Monthly"):
        return None
    bank_name, kind_folder = parts[-2], parts[-1]
    code = next((b.CODE for b in engine.BANKS.values() if b.NAME.lower() == bank_name.lower()), None)
    if code is None:
        return None
    stem = Path(name).stem if Path(name).suffix.lower() in (".xlsx", ".xls") else name
    stem = re.sub(r"\s*\(\d+\)$", "", stem)                        # a Drive copy: "... (1)"
    m = re.match(rf"^{re.escape(bank_name)}_report_(?P<rest>.+)$", stem, re.IGNORECASE)
    if not m:
        return None
    rest = m["rest"]
    if kind_folder == "Monthly":
        mm = _NB_MONTH_RE.match(rest)
        if not mm or mm["month"].lower() not in _NB_MONTHS:
            return None
        start = pd.Timestamp(year=int(mm["year"]), month=_NB_MONTHS[mm["month"].lower()], day=1)
        kind = "month"
    else:
        wm = _NB_WEEK_RE.match(rest)
        if not wm or wm["month"].lower() not in _NB_MONTHS:
            return None
        month = _NB_MONTHS[wm["month"].lower()]
        year = int(wm["year"]) if wm["year"] else _notebook_year(modified, month)
        start = _nth_week_of_month(_NB_ORDINALS[wm["ord"].lower()], month, year)
        if start is None:
            return None
        kind = "week"
    period = engine.make_period(kind, start)
    label = f"{period.start:%d-%m-%Y}_to_{period.last_day:%d-%m-%Y}"
    return {"code": code, "kind": kind, "start": period.start, "label": label, "version": 1, "ext": "xlsx",
           "stem": stem, "notebook": True}


# ---- saving ----------------------------------------------------------------------------------------------------------
@dataclass
class SaveResult:
    name: str
    ok: bool                         # True only when Google Drive confirmed the stored file (size + checksum)
    file: DriveFile | None = None
    error: str = ""
    replaced: bool = False


def _month_folder_id(client, root_id, bank, period, create: bool):
    bank_id = (client.find_or_create_folder if create else client.find_folder)(root_id, bank_folder(bank))
    if not bank_id:
        return None
    return (client.find_or_create_folder if create else client.find_folder)(bank_id, month_folder(period))


def existing_reports(client, root_id: str, bank, period) -> list:
    """Every file already saved for this report on Drive (all versions, Excel and PDF), newest first."""
    folder = _month_folder_id(client, root_id, bank, period, create=False)
    if not folder:
        return []
    stem = report_stem(bank, period)
    return [f for f in client.list_children(folder)
            if (p := parse_name(f.name)) and p["label"] == period.label and p["code"]
            == engine.get_bank(bank).CODE]


def next_version(existing: list) -> int:
    """The next free version number (2 when only the first report exists)."""
    return max([parse_name(f.name)["version"] for f in existing if parse_name(f.name)] + [1]) + 1


def save_report(client, root_id: str, bank, period, files: dict, mode: str = "new", existing: list | None = None,
                only: set | None = None) -> list:
    """Upload the report files {ext: bytes} (ext = 'xlsx' / 'pdf') to <bank>/<YYYY-MM>/.
    mode: 'new'      first save (no file of this report on Drive yet)
          'replace'  replace the content of the existing files of the same name (Drive keeps their history)
          'version'  save next to them as _v<N>
    only: the names to (re)try (a retry after a failure). Returns one SaveResult per file; ok only when Drive
    confirmed it. Nothing is deleted on Drive."""
    stem = report_stem(bank, period)
    existing = existing or []
    if mode == "version":
        stem = f"{stem}_v{next_version(existing)}"
    names = {ext: f"{stem}.{ext}" for ext in files}
    targets = [(ext, name) for ext, name in names.items() if only is None or name in only]
    try:
        folder = _month_folder_id(client, root_id, bank, period, create=True)
    except DriveError as err:
        return [SaveResult(name, False, error=str(err)) for _, name in targets]
    by_name = {f.name: f for f in existing}

    def one(ext, name, http):
        replace = by_name.get(name) if mode == "replace" else None
        if mode == "new" and name in by_name:          # never overwrite without being asked
            return SaveResult(name, False, error="a file with this name already exists on Drive")
        try:
            stored = client.upload(folder, name, files[ext], MIME[ext], replace_id=replace.id if replace else None,
                                   http=http)
            return SaveResult(name, True, stored, replaced=replace is not None)
        except DriveError as err:
            return SaveResult(name, False, error=str(err))

    # the Excel and the PDF at the same time, each on a connection of its own (one after the other when there is
    # no connection to copy, e.g. in tests)
    https = [client.thread_http() for _ in targets] if len(targets) > 1 else []
    if https and all(h is not None for h in https):
        with ThreadPoolExecutor(len(targets)) as pool:
            results = list(pool.map(lambda job: one(*job[0], job[1]), zip(targets, https)))
    else:
        results = [one(ext, name, None) for ext, name in targets]
    if not all(r.ok for r in results):
        client.forget_folders()                        # a folder may have been removed: look it up again next time
    return results


# ---- finding saved reports -------------------------------------------------------------------------------------------
@dataclass
class Record:
    """One saved report file, from Drive or from this computer."""
    source: str                      # "Google Drive" or "This computer"
    bank: str                        # bank code ('IM')
    kind: str                        # day / week / month
    start: pd.Timestamp
    title: str                       # 'Month · September 2026'
    name: str
    ext: str                         # xlsx / pdf
    version: int = 1
    size: int | None = None
    modified: pd.Timestamp | None = None
    file_id: str = ""                # Drive
    path: Path | None = None         # this computer
    folder: str = ""                 # where it is: 'I&M Bank / 2026-09' ('' = the top of Reconciliation Reports)
    stem: str = ""                   # name without extension: the Excel and the PDF of one report share it
    mime: str = ""                   # Drive's own mimeType ('': this computer, or the app's own .xlsx/.pdf upload)
    notebook: bool = False           # made by the reconciliation notebooks, not this app (engine.read_notebook_report)

    @property
    def month(self) -> str:
        """'2026-09': the month the period starts in (from the name, wherever the file is)."""
        return f"{self.start:%Y-%m}"


@dataclass
class DriveIndex:
    """Everything inside Reconciliation Reports (listed once; nothing downloaded)."""
    records: list                    # the reports, recognised by their name wherever they are
    others: list                     # every other file (DriveFile), shown as it is
    folders: list                    # every folder path ('I&M Bank', 'I&M Bank / 2026-09', ...)


def drive_index(client, root_id: str, depth: int = 3) -> DriveIndex:
    """List Reconciliation Reports once: the files at the top and in its folders (up to `depth` levels). A report is
    recognised by its name (<CODE>_recon_summary_<first>_to_<last>[_vN].xlsx|pdf), wherever it is: in the app's
    <bank>/<YYYY-MM> folders or put anywhere by hand - or by the reconciliation notebooks' own naming
    (parse_notebook_name), in their <bank name>/Weekly|Monthly folders. Nothing is moved or downloaded."""
    records, others = [], []
    files, folders = client.list_tree(root_id, depth)              # one request per level, folders included
    for f in sorted(files, key=lambda f: f.modified, reverse=True):
        p = parse_name(f.name) or parse_notebook_name(f.name, f.folder, f.modified)
        if p:
            records.append(_record(f, p))
        else:
            others.append(f)
    return DriveIndex(_sorted(records), others, sorted(set(folders)))


def _record(f: DriveFile, p: dict) -> Record:
    """A report file on Drive as a Record (p = parse_name or parse_notebook_name of its name)."""
    return Record("Google Drive", p["code"], p["kind"], p["start"], engine.period_title(p["kind"], p["label"]), f.name,
                  p["ext"], p["version"], f.size, pd.to_datetime(f.modified, errors="coerce", utc=True), file_id=f.id,
                  folder=f.folder, stem=p["stem"], mime=f.mime, notebook=bool(p.get("notebook")))


def remember_saved(index: DriveIndex, results: list, bank, period) -> DriveIndex:
    """The listing with the files just saved (SaveResult confirmed by Drive) added - a replaced file keeps its ID
    and takes its new size and date - so Report History, Compare and the Dashboard need no new listing."""
    folder = f"{bank_folder(bank)} / {month_folder(period)}"
    saved = {r.file.id: r.file for r in results if r.ok and r.file is not None}
    records = [r for r in index.records if r.file_id not in saved]
    for f in saved.values():
        p = parse_name(f.name)
        if p:
            records.append(_record(DriveFile(**{**f.__dict__, "folder": folder}), p))
    folders = sorted(set(index.folders) | {bank_folder(bank), folder})
    return DriveIndex(_sorted(records), index.others, folders)


def filter_records(records: list, bank: str | None = None, month: str | None = None, query: str = "") -> list:
    """The reports of one bank (code) and/or one month ('2026-09'), matching the search text."""
    return _filtered([r for r in records if (not bank or r.bank == bank) and (not month or r.month == month)], query)


def drive_months(client, root_id: str, bank) -> list:
    """The months ('2026-09', ...) that have reports of a bank on Drive, newest first (from the file names)."""
    code = engine.get_bank(bank).CODE
    return sorted({r.month for r in drive_index(client, root_id).records if r.bank == code}, reverse=True)


def drive_records(client, root_id: str, bank, month: str | None = None, query: str = "") -> list:
    """The reports of a bank on Drive (one month, or every month), newest period first, wherever they are inside
    Reconciliation Reports. Only names are listed: nothing is downloaded."""
    return filter_records(drive_index(client, root_id).records, engine.get_bank(bank).CODE, month, query)


def local_records(base_dir: Path, bank, query: str = "") -> list:
    """Older reports still in output/ on this computer (listed for compatibility; never copied anywhere)."""
    out = []
    for title, path in engine.report_list(base_dir, bank):
        p = parse_name(path.name)
        if p:
            out.append(Record("This computer", p["code"], p["kind"], p["start"], title, path.name, p["ext"],
                              p["version"], path.stat().st_size, pd.Timestamp(path.stat().st_mtime, unit="s"),
                              path=path, folder=path.parent.name, stem=p["stem"]))     # folder inside output/
    return _filtered(out, query)


def _sorted(records: list) -> list:
    return sorted(records, key=lambda r: (r.start, r.kind, r.version, r.ext), reverse=True)


def _filtered(records: list, query: str) -> list:
    q = query.strip().lower()
    return _sorted([r for r in records if not q or q in f"{r.title} {r.name} {r.kind} {r.folder}".lower()])


def find_previous(client, root_id: str, bank, period, index: DriveIndex | None = None) -> DriveFile | None:
    """The Excel of the report of `period` on Drive, wherever it is inside Reconciliation Reports: this app's own
    report when there is one, else the reconciliation notebooks' own (engine.read_report_counts reads either) -
    the newest version; at equal standing the newest saved. `index`: a listing already made (Drive is not listed
    again). Matched by kind and start date, not the re-parsed name: a notebook record's name is not this app's own
    (parse_name)."""
    code = engine.get_bank(bank).CODE
    hits = [r for r in (index or drive_index(client, root_id)).records
            if r.bank == code and r.ext == "xlsx" and r.kind == period.report_type and r.start == period.start]
    if not hits:
        return None
    best = max(hits, key=lambda r: (not r.notebook, r.version,
                                    r.modified if r.modified is not None else pd.Timestamp(0, tz="UTC")))
    return DriveFile(best.file_id, best.name, size=best.size, folder=best.folder, mime=best.mime)


def read_drive_summary(client, file_id: str, name: str, session) -> dict:
    """Only engine.read_report(...) of a saved Excel report (for the dashboard: no PDF is made). The file exists in
    the session's temporary folder only while it is read."""
    data = client.download(file_id)
    p = parse_name(name)
    if not p or p["ext"] != "xlsx":
        raise DriveError(f"{name}: only an Excel report can be read.")
    path = session.file(p["kind"], f"{p['code']}{engine.REPORT_SUFFIX}{p['label']}.xlsx")
    try:
        path.write_bytes(data)
        return engine.read_report(path)
    finally:
        session.remove(path)


def read_drive_facts(client, record, session, data: bytes | None = None) -> list:
    """The dashboard's facts of one saved Excel report (services/dashboard.py): its summary blocks and the late
    payments made over the weekend and booked Monday (from its Payment details). Works for a report this app made
    or one the reconciliation notebooks made (record.notebook), read by engine.read_notebook_report instead - its
    own name is not re-parsed, so a notebook report (which parse_name cannot read) still works. The file exists in
    the session's temporary folder only while it is read. data: the file already downloaded (download_many), else
    it is."""
    from services import dashboard
    data = client.download(record.file_id, record.mime, as_xlsx=True) if data is None else data
    if record.ext != "xlsx":
        raise DriveError(f"{record.name}: only an Excel report can be read.")
    period = engine.make_period(record.kind, record.start)
    path = session.file(record.kind, f"{record.bank}{engine.REPORT_SUFFIX}{period.label}.xlsx")
    try:
        path.write_bytes(data)
        no_tx = dashboard.no_transaction_fact(path, record)          # a No-Transaction Report: every count 0
        if no_tx is not None:
            return [no_tx]
        rep = (engine.read_notebook_report(path, record.bank, period) if record.notebook
              else engine.read_report(path, count_reviewed=False))  # the dashboard does not use that count
        try:
            sheets = pd.read_excel(path, sheet_name=["Payment details", "Pending notification",
                                                     "Duplicate payment other channel"], engine=engine.EXCEL_ENGINE)
            details = sheets["Payment details"]
            weekend = dashboard.weekend_late(details, record.bank)
        except Exception:                                  # an older report without these sheets: no days
            sheets, weekend = None, None
        facts = dashboard.facts_from_report(rep, record, weekend)
        if sheets is not None and not record.notebook:      # the days, from its rows (checked against it): of a
            whole = "month" if record.kind == "month" else "week"   # monthly report, or of a weekly one (27/10/2026)
            for month in [f for f in facts if f.kind == whole]:     # not known for the notebooks' own reports
                facts += dashboard.daily_facts(month, sheets["Payment details"], sheets["Pending notification"],
                                               sheets["Duplicate payment other channel"], record.bank, weekend)
        return facts
    finally:
        session.remove(path)


def download_many(client, records: list, workers: int = 4, done=None) -> dict:
    """{file_id: bytes or DriveError} of several reports, downloaded at the same time (each on a connection of its
    own; one after the other when the client has none to give). done(n): called after each file."""
    from concurrent.futures import ThreadPoolExecutor, as_completed

    def one(rec):
        try:
            return rec.file_id, client.download(rec.file_id, rec.mime, http=client.thread_http(), as_xlsx=True)
        except DriveError as err:
            return rec.file_id, err

    import inspect
    out = {}
    # a client made before download() took http (a running app keeps its Drive client until it restarts): one by one
    takes_http = "http" in inspect.signature(client.download).parameters
    parallel = len(records) > 1 and takes_http and client.thread_http() is not None
    if not parallel:
        for i, rec in enumerate(records, 1):
            try:
                out[rec.file_id] = client.download(rec.file_id, rec.mime, as_xlsx=True)
            except DriveError as err:
                out[rec.file_id] = err
            if done:
                done(i)
        return out
    with ThreadPoolExecutor(max_workers=workers) as pool:
        for i, fut in enumerate(as_completed([pool.submit(one, rec) for rec in records]), 1):
            fid, data = fut.result()
            out[fid] = data
            if done:
                done(i)
    return out


def read_drive_report(client, file_id: str, name: str, session, mime: str = "") -> dict:
    """A saved report read back from Drive (for the viewer): {"report": engine.read_report(...), "xlsx": bytes,
    "pdf": bytes}. The file exists in the session's temporary folder only while it is read."""
    data = client.download(file_id, mime, as_xlsx=True)
    p = parse_name(name)
    if not p or p["ext"] != "xlsx":
        raise DriveError(f"{name}: only the Excel report can be opened in the app.")
    path = session.file(p["kind"], f"{p['code']}{engine.REPORT_SUFFIX}{p['label']}.xlsx")
    try:
        path.write_bytes(data)
        return {"report": engine.read_report(path), "xlsx": data, "pdf": pdf_report.report_pdf(path)}
    finally:
        session.remove(path)
