"""
SAVED REPORTS: GOOGLE DRIVE FIRST
=================================
Google Drive is the permanent storage of the standard reports (Excel + PDF). This file saves them there and finds
them again; it uses integrations/google_drive.py for every Drive call and engine.py for names and titles.

    On Drive:   <Reconciliation Reports>/<top>/<second>/<CODE>_recon_summary_<first>_to_<last>[_vN].xlsx|pdf
                <top>/<second> depends on the report_layout setting (config.py): "notebook" (default) is
                <bank name>/<Weekly|Monthly>, the same folders the reconciliation notebooks use; "app" is
                <bank full name>/<YYYY-MM>. _vN = a new version saved next to an earlier one. A report is found
                again by its file name alone, so reading (Dashboard, Compare with previous period) works for
                either layout without knowing the setting.

    existing_reports()   the files already saved for a report (duplicate check before saving)
    save_report()        upload Excel + PDF (at the same time): as new files, replacing the existing ones, or as a
                         new version
    remember_saved()     add what was just saved to a listing already made (Drive is not listed again)
    drive_months()       the month folders of a bank          drive_records()  the reports of a bank and month
    local_records()      older reports still in output/ (listed only; never copied)
    find_previous()      the latest saved report of a period (for "Compare with previous period")
    read_drive_facts()   the Dashboard's figures of a report read from Drive
    read_local_facts()   the same, for a report kept in output/ (the Dashboard falls back to this when Drive is
                         not connected, and always merges the two so a report is never missed)
"""
from __future__ import annotations

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
                     r"(?P<a>\d\d-\d\d-\d{4})_to_(?P<b>\d\d-\d\d-\d{4})(?:_v(?P<v>\d+))?"
                     r"(?: \((?P<copy>\d+)\))?\.(?P<ext>xlsx|pdf)$", re.IGNORECASE)   # ' (1)': a copy saved by hand


# ---- where a report goes ---------------------------------------------------------------------------------------------
def bank_folder(bank) -> str:
    """'I&M Bank', 'Ecobank': the bank's folder on Drive (its full name, from banks/)."""
    return engine.get_bank(bank).FULL_NAME


def month_folder(period: engine.Period) -> str:
    """'2026-09': the month the period starts in."""
    return f"{period.start:%Y-%m}"


def _kind_folder(period: engine.Period) -> str:
    """'Weekly' or 'Monthly' - the notebooks' own second folder (a day uses Weekly, the closest fit: the
    notebooks only make weekly and monthly reports)."""
    return "Monthly" if period.report_type == "month" else "Weekly"


def location_parts(bank, period, layout: str) -> tuple:
    """(top folder, second folder) of a report, for either layout:
        "notebook"  <bank name>/<Weekly|Monthly>/            - what the reconciliation notebooks already use
        "app"       <bank full name>/<YYYY-MM>/              - the app's own first layout
    Reading (drive_index, find_previous, the Dashboard) does not need this: a report is recognised by its file
    name alone, wherever it is saved."""
    if layout == "notebook":
        return engine.get_bank(bank).NAME, _kind_folder(period)
    return bank_folder(bank), month_folder(period)


def drive_location(bank, period, layout: str = "notebook") -> str:
    top, second = location_parts(bank, period, layout)
    return f"{REPORTS_FOLDER}/{top}/{second}"


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


# ---- saving ----------------------------------------------------------------------------------------------------------
@dataclass
class SaveResult:
    name: str
    ok: bool                         # True only when Google Drive confirmed the stored file (size + checksum)
    file: DriveFile | None = None
    error: str = ""
    replaced: bool = False


def _report_folder_id(client, root_id, bank, period, layout, create: bool):
    top, second = location_parts(bank, period, layout)
    top_id = (client.find_or_create_folder if create else client.find_folder)(root_id, top)
    if not top_id:
        return None
    return (client.find_or_create_folder if create else client.find_folder)(top_id, second)


def existing_reports(client, root_id: str, bank, period, layout: str = "notebook") -> list:
    """Every file already saved for this report on Drive in this layout's folder (all versions, Excel and PDF),
    newest first."""
    folder = _report_folder_id(client, root_id, bank, period, layout, create=False)
    if not folder:
        return []
    stem = report_stem(bank, period)
    return [f for f in client.list_children(folder)
            if (p := parse_name(f.name)) and p["label"] == stem.split(engine.REPORT_SUFFIX)[1] and p["code"]
            == engine.get_bank(bank).CODE]


def next_version(existing: list) -> int:
    """The next free version number (2 when only the first report exists)."""
    return max([parse_name(f.name)["version"] for f in existing if parse_name(f.name)] + [1]) + 1


def save_report(client, root_id: str, bank, period, files: dict, mode: str = "new", existing: list | None = None,
                only: set | None = None, layout: str = "notebook") -> list:
    """Upload the report files {ext: bytes} (ext = 'xlsx' / 'pdf') to the folder this layout gives (see
    location_parts()).
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
        folder = _report_folder_id(client, root_id, bank, period, layout, create=True)
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
    <bank>/<YYYY-MM> folders or put anywhere by hand. Nothing is moved or downloaded."""
    records, others = [], []
    files, folders = client.list_tree(root_id, depth)              # one request per level, folders included
    for f in sorted(files, key=lambda f: f.modified, reverse=True):
        p = parse_name(f.name)
        if p:
            records.append(_record(f, p))
        else:
            others.append(f)
    return DriveIndex(_sorted(records), others, sorted(set(folders)))


def _record(f: DriveFile, p: dict) -> Record:
    """A report file on Drive as a Record (p = parse_name of its name)."""
    return Record("Google Drive", p["code"], p["kind"], p["start"], engine.period_title(p["kind"], p["label"]), f.name,
                  p["ext"], p["version"], f.size, pd.to_datetime(f.modified, errors="coerce", utc=True), file_id=f.id,
                  folder=f.folder, stem=p["stem"])


def remember_saved(index: DriveIndex, results: list, bank, period, layout: str = "notebook") -> DriveIndex:
    """The listing with the files just saved (SaveResult confirmed by Drive) added - a replaced file keeps its ID
    and takes its new size and date - so Report History, Compare and the Dashboard need no new listing."""
    top, second = location_parts(bank, period, layout)
    folder = f"{top} / {second}"
    saved = {r.file.id: r.file for r in results if r.ok and r.file is not None}
    records = [r for r in index.records if r.file_id not in saved]
    for f in saved.values():
        p = parse_name(f.name)
        if p:
            records.append(_record(DriveFile(**{**f.__dict__, "folder": folder}), p))
    folders = sorted(set(index.folders) | {top, folder})
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
    """The Excel of the report of `period` on Drive, wherever it is inside Reconciliation Reports (the newest
    version; at equal version the last saved), or None. `index`: a listing already made (Drive is not listed again)."""
    code = engine.get_bank(bank).CODE
    label = report_stem(bank, period).split(engine.REPORT_SUFFIX)[1]
    hits = [r for r in (index or drive_index(client, root_id)).records
            if r.bank == code and r.ext == "xlsx" and parse_name(r.name)["label"] == label]
    if not hits:
        return None
    best = max(hits, key=lambda r: (r.version, r.modified if r.modified is not None else pd.Timestamp(0, tz="UTC")))
    return DriveFile(best.file_id, best.name, size=best.size, folder=best.folder)


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


def _report_facts(path: Path, record) -> list:
    """Shared by read_drive_facts and read_local_facts: the dashboard's facts of the Excel report at `path`,
    already identified as `record` (its summary blocks and the late payments made over the weekend and booked
    Monday, from its Payment details)."""
    from services import dashboard
    p = parse_name(record.name)
    if not p or p["ext"] != "xlsx":
        raise DriveError(f"{record.name}: only an Excel report can be read.")
    no_tx = dashboard.no_transaction_fact(path, record)          # a No-Transaction Report: every count 0
    if no_tx is not None:
        return [no_tx]
    rep = engine.read_report(path)
    try:
        sheets = pd.read_excel(path, sheet_name=["Payment details", "Pending notification",
                                                 "Duplicate payment other channel"], engine=engine.EXCEL_ENGINE)
        details = sheets["Payment details"]
        weekend = dashboard.weekend_late(details, p["code"])
    except Exception:                                  # an older report without these sheets: no days
        sheets, weekend = None, None
    facts = dashboard.facts_from_report(rep, record, weekend)
    if sheets is not None:                             # a month's days, from its rows (checked against it)
        for month in [f for f in facts if f.kind == "month"]:
            facts += dashboard.daily_facts(month, sheets["Payment details"], sheets["Pending notification"],
                                           sheets["Duplicate payment other channel"], p["code"], weekend)
    return facts


def read_drive_facts(client, record, session) -> list:
    """The dashboard's facts of one saved Excel report, read from Google Drive. The file exists in the
    session's temporary folder only while it is read."""
    data = client.download(record.file_id)
    p = parse_name(record.name)
    if not p or p["ext"] != "xlsx":
        raise DriveError(f"{record.name}: only an Excel report can be read.")
    path = session.file(p["kind"], f"{p['code']}{engine.REPORT_SUFFIX}{p['label']}.xlsx")
    try:
        path.write_bytes(data)
        return _report_facts(path, record)
    finally:
        session.remove(path)


def read_local_facts(record) -> list:
    """The dashboard's facts of one report already saved in output/ on this computer (local_records()): read in
    place, nothing downloaded."""
    return _report_facts(record.path, record)


def read_drive_report(client, file_id: str, name: str, session) -> dict:
    """A saved report read back from Drive (for the viewer): {"report": engine.read_report(...), "xlsx": bytes,
    "pdf": bytes}. The file exists in the session's temporary folder only while it is read."""
    data = client.download(file_id)
    p = parse_name(name)
    if not p or p["ext"] != "xlsx":
        raise DriveError(f"{name}: only the Excel report can be opened in the app.")
    path = session.file(p["kind"], f"{p['code']}{engine.REPORT_SUFFIX}{p['label']}.xlsx")
    try:
        path.write_bytes(data)
        return {"report": engine.read_report(path), "xlsx": data, "pdf": pdf_report.report_pdf(path)}
    finally:
        session.remove(path)
