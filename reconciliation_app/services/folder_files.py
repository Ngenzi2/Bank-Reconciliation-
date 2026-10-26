"""
FINDING FILES FOR A PERIOD (dianah branch: folders instead of uploads)
========================================================================
Given a bank's statements/exports folders (config.load_folders), finds the files a period needs by itself, the
way the team's notebooks do (read_with_cache): list the folder, read each file once, and keep what the period
needs. A folder can be local (any sub-folders allowed) or a Google Drive folder.

    find_files(period, bank, folders, drive=None) -> FoundFiles
        .statements   inside the period's usual window (the same window engine.accept_file uses)
        .reference    older statement files (back to engine.RESEND_LOOKBACK_DAYS), kept only to spot re-sends:
                      never required, never counted
        .exports      every IremboPay export file in the exports folder
        .problems     what looks wrong, in plain words (file of another bank, folder not found, ...); missing
                      days, a short payment file and a currency/stale-export warning are engine.check_period's job

    A local statements folder is also checked against config.DriveSettings.statements_root_folder_id, every time,
    when Drive is connected and that setting is filled in - the shared "Statements" folder the reconciliation
    notebooks already use (month folders, every bank's files together), the same way the notebooks pick up a new
    day's statement themselves. Nothing needs to be downloaded by hand first; a file already found locally is
    never replaced by its Drive copy, approved 2026-10-26.
    describe(found) -> ["Statements found: ...", "IremboPay export: ..."]   (the Check screen's file list)
    all_files(found) -> found.statements + found.reference + found.exports (what engine.check_period / run need)

Caching: one parsed engine.UploadedFile per (path, mtime, size), or (Drive file id, its modifiedTime); a changed
file is read again, nothing else is. The cache is kept here for as long as the app keeps running.
"""
from __future__ import annotations

from dataclasses import dataclass, field, replace
from pathlib import Path

import pandas as pd

import config
import engine

_CACHE: dict = {}


@dataclass
class FoundFiles:
    statements: list = field(default_factory=list)
    reference: list = field(default_factory=list)
    exports: list = field(default_factory=list)
    problems: list = field(default_factory=list)


def all_files(found: FoundFiles) -> list:
    return found.statements + found.reference + found.exports


def describe(found: FoundFiles) -> list:
    """Plain-language lines for the Check screen, e.g. 'Statements found: 1 Oct - 7 Oct, 7 files (2 from Drive)'."""
    lines = []
    if found.statements:
        first, last = min(f.first for f in found.statements), max(f.last for f in found.statements)
        n = len(found.statements)
        from_drive = sum(1 for f in found.statements + found.reference if f.from_drive)
        extra = f" ({from_drive} from Drive)" if from_drive else ""
        lines.append(f"Statements found: {first:%d %b} – {last:%d %b}, {n} file{'s' if n != 1 else ''}{extra}")
    else:
        lines.append("Statements found: none")
    if found.exports:
        taken = [f.taken_at for f in found.exports if f.taken_at is not None]
        when = f" taken {max(taken):%d %b %H:%M}" if taken else ""
        n = len(found.exports)
        lines.append(f"IremboPay export:{when} ({n} file{'s' if n != 1 else ''})")
    else:
        lines.append("IremboPay export: none found")
    return lines


def _other_bank_name(name: str, content: bytes, bank) -> str | None:
    """Which other active bank this file actually looks like, or None."""
    for other in engine.BANKS.values():
        if other is bank:
            continue
        try:
            other.read_statement(content, name, engine.EXCEL_ENGINE)
            return other.NAME
        except Exception:
            continue
    return None


# =====================================================================================================================
# LOCAL FOLDERS
# =====================================================================================================================
def _local_candidates(folder: Path, extensions) -> list:
    if not folder.is_dir():
        return []
    exts = {str(x).lower().lstrip(".") for x in extensions}
    return [p for p in sorted(folder.rglob("*"))
            if p.is_file() and not p.name.startswith(("~$", ".")) and p.suffix.lower().lstrip(".") in exts]


def _read_bank_cached(path: Path, bank):
    st = path.stat()
    key = ("bank", bank.CODE, str(path), st.st_mtime, st.st_size)
    if key in _CACHE:
        return _CACHE[key], None
    content = path.read_bytes()
    try:
        f = engine.read_bank_file(path.name, content, bank)
    except engine.ReconError as err:
        other = _other_bank_name(path.name, content, bank)
        return None, (f"{path.name}: this is a {other} statement, not {bank.NAME}." if other else str(err))
    _CACHE[key] = f
    return f, None


def _read_export_cached(path: Path, bank):
    st = path.stat()
    key = ("pay", bank.CODE, str(path), st.st_mtime, st.st_size)
    if key in _CACHE:
        return _CACHE[key], None
    content = path.read_bytes()
    try:
        f = engine.read_payment_file(path.name, content, bank)
    except engine.ReconError as err:
        return None, str(err)
    f = replace(f, order=st.st_mtime, taken_at=pd.Timestamp(st.st_mtime, unit="s"))
    _CACHE[key] = f
    return f, None


def _local_statements(folder: str, bank, problems: list) -> list:
    path = Path(folder).expanduser()
    if not path.is_dir():
        problems.append(f"The {bank.NAME} statements folder was not found: {path}")
        return []
    out = []
    for p in _local_candidates(path, bank.FILE_TYPES):
        f, problem = _read_bank_cached(p, bank)
        (problems.append(problem) if problem else out.append(f))
    return out


def _local_exports(folder: str, bank, problems: list) -> list:
    path = Path(folder).expanduser()
    if not path.is_dir():
        problems.append(f"The {bank.NAME} IremboPay exports folder was not found: {path}")
        return []
    types = engine.option(bank, "PAYMENT_FILE_TYPES", ["csv"])
    out = []
    for p in _local_candidates(path, types):
        f, problem = _read_export_cached(p, bank)
        (problems.append(problem) if problem else out.append(f))
    return out


# =====================================================================================================================
# GOOGLE DRIVE FOLDERS
# =====================================================================================================================
def _read_bank_cached_drive(df, drive, bank):
    key = ("bank", bank.CODE, df.id, df.modified)
    if key in _CACHE:
        return _CACHE[key], None
    content = drive.download(df.id)
    try:
        f = engine.read_bank_file(df.name, content, bank)
    except engine.ReconError as err:
        other = _other_bank_name(df.name, content, bank)
        return None, (f"{df.name}: this is a {other} statement, not {bank.NAME}." if other else str(err))
    _CACHE[key] = f
    return f, None


def _drive_modified(df) -> pd.Timestamp:
    taken = pd.to_datetime(df.modified, utc=True, errors="coerce")
    if pd.isna(taken):
        return pd.Timestamp.now()
    return taken.tz_localize(None)


def _read_export_cached_drive(df, drive, bank):
    key = ("pay", bank.CODE, df.id, df.modified)
    if key in _CACHE:
        return _CACHE[key], None
    content = drive.download(df.id)
    try:
        f = engine.read_payment_file(df.name, content, bank)
    except engine.ReconError as err:
        return None, str(err)
    taken = _drive_modified(df)
    f = replace(f, order=taken.timestamp(), taken_at=taken)
    _CACHE[key] = f
    return f, None


def _drive_statements(folder_id: str, bank, drive, problems: list) -> list:
    if drive is None:
        problems.append(f"The {bank.NAME} statements folder is a Google Drive folder, but Drive is not connected.")
        return []
    out = []
    for df in drive.list_files(folder_id, extensions=tuple(bank.FILE_TYPES)):
        f, problem = _read_bank_cached_drive(df, drive, bank)
        (problems.append(problem) if problem else out.append(f))
    return out


def _drive_exports(folder_id: str, bank, drive, problems: list) -> list:
    if drive is None:
        problems.append(f"The {bank.NAME} IremboPay exports folder is a Google Drive folder, but Drive is not "
                        "connected.")
        return []
    types = tuple(engine.option(bank, "PAYMENT_FILE_TYPES", ["csv"]))
    out = []
    for df in drive.list_files(folder_id, extensions=types):
        f, problem = _read_export_cached_drive(df, drive, bank)
        (problems.append(problem) if problem else out.append(f))
    return out


# =====================================================================================================================
# PUTTING IT TOGETHER: what a period needs
# =====================================================================================================================
def _window(period, bank):
    """(window_start, window_end): the days a statement file may cover for this period (engine.accept_file)."""
    window_start, window_end = period.day_before, engine.extra_monday(period, bank) or period.last_day
    earlier = engine.option(bank, "ACCEPT_EARLIER_DAYS", 0)
    if earlier:
        window_start = period.start - pd.Timedelta(days=earlier)
    return window_start, window_end


def _needed_days(period, bank) -> set:
    """Every day a period could use from a statement file: the usual window, plus the re-send lookback before it."""
    window_start, window_end = _window(period, bank)
    lookback_start = period.start - pd.Timedelta(days=engine.RESEND_LOOKBACK_DAYS)
    return set(pd.date_range(lookback_start, window_end))


def _month_folders(drive, root: str, days: set) -> list:
    """The shared Statements folder's own '<Month Year>' sub-folder ids touching `days` - so this only lists and
    downloads the few months actually needed, never the whole shared folder (every bank's files, years of
    them)."""
    names = sorted({d.strftime("%B %Y") for d in days})
    return [sub for sub in (drive.find_folder(root, name) for name in names) if sub]


def _sort_statements(candidates: list, period, bank) -> tuple[list, list]:
    """Candidates -> (statements inside the usual window, older ones kept only to spot re-sends)."""
    window_start, window_end = _window(period, bank)
    lookback_start = period.start - pd.Timedelta(days=engine.RESEND_LOOKBACK_DAYS)
    statements, reference, seen = [], [], set()
    for f in candidates:
        if f.digest in seen or f.last < lookback_start or f.first > window_end:
            continue
        seen.add(f.digest)
        if f.last < window_start:
            reference.append(replace(f, reference=True))
        else:
            statements.append(f)
    return statements, reference


def find_files(period, bank, folders: dict, drive=None) -> FoundFiles:
    """folders: config.load_folders()'s result, {bank code: {"statements": FolderSpec|None, "exports": ...}}.
    drive: an already-connected integrations.google_drive.DriveClient, when a folder is on Drive (None: Drive
    folders become a problem line instead of being read)."""
    bank = engine.get_bank(bank)
    spec = folders.get(bank.CODE) or {}
    found = FoundFiles()
    stm_spec, exp_spec = spec.get("statements"), spec.get("exports")

    if stm_spec is None:
        found.problems.append(f"No statements folder is set for {bank.NAME}.")
    else:
        candidates = (_local_statements(stm_spec.value, bank, found.problems) if stm_spec.kind == "local"
                     else _drive_statements(stm_spec.value, bank, drive, found.problems))
        if stm_spec.kind == "local" and drive is not None:
            root = config.load().statements_root_folder_id
            if root:                                      # checked every time, the way the notebooks do it
                have = {f.digest for f in candidates}
                for sub in _month_folders(drive, root, _needed_days(period, bank)):
                    candidates += [replace(f, from_drive=True) for f in _drive_statements(sub, bank, drive, found.problems)
                                  if f.digest not in have]
        found.statements, found.reference = _sort_statements(candidates, period, bank)

    if exp_spec is None:
        found.problems.append(f"No IremboPay exports folder is set for {bank.NAME}.")
    else:
        found.exports = (_local_exports(exp_spec.value, bank, found.problems) if exp_spec.kind == "local"
                        else _drive_exports(exp_spec.value, bank, drive, found.problems))
    return found
