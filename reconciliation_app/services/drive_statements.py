"""
BANK STATEMENTS READ FROM GOOGLE DRIVE (Ecobank, 26/10/2026)
============================================================
The Ecobank "IREMBO PAY" exports are kept on Drive (Ecobank Recorn → Statements → <month> folders). In step 2 the
app reads them itself: no list to choose from. Each file's days come from INSIDE the file (its Request Dates, as for
an uploaded file), never from its name; the files that cover a day of the period are added.

    - only the files whose name starts with "IREMBO PAY" are considered (TACCT, IREMBO_MONTHLY_TRAN, ... are not)
    - at most 10 files are downloaded per second (RATE)
    - the days of each file are remembered on this computer (output/.drive_statements.pkl, by file ID and last
      change on Drive): the next time only the new or changed files, and the files of the period, are downloaded
    - nothing is written on Drive; the content is kept in memory only

    pick(client, folder_id, bank, start, end, cache_path, done=None) -> Picked
"""
from __future__ import annotations

import pickle
import threading
import time
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd

import engine
from integrations.google_drive import DriveError

PREFIX = "IREMBO PAY"
RATE = 10                      # files downloaded per second, at most
CACHE_VERSION = 1


@dataclass
class Picked:
    files: list = field(default_factory=list)        # [(name, content)] of the files that cover the period
    considered: int = 0                              # "IREMBO PAY" files on Drive
    downloaded: int = 0
    unreadable: list = field(default_factory=list)   # names of the files that are not a statement
    errors: list = field(default_factory=list)       # downloads that failed (plain words)


def candidates(files: list) -> list:
    """The Drive files to consider: the "IREMBO PAY" exports (any case), not the folders' other files."""
    return [f for f in files if f.name.strip().upper().startswith(PREFIX)]


def _load(path: Path) -> dict:
    try:
        with open(path, "rb") as fh:
            saved = pickle.load(fh)
        return saved["files"] if saved.get("version") == CACHE_VERSION else {}
    except Exception:
        return {}


def _save(path: Path, data: dict) -> None:
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        tmp = path.with_suffix(".tmp")
        with open(tmp, "wb") as fh:
            pickle.dump({"version": CACHE_VERSION, "files": data}, fh)
        tmp.replace(path)
    except OSError:
        pass


class _Limiter:
    """Lets at most `rate` downloads start per second (shared by the threads)."""

    def __init__(self, rate: int):
        self.gap, self.next, self.lock = 1.0 / rate, time.monotonic(), threading.Lock()

    def wait(self):
        with self.lock:
            now = time.monotonic()
            start = max(now, self.next)
            self.next = start + self.gap
        time.sleep(max(0.0, start - now))


def _overlaps(first, last, start, end) -> bool:
    return first is not None and last is not None and first <= end and last >= start


def pick(client, folder_id: str, bank, start, end, cache_path: Path, done=None) -> Picked:
    """The statements on Drive that hold a day from start to end (both included), read from inside each file.
    done(n, total): called after each download."""
    bank = engine.get_bank(bank)
    start, end = pd.Timestamp(start).normalize(), pd.Timestamp(end).normalize()
    files = candidates(client.list_files(folder_id, bank.FILE_TYPES))
    known = _load(cache_path)
    out = Picked(considered=len(files))
    # a file not seen yet (or changed on Drive) must be opened to learn its days; a known one only if it covers
    # the period
    todo = []
    for f in files:
        k = known.get(f.id)
        if k is None or k["modified"] != f.modified:
            todo.append(f)
        elif _overlaps(k["first"], k["last"], start, end):
            todo.append(f)
    limiter = _Limiter(RATE)
    takes_http = client.thread_http() is not None
    local = threading.local()                    # one connection per thread, kept for its next downloads

    def one(f):
        limiter.wait()
        if takes_http and getattr(local, "http", None) is None:
            local.http = client.thread_http()
        try:
            return f, client.download(f.id, f.mime, http=local.http if takes_http else None), None
        except DriveError as err:
            return f, None, str(err)

    # a download takes ~1.5 s: 2 x RATE at the same time, so that RATE can start every second
    with ThreadPoolExecutor(max_workers=2 * RATE if takes_http else 1) as pool:
        for i, (f, content, err) in enumerate(pool.map(one, todo), 1):
            if done:
                done(i, len(todo))
            if err:
                out.errors.append(f"{f.name}: {err}")
                continue
            out.downloaded += 1
            try:
                read = engine.read_bank_file(f.name, content, bank)
                first, last = read.first, read.last
            except Exception:
                first = last = None
                out.unreadable.append(f.name)
            known[f.id] = {"modified": f.modified, "name": f.name, "first": first, "last": last}
            if _overlaps(first, last, start, end):
                out.files.append((f.name, content, first))
    _save(cache_path, {k: v for k, v in known.items() if k in {f.id for f in files}})
    out.files = [(n, c) for n, c, _ in sorted(out.files, key=lambda x: (x[2], x[0]))]
    return out
