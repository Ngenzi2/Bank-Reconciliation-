"""
BANK STATEMENTS READ FROM GOOGLE DRIVE (Ecobank, 26/10/2026)
============================================================
The Ecobank "IREMBO PAY" exports are kept on Drive (Ecobank Recorn → Statements → <month> folders). In step 2 the
app reads them itself: no list to choose from. Each file's days come from INSIDE the file (its Request Dates, as for
an uploaded file), never from its name; the files that cover a day of the period are added.

    - only the files whose name starts with "IREMBO PAY" are considered (TACCT, IREMBO_MONTHLY_TRAN, ... are not)
    - at most 10 files are downloaded per second (RATE); a file not come after 3 s is asked again
    - the days of each file are remembered on this computer (output/.drive_statements.pkl, by file ID and last
      change on Drive): the next time only the new or changed files, and the files of the period, are downloaded
    - nothing is written on Drive; the content is kept in memory only

    pick(client, folder_id, bank, start, end, cache_path, done=None) -> Picked
"""
from __future__ import annotations

import pickle
import threading
import time
from concurrent.futures import FIRST_COMPLETED, ThreadPoolExecutor, wait
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd

import engine
from integrations.google_drive import DriveError

PREFIX = "IREMBO PAY"
RATE = 10                      # files downloaded per second, at most
TIMEOUT, TRIES, HEDGE = 20, 3, 3   # a request's time limit, requests per file at most, seconds before asking again
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


# 27/10/2026: kept in memory while the app runs (never on disk): a statement already read is not downloaded again
# (by file ID and last change on Drive), and the folder's list is kept LISTING_SECONDS (the ↻ icon reads it again)
_CONTENT: dict = {}
_LISTING: dict = {}
LISTING_SECONDS = 120


def _listing(client, folder_id: str, types, fresh: bool) -> list:
    hit = _LISTING.get(folder_id)
    if hit and not fresh and time.monotonic() - hit[0] < LISTING_SECONDS:
        return hit[1]
    files = client.list_files(folder_id, types)
    _LISTING[folder_id] = (time.monotonic(), files)
    return files


def pick(client, folder_id: str, bank, start, end, cache_path: Path, done=None, fresh: bool = False) -> Picked:
    """The statements on Drive that hold a day from start to end (both included), read from inside each file.
    done(n, total): called after each download. fresh: list the folder again (not the kept list)."""
    bank = engine.get_bank(bank)
    start, end = pd.Timestamp(start).normalize(), pd.Timestamp(end).normalize()
    files = candidates(_listing(client, folder_id, bank.FILE_TYPES, fresh))
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
    def connection():
        try:
            return client.thread_http(timeout=TIMEOUT)
        except TypeError:                         # a client without the timeout (older running app / test double)
            return client.thread_http()

    def attempt(f):
        kept = _CONTENT.get((f.id, f.modified))
        if kept is not None:                      # read before (this file, unchanged on Drive): no download
            return kept
        limiter.wait()                            # every request counts in the 10 per second
        data = client.download(f.id, f.mime, http=connection() if takes_http else None)
        _CONTENT[(f.id, f.modified)] = data
        return data

    def one(f):
        """27/10/2026: now and then Google is slow to answer one file (20-40 s). If it has not come after HEDGE
        seconds, the same file is asked again on a new connection (up to TRIES requests); the first answer is
        used, the slow one is left to end on its own."""
        if not takes_http:                        # one shared connection: one request at a time
            try:
                return f, attempt(f), None
            except DriveError as e:
                return f, None, str(e)
        pool = ThreadPoolExecutor(max_workers=TRIES)
        futs, started, err = [pool.submit(attempt, f)], 1, None
        try:
            while True:
                fin, _ = wait(futs, timeout=HEDGE, return_when=FIRST_COMPLETED)
                for fut in fin:
                    futs.remove(fut)
                    if fut.exception() is None:
                        return f, fut.result(), None
                    err = str(fut.exception())
                if started < TRIES:                # slow or failed: the same file asked again
                    futs.append(pool.submit(attempt, f))
                    started += 1
                elif not futs:                     # every request failed
                    return f, None, err
        finally:
            pool.shutdown(wait=False, cancel_futures=True)

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


# =====================================================================================================================
# 27/10/2026: READING IN THE BACKGROUND - the page is drawn at once and stays usable; the card shows the loader and
# the files appear when the reading is done (app.py: a small part of the page checks the job every half second)
# =====================================================================================================================
JOBS: dict = {}


def start_job(key, connect, folder_id: str, bank, start, end, cache_path: Path, fresh: bool = False) -> dict:
    """The job reading the statements for `key` (started now if there is none). connect(): a Drive client of the
    job's own (one connection cannot serve two threads). Returns {"done", "i", "n", "result", "error"}."""
    job = JOBS.get(key)
    if job is not None:
        return job
    job = {"done": False, "i": 0, "n": 0, "result": None, "error": None}
    JOBS[key] = job

    def run():
        try:
            job["result"] = pick(connect(), folder_id, bank, start, end, cache_path,
                                 done=lambda i, n: job.update(i=i, n=n), fresh=fresh)
        except Exception as err:                  # a Drive problem: said in the card
            job["error"] = str(err)
        finally:
            job["done"] = True
    threading.Thread(target=run, daemon=True, name=f"statements {key[0]}").start()
    return job
