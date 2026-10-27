"""
BANK STATEMENTS READ FROM GOOGLE DRIVE (Ecobank, 26/10/2026)
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

RATE = 10                      # files downloaded per second, at most
TIMEOUT, TRIES, HEDGE = 20, 3, 3   # a request's time limit, requests per file at most, seconds before asking again
CACHE_VERSION = 1


@dataclass
class Picked:
    files: list = field(default_factory=list)        # [(name, content)] of the files that cover the period
    considered: int = 0                              # candidate files on Drive (the bank's own file types/prefix)
    downloaded: int = 0
    unreadable: list = field(default_factory=list)   # names of the files that are not a statement
    errors: list = field(default_factory=list)       # downloads that failed (plain words)


def candidates(files: list, prefix: str = "") -> list:
    """The Drive files to consider: every one, or only those named like `prefix` (any case) when the bank has one -
    Ecobank's folder also holds IREMBO_TACCT_TRANS_..., other exports, that are never statements."""
    if not prefix:
        return list(files)
    return [f for f in files if f.name.strip().upper().startswith(prefix.upper())]


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
    files = candidates(client.list_files(folder_id, bank.FILE_TYPES), engine.option(bank, "DRIVE_STATEMENTS_PREFIX", ""))
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
    def fresh():
        try:
            return client.thread_http(timeout=TIMEOUT)
        except TypeError:                         # a client without the timeout (older running app / test double)
            return client.thread_http()

    def attempt(f):
        limiter.wait()                            # every request counts in the 10 per second
        return client.download(f.id, f.mime, http=fresh() if takes_http else None)

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
