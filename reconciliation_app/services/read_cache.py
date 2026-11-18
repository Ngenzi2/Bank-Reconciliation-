"""
FILES READ ONCE (17/11/2026)
============================
Opening a statement or an IremboPay file and checking it takes seconds, and a folder of a few months holds dozens of
files: each new period read every one of them again (a BPR folder: 87 files, about a minute). Now:
    - the days a file covers (and its bank, currency, providers) are remembered on this computer, by the file's
      fingerprint (output/.file_days.pkl): a file outside the period is not opened again; it is listed as before,
      with the same reason, and opened only if a later period needs it;
    - a file opened once is kept in memory while the app runs (the same file is not read twice, in any session).
Only these facts are written on this computer, never the content of a file (uploaded files are never kept).

    read(name, content, kind, bank, read_file)   the UploadedFile of a file (read_file: engine.read_bank_file or
                                                 engine.read_payment_file); its ReconError as it was
    outside(name, content, kind, bank, period, facts_path, digest)   the reason the file is not used for `period`,
                                                 when it is already known without opening it (else None)
    forget(content, kind, bank)                  drop a file from memory (not needed for the period)
"""
from __future__ import annotations

import pickle
import threading
import uuid
from collections import OrderedDict
from pathlib import Path

import engine

VERSION = 1
KEEP = 24                                   # files kept in memory (a BK month: 5 files of up to 100,000 rows)
_MEMORY: "OrderedDict[tuple, object]" = OrderedDict()
_LOCK = threading.Lock()
_FACTS: dict = {}                           # path -> {key: facts}, loaded once


def _key(digest: str, kind: str, bank) -> tuple:
    return digest, kind, engine.get_bank(bank).CODE, getattr(engine, "_loaded_mtime", 0)


def _load(path: Path) -> dict:
    if str(path) not in _FACTS:
        try:
            with open(path, "rb") as fh:
                saved = pickle.load(fh)
            _FACTS[str(path)] = saved["files"] if saved.get("version") == VERSION else {}
        except Exception:
            _FACTS[str(path)] = {}
    return _FACTS[str(path)]


def _save(path: Path) -> None:
    tmp = path.with_name(f".{path.name}.{uuid.uuid4().hex[:8]}.tmp")
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        with open(tmp, "wb") as fh:
            pickle.dump({"version": VERSION, "files": _FACTS[str(path)]}, fh)
        tmp.replace(path)
    except Exception:                           # not saved: the file is opened again next time
        tmp.unlink(missing_ok=True)


def read(name: str, content: bytes, kind: str, bank, read_file, facts_path: Path | None = None):
    """The file read (engine.UploadedFile), from memory when it was read before; raises its ReconError."""
    digest = engine.file_digest(content)
    key = _key(digest, kind, bank)
    with _LOCK:
        hit = _MEMORY.get(key)
        if hit is not None:
            _MEMORY.move_to_end(key)
    if hit is None:
        try:
            hit = read_file(name, content, bank)
        except engine.ReconError as err:
            hit = err
        with _LOCK:
            _MEMORY[key] = hit
            while len(_MEMORY) > KEEP:
                _MEMORY.popitem(last=False)
        if facts_path is not None:
            facts = _load(facts_path)
            facts[key[:3]] = (None if isinstance(hit, engine.ReconError) else
                              {"bank": hit.bank, "currency": hit.currency, "first": hit.first, "last": hit.last,
                               "rows": hit.rows, "providers": hit.providers})
            _save(facts_path)
    if isinstance(hit, engine.ReconError):
        raise hit
    return engine.UploadedFile(**{**hit.__dict__, "name": name})


def forget(content: bytes, kind: str, bank) -> None:
    """Drop a file read from memory (it is not needed for the period; its days stay known)."""
    with _LOCK:
        _MEMORY.pop(_key(engine.file_digest(content), kind, bank), None)


def outside(name: str, content: bytes, kind: str, bank, period, facts_path: Path, digest: str = "") -> str | None:
    """The reason a file is not used for `period` (outside it, another bank, no payments of this bank), when it is
    known from an earlier reading without opening the file; None when it may be used or is not known yet.
    digest: the file's fingerprint when already known (a BK statement is 20 MB: not computed again)."""
    digest = digest or engine.file_digest(content)
    facts = _load(facts_path).get((digest, kind, engine.get_bank(bank).CODE))
    if not facts:
        return None
    stub = engine.UploadedFile(name, kind, digest, None, facts["currency"], facts["first"],
                               facts["last"], facts["rows"], bank=facts["bank"], providers=facts["providers"])
    return engine.accept_file(period, stub, [], bank)
