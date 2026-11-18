"""Files read once (services/read_cache.py, 17/11/2026): a file is not read twice while the app runs, and a file known
(from an earlier reading) to be outside the period is not opened again; only its days are kept on this computer."""
import pickle

import pandas as pd
import pytest

import engine as e
from services import read_cache


@pytest.fixture(autouse=True)
def empty_cache():
    read_cache._MEMORY.clear()
    read_cache._FACTS.clear()
    yield
    read_cache._MEMORY.clear()
    read_cache._FACTS.clear()


def reader(calls, first="2026-08-01", last="2026-08-20"):
    def read_file(name, content, bank):
        calls.append(name)
        if content.startswith(b"bad"):
            raise e.ReconError(f"{name}: this is not a statement.")
        return e.UploadedFile(name, "bank", e.file_digest(content), pd.DataFrame({"x": [1, 2]}), "RWF",
                              pd.Timestamp(first), pd.Timestamp(last), 2, bank="IM")
    return read_file


def test_a_file_is_read_once_and_a_refusal_is_kept_too(tmp_path):
    calls, days = [], tmp_path / "days.pkl"
    a = read_cache.read("aug.xlsx", b"august", "bank", "IM", reader(calls), days)
    b = read_cache.read("copy of aug.xlsx", b"august", "bank", "IM", reader(calls), days)
    assert calls == ["aug.xlsx"] and b.name == "copy of aug.xlsx" and a.rows == b.rows == 2
    for _ in range(2):
        with pytest.raises(e.ReconError, match="not a statement"):
            read_cache.read("bad.xlsx", b"bad file", "bank", "IM", reader(calls), days)
    assert calls == ["aug.xlsx", "bad.xlsx"]


def test_a_file_outside_the_period_is_known_without_opening_it(tmp_path):
    days = tmp_path / "days.pkl"
    assert read_cache.outside("aug.xlsx", b"august", "bank", "IM", e.make_period("month", "2026-09-01"), days) is None
    read_cache.read("aug.xlsx", b"august", "bank", "IM", reader([]), days)
    read_cache._FACTS.clear()                                  # another time the app is opened: read from the disk
    why = read_cache.outside("aug.xlsx", b"august", "bank", "IM", e.make_period("month", "2026-09-01"), days)
    assert why.startswith("aug.xlsx: this file covers 01/08/2026 to 20/08/2026, outside September 2026")
    assert read_cache.outside("aug.xlsx", b"august", "bank", "IM", e.make_period("month", "2026-08-01"), days) is None
    saved = pickle.loads(days.read_bytes())["files"]
    assert set(next(iter(saved.values()))) == {"bank", "currency", "first", "last", "rows", "providers"}   # no content
