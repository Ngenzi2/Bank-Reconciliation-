"""02/11/2026: errors said in plain words (never a traceback), and files saved by two steps at the same time never
mixed up (each save writes under a name of its own)."""
import threading
import zipfile

import engine
from integrations.google_drive import DriveError
from services import app_settings, facts_cache, followup, friendly


def test_the_apps_own_messages_are_kept_as_they_are():
    assert friendly.message(engine.ReconError("No I&M report found for September 2026.")) == \
        "No I&M report found for September 2026."
    assert friendly.message(DriveError("Could not download the file.")) == "Could not download the file."


def test_the_usual_problems_are_said_in_plain_words():
    missing = FileNotFoundError(2, "No such file or directory", "/var/folders/x/IM_recon.xlsx")
    assert friendly.message(missing, "The reconciliation") == \
        "The reconciliation did not work: a temporary file was missing. Please try again."
    assert "open in another program" in friendly.message(PermissionError(13, "denied"))
    assert "not a valid Excel file" in friendly.message(zipfile.BadZipFile("File is not a zip file"))
    assert "internet could not be reached" in friendly.message(ConnectionError("reset"))


def test_anything_else_is_short_and_says_what_to_do():
    msg = friendly.message(KeyError("Total Amount"), "Reading this report")
    assert msg == ("Reading this report did not work. Please try again; if it happens again, tell the developer "
                   "(KeyError: 'Total Amount').")
    long = friendly.message(ValueError("x" * 500 + "\nTraceback line"))
    assert len(long) < 260 and "Traceback" not in long


def _at_once(save, n=8):
    threads = [threading.Thread(target=save, args=(i,)) for i in range(n)]
    for t in threads:
        t.start()
    for t in threads:
        t.join()


def test_caches_saved_at_the_same_time_stay_readable_and_leave_no_temporary_file(tmp_path):
    fu, facts = tmp_path / "followup.pkl", tmp_path / "facts.pkl"
    _at_once(lambda i: followup.cache_save(fu, {f"r{i}": [{"key": i}]}))
    _at_once(lambda i: facts_cache.save(facts, {}))
    assert followup.cache_load(fu)                       # a whole file, never a half-written one
    assert facts.exists()
    assert not list(tmp_path.glob("*.tmp")) and not list(tmp_path.glob(".*.tmp"))


def test_settings_saved_at_the_same_time_stay_readable(tmp_path):
    path = tmp_path / "app_settings.json"
    _at_once(lambda i: app_settings._save({"editors": [f"user{i}@irembo.com"]}, path))
    assert app_settings.load(path)["editors"]
    assert not list(tmp_path.glob(".*.tmp"))


def test_review_decisions_saved_at_the_same_time_stay_readable(tmp_path):
    import pandas as pd
    store = pd.DataFrame({"bank": ["IM"], "period": ["2026-09"]})
    _at_once(lambda i: engine._write_store(tmp_path, store))
    assert pd.read_csv(engine.store_path(tmp_path)).shape[0] == 1
    assert not list(engine.store_path(tmp_path).parent.glob(".*.tmp"))
