"""The folder picker (Choose: "Browse..."): always a separate process, so a closed window, a missing toolkit or
a slow response never hangs or crashes the app - it just falls back to the text box."""
import subprocess
from types import SimpleNamespace

import pytest

from services import folder_picker as fp


def fake_run(returncode=0, stdout="", stderr=""):
    def run(*a, **k):
        return SimpleNamespace(returncode=returncode, stdout=stdout, stderr=stderr)
    return run


def test_windows_a_chosen_folder_is_returned(monkeypatch):
    monkeypatch.setattr(fp.platform, "system", lambda: "Windows")
    monkeypatch.setattr(fp.subprocess, "run", fake_run(stdout="C:\\data\\BPR"))
    path, error = fp.browse_folder("Choose a folder")
    assert path == "C:\\data\\BPR" and error is None


def test_windows_a_closed_window_is_not_an_error(monkeypatch):
    monkeypatch.setattr(fp.platform, "system", lambda: "Windows")
    monkeypatch.setattr(fp.subprocess, "run", fake_run(stdout=""))     # tkinter: empty path = cancelled
    path, error = fp.browse_folder("Choose a folder")
    assert path is None and error is None


def test_windows_a_failure_to_open_is_an_error_with_a_message(monkeypatch):
    monkeypatch.setattr(fp.platform, "system", lambda: "Windows")
    monkeypatch.setattr(fp.subprocess, "run", fake_run(returncode=1, stderr="TclError: no display name"))
    path, error = fp.browse_folder("Choose a folder")
    assert path is None and error and "no display name" in error


def test_macos_a_chosen_folder_is_returned(monkeypatch):
    monkeypatch.setattr(fp.platform, "system", lambda: "Darwin")
    monkeypatch.setattr(fp.subprocess, "run", fake_run(stdout="/Users/x/data/BPR"))
    path, error = fp.browse_folder("Choose a folder")
    assert path == "/Users/x/data/BPR" and error is None


def test_macos_cancelling_the_window_is_not_an_error(monkeypatch):
    monkeypatch.setattr(fp.platform, "system", lambda: "Darwin")
    monkeypatch.setattr(fp.subprocess, "run", fake_run(returncode=1, stderr="User canceled. (-128)"))
    path, error = fp.browse_folder("Choose a folder")
    assert path is None and error is None


def test_macos_a_real_failure_is_an_error_with_a_message(monkeypatch):
    monkeypatch.setattr(fp.platform, "system", lambda: "Darwin")
    monkeypatch.setattr(fp.subprocess, "run", fake_run(returncode=1, stderr="osascript: command not found"))
    path, error = fp.browse_folder("Choose a folder")
    assert path is None and error and "osascript" in error


def test_a_slow_window_times_out_instead_of_hanging(monkeypatch):
    monkeypatch.setattr(fp.platform, "system", lambda: "Windows")

    def slow(*a, **k):
        raise subprocess.TimeoutExpired(cmd="python", timeout=fp.TIMEOUT)
    monkeypatch.setattr(fp.subprocess, "run", slow)
    path, error = fp.browse_folder("Choose a folder")
    assert path is None and error and "too long" in error


def test_a_typed_path_that_does_not_exist_is_not_passed_as_a_starting_point(monkeypatch, tmp_path):
    monkeypatch.setattr(fp.platform, "system", lambda: "Windows")
    captured = {}

    def run(cmd, **k):
        captured["script"] = cmd[-1]
        return SimpleNamespace(returncode=0, stdout=str(tmp_path), stderr="")
    monkeypatch.setattr(fp.subprocess, "run", run)
    fp.browse_folder("Choose a folder", start=str(tmp_path / "does-not-exist"))
    assert "does-not-exist" not in captured["script"]
    fp.browse_folder("Choose a folder", start=str(tmp_path))     # a real folder: used as the starting point
    assert repr(str(tmp_path)) in captured["script"]
