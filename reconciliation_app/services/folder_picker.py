"""
FOLDER PICKER (Choose: "Browse..." next to a bank's statements/exports folder)
===============================================================================
Opens the computer's own "choose a folder" window, so a folder can be picked instead of typed. A GUI toolkit and
Streamlit's own event loop do not mix safely in one process, so this always runs as a separate small process (a
one-off Python/tkinter process on Windows, osascript on macOS) with a timeout - it can never hang or crash the
app. On any problem (no display, the toolkit is missing, it took too long) it returns an error message instead of
raising, so the caller can fall back to the text box that is there anyway. A closed/cancelled window is not an
error: the caller just keeps the text box as it was.

    browse_folder(title, start="") -> (path or None, error or None)
"""
from __future__ import annotations

import platform
import subprocess
import sys
from pathlib import Path

TIMEOUT = 120   # seconds: generous for someone actually browsing, but always bounded


def browse_folder(title: str, start: str = "") -> tuple[str | None, str | None]:
    """Opens the OS's own folder window. (path, None) when one was chosen; (None, None) when the window was
    closed or cancelled; (None, error) when the window could not be opened at all."""
    if start and not Path(start).expanduser().is_dir():
        start = ""                                # a typed path that does not (yet) exist: open with no default
    try:
        if platform.system() == "Darwin":
            return _browse_macos(title, start)
        return _browse_windows(title, start)
    except subprocess.TimeoutExpired:
        return None, "The folder window took too long to respond."
    except FileNotFoundError:
        return None, "The folder window could not be opened on this computer."
    except Exception as err:
        return None, _short(str(err))


def _short(text: str) -> str:
    text = (text or "").strip().splitlines()
    return (text[-1][:200] if text else "The folder window could not be opened.")


def _browse_windows(title: str, start: str) -> tuple[str | None, str | None]:
    script = (
        "import sys, tkinter as tk\n"
        "from tkinter import filedialog\n"
        "root = tk.Tk()\n"
        "root.withdraw()\n"
        "root.attributes('-topmost', True)\n"
        f"path = filedialog.askdirectory(title={title!r}, initialdir={start!r} or None)\n"
        "root.destroy()\n"
        "sys.stdout.write(path)\n"
    )
    result = subprocess.run([sys.executable, "-c", script], capture_output=True, text=True, timeout=TIMEOUT)
    if result.returncode != 0:
        return None, _short(result.stderr)
    return (result.stdout.strip() or None), None


def _applescript_text(value: str) -> str:
    """A Python string as an AppleScript double-quoted literal."""
    return '"' + value.replace("\\", "\\\\").replace('"', '\\"') + '"'


def _browse_macos(title: str, start: str) -> tuple[str | None, str | None]:
    where = f" default location (POSIX file {_applescript_text(start)})" if start else ""
    script = f"POSIX path of (choose folder with prompt {_applescript_text(title)}{where})"
    result = subprocess.run(["osascript", "-e", script], capture_output=True, text=True, timeout=TIMEOUT)
    if result.returncode != 0:
        return (None, None) if "User canceled" in result.stderr else (None, _short(result.stderr))
    return (result.stdout.strip() or None), None
