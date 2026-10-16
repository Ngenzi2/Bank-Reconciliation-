"""
TEMPORARY FILES OF ONE SESSION
==============================
Reports are kept on Google Drive, not on this computer. Some steps still need a file for a moment (the engine
writes the Excel report to a path, a report read back is opened from a path), so each browser session gets its own
private temporary folder, and every file is deleted as soon as it has been read into memory.

    <system temp>/bank-reconciliation/session-<random id>/     (only this computer's user can open it)
    (the environment variable BANK_RECON_TMP can replace <system temp>/bank-reconciliation)

The folder is made only when a file is needed and removed again as soon as it is empty, so nothing stays behind
after a run. A session can only delete inside its own folder: every delete checks the path first. Nothing here ever
touches another session's folder, output/, or any other folder.
"""
from __future__ import annotations

import os
import shutil
import tempfile
import uuid
from pathlib import Path

ROOT = Path(tempfile.gettempdir()) / "bank-reconciliation"


class SessionFolder:
    def __init__(self, root: Path | None = None):
        self.root = Path(root or os.environ.get("BANK_RECON_TMP") or ROOT)
        self.id = uuid.uuid4().hex[:8]
        self._path: Path | None = None             # made on first use

    @property
    def path(self) -> Path:
        """This session's folder (made now if it does not exist; mode 0700)."""
        if self._path is None or not self._path.exists():
            self.root.mkdir(mode=0o700, parents=True, exist_ok=True)
            self._path = Path(tempfile.mkdtemp(prefix=f"session-{self.id}-", dir=self.root))
        return self._path

    def file(self, *parts: str) -> Path:
        """A path inside this session's folder (its parent folders are made)."""
        target = self.path.joinpath(*parts)
        if not self.owns(target):
            raise ValueError("A temporary file must stay inside the session's folder.")
        target.parent.mkdir(parents=True, exist_ok=True)
        return target

    def owns(self, path) -> bool:
        """True when `path` is inside this session's folder (and is not the folder itself)."""
        if path is None or self._path is None:
            return False
        try:
            resolved, base = Path(path).resolve(), self._path.resolve()
        except OSError:
            return False
        return resolved != base and base in resolved.parents

    def remove(self, path) -> bool:
        """Delete one temporary file of this session (then the folder too when it is empty). Anything outside the
        session's folder is never touched."""
        if not self.owns(path):
            return False
        Path(path).unlink(missing_ok=True)
        self._tidy()
        return True

    def clear(self):
        """Delete everything inside this session's folder, and the folder."""
        if self._path is not None and self._path.exists():
            for child in self._path.iterdir():
                if child.is_dir() and not child.is_symlink():
                    shutil.rmtree(child, ignore_errors=True)
                else:
                    child.unlink(missing_ok=True)
        self._tidy()

    def _tidy(self):
        """Remove this session's folder when no file is left in it (it is made again when needed)."""
        if self._path is not None and self._path.exists() and not self.files():
            shutil.rmtree(self._path, ignore_errors=True)
            self._path = None

    def files(self) -> list:
        """Files still in the folder (used to check the clean-up)."""
        if self._path is None or not self._path.exists():
            return []
        return sorted(p for p in self._path.rglob("*") if p.is_file())
