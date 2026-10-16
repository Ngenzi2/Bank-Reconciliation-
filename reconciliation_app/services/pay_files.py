"""
THE IREMBOPAY FILES OF THE CHOSEN BANK (step 2, "IremboPay file" from Google Drive)
==================================================================================
The IremboPay folder on Drive has one folder per bank ("I&M", "Eco Bank", "Bank of Africa"…), each with month
folders ("09/2026"). Step 2 lists only the files of the bank being reconciled; files put directly in the IremboPay
folder (in no bank folder) stay available to every bank. Only the list changes: how a file is read is not touched.

    bank_of_folder()    the bank (code) a Drive folder path belongs to, or None
    files_for_bank()    (the bank's files, the files in no bank folder, whether the bank has a folder)
"""
from __future__ import annotations

import re


def _norm(text: str) -> str:
    """'Eco Bank' -> 'ecobank', 'I&M' -> 'im', 'Bank of Africa' -> 'bankofafrica' (spaces, capitals, & ignored)."""
    return re.sub(r"[^a-z0-9]", "", str(text).lower())


def _aliases(bank) -> set:
    return {a for a in (_norm(bank.CODE), _norm(bank.NAME), _norm(bank.FULL_NAME)) if a}


def top_folder(folder: str) -> str:
    """'Eco Bank / 09/2026' -> 'Eco Bank' (the folder right under the IremboPay folder; '' = the top itself)."""
    return str(folder or "").split(" / ")[0].strip()


def bank_of_folder(folder: str, banks: dict) -> str | None:
    """The code of the bank whose name the top folder carries ('Eco Bank' -> 'ECO'), or None."""
    top = _norm(top_folder(folder))
    if not top:
        return None
    for code, bank in banks.items():
        names = _aliases(bank)
        if top in names or any(top.startswith(n) for n in names if len(n) >= 4):   # 'Bank of Africa Rwanda'
            return code
    return None


def files_for_bank(files: list, bank, banks: dict) -> tuple[list, list, bool]:
    """(files in the bank's folder, files in no bank folder, True when the bank has a folder on Drive)."""
    mine, loose = [], []
    for f in files:
        owner = bank_of_folder(f.folder, banks)
        if owner == bank.CODE:
            mine.append(f)
        elif owner is None and not top_folder(f.folder):          # directly in the IremboPay folder
            loose.append(f)
    has_folder = any(bank_of_folder(f.folder, banks) == bank.CODE for f in files)
    return mine, loose, has_folder
