"""
APP SETTINGS: WHAT AUTHORIZED USERS CHANGE IN THE APP, KEPT AFTER A RESTART
==========================================================================
Business settings changed in the app (not in the code), saved in app_settings.json (project folder; the path can be
changed with BANK_RECON_APP_SETTINGS). Every change records who made it and when.

    targets      the Dashboard and Slack targets: notifications received >= receipt, payments on time >= ontime
    banks        per bank: sla_days = the days allowed between the bank date and IremboPay paid_at (banks whose SLA
                 rule uses days; a bank with a fixed rule, e.g. Ecobank's same minute, has none)
    editors      the Google accounts allowed to change these settings (managed by an administrator)
    history      the changes: when, who, what, before -> after (the last 500)

Who may change what (the account signed in to Google Drive in the app):
    administrator   listed in drive_settings.toml, [access] admins (a file on this computer, set up by a developer),
                    or in BANK_RECON_ADMINS: every setting + the editors; sees the technical settings of the banks
    editor          listed by an administrator in the app: targets and SLA days
    viewer          anyone else (or not signed in): sees the settings, changes nothing
Technical settings (statement formats, columns, matching rules) stay in the code (banks/<bank>.py): developers only.
"""
from __future__ import annotations

import json
import os
from copy import deepcopy
from datetime import datetime
from pathlib import Path

import config
import engine

FILE = Path(os.environ.get("BANK_RECON_APP_SETTINGS") or config.PROJECT_DIR / "app_settings.json")
DEFAULTS = {"targets": {"receipt": 0.99, "ontime": 0.90}, "banks": {}, "editors": [], "history": []}
SLA_DAYS_MAX = 10
ROLES = {"admin": "Administrator", "editor": "Editor", "viewer": "Viewer"}


class SettingsError(Exception):
    """A setting could not be saved, in plain words."""


def load(path: Path | None = None) -> dict:
    """The settings, every missing value taken from DEFAULTS (a missing or unreadable file = the defaults)."""
    path = Path(path or FILE)
    data = deepcopy(DEFAULTS)
    if path.is_file():
        try:
            saved = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            return data
        data["targets"].update({k: float(v) for k, v in (saved.get("targets") or {}).items() if k in data["targets"]})
        data["banks"] = {str(k): dict(v) for k, v in (saved.get("banks") or {}).items()}
        data["editors"] = sorted({str(e).strip().lower() for e in saved.get("editors") or [] if str(e).strip()})
        data["history"] = list(saved.get("history") or [])[-500:]
    return data


def _save(data: dict, path: Path | None = None):
    path = Path(path or FILE)
    tmp = path.with_suffix(".tmp")
    try:
        tmp.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
        os.replace(tmp, path)
    except OSError as err:
        raise SettingsError(f"The settings could not be saved ({err.strerror or err}).") from None


def admins() -> set:
    """The administrators (lower-case e-mail addresses): drive_settings.toml [access] admins + BANK_RECON_ADMINS."""
    return config.load_admins()


def role(email: str | None, data: dict | None = None) -> str:
    """'admin', 'editor' or 'viewer' for the account signed in to Google Drive (None = not signed in)."""
    who = (email or "").strip().lower()
    if not who:
        return "viewer"
    if who in admins():
        return "admin"
    if who in (data or load())["editors"]:
        return "editor"
    return "viewer"


def can_edit(r: str) -> bool:
    return r in ("admin", "editor")


def sla_days(data: dict, bank) -> int:
    """The SLA days of a bank (0 = paid_at on the same day as the bank date); 0 for a bank with a fixed rule."""
    bank = engine.get_bank(bank)
    if not engine.option(bank, "SLA_DAYS_SETTING", True):
        return 0
    return int((data["banks"].get(bank.CODE) or {}).get("sla_days", 0))


def change(what: str, before, after, who: str, data: dict | None = None, path: Path | None = None) -> dict:
    """Record one change made by `who` (it must already be applied to `data`), save, and return the settings."""
    data = data if data is not None else load(path)
    data["history"] = (data["history"] + [{"at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "by": who,
                                           "what": what, "before": before, "after": after}])[-500:]
    _save(data, path)
    return data


def set_targets(receipt: float, ontime: float, who: str, path: Path | None = None) -> dict:
    """Save the targets (0..1) as `who`; each one changed is recorded."""
    for name, v in (("received", receipt), ("on time", ontime)):
        if not 0 < v <= 1:
            raise SettingsError(f"The {name} target must be between 0% and 100%.")
    data = load(path)
    for key, label, v in (("receipt", "Target: notifications received", receipt), ("ontime", "Target: on time", ontime)):
        old = data["targets"][key]
        if round(old, 6) != round(v, 6):
            data["targets"][key] = round(float(v), 6)
            data = change(label, f"{old:.2%}", f"{v:.2%}", who, data, path)
    return data


def set_sla_days(bank, days: int, who: str, path: Path | None = None) -> dict:
    """Save the SLA days of a bank as `who` (0..SLA_DAYS_MAX)."""
    bank = engine.get_bank(bank)
    if not engine.option(bank, "SLA_DAYS_SETTING", True):
        raise SettingsError(f"{bank.NAME} has a fixed SLA rule: there are no days to set.")
    if not 0 <= int(days) <= SLA_DAYS_MAX:
        raise SettingsError(f"The SLA days must be between 0 and {SLA_DAYS_MAX}.")
    data = load(path)
    old = sla_days(data, bank)
    if old != int(days):
        data["banks"].setdefault(bank.CODE, {})["sla_days"] = int(days)
        data = change(f"{bank.NAME}: SLA days", old, int(days), who, data, path)
    return data


def set_editors(editors: list, who: str, path: Path | None = None) -> dict:
    """Save the list of editors as `who` (an administrator)."""
    data = load(path)
    new = sorted({str(e).strip().lower() for e in editors if str(e).strip()})
    bad = [e for e in new if "@" not in e]
    if bad:
        raise SettingsError(f"Not an e-mail address: {', '.join(bad)}")
    if new != data["editors"]:
        old = ", ".join(data["editors"]) or "none"
        data["editors"] = new
        data = change("Editors", old, ", ".join(new) or "none", who, data, path)
    return data
