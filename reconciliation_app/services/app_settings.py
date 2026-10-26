"""
APP SETTINGS: WHAT AUTHORIZED USERS CHANGE IN THE APP, KEPT AFTER A RESTART
==========================================================================
Business settings changed in the app (not in the code), saved in app_settings.json (project folder; the path can be
changed with BANK_RECON_APP_SETTINGS). Every change records who made it and when.

    targets      the Dashboard and Slack targets: notifications received >= receipt, payments on time >= ontime
    banks        per bank: sla_days = the days allowed between the bank date and IremboPay paid_at (banks whose SLA
                 rule uses days; a bank with a fixed rule, e.g. Ecobank's same minute, has none)
    holidays     Rwanda public holidays: a payment made on a weekend or a holiday is posted on the next working day,
                 so a period ending on one needs that day's statement (I&M). Some dates move every year (Easter, Eid)
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
# Rwanda public holidays 2026 (Easter and the two Eid dates move every year; the Eid dates are announced late: check
# them in Settings)
HOLIDAYS_2026 = [
    {"date": "2026-01-01", "name": "New Year's Day"}, {"date": "2026-01-02", "name": "Day after New Year"},
    {"date": "2026-02-01", "name": "National Heroes' Day"}, {"date": "2026-03-20", "name": "Eid al-Fitr (check)"},
    {"date": "2026-04-03", "name": "Good Friday"}, {"date": "2026-04-06", "name": "Easter Monday"},
    {"date": "2026-04-07", "name": "Genocide against the Tutsi Memorial Day"}, {"date": "2026-05-01", "name": "Labour Day"},
    {"date": "2026-05-27", "name": "Eid al-Adha (check)"}, {"date": "2026-07-01", "name": "Independence Day"},
    {"date": "2026-07-04", "name": "Liberation Day"}, {"date": "2026-08-07", "name": "Umuganura"},
    {"date": "2026-08-15", "name": "Assumption Day"}, {"date": "2026-12-25", "name": "Christmas Day"},
    {"date": "2026-12-26", "name": "Boxing Day"},
]
DEFAULTS = {"targets": {"receipt": 0.99, "ontime": 0.90}, "banks": {}, "editors": [], "history": [],
            "holidays": HOLIDAYS_2026,
            # Follow-up page (22/10/2026): the reports followed up (period from this date) and the days before an open
            # item is overdue (red, and in the Slack reminder)
            "followup": {"from": "2026-10-01", "days": 7}}
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
        if isinstance(saved.get("followup"), dict):
            data["followup"].update({k: v for k, v in saved["followup"].items() if k in ("from", "days")})
        if isinstance(saved.get("holidays"), list):
            data["holidays"] = _clean_holidays(saved["holidays"])
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
    """The SLA setting of a bank: days (0 = paid_at on the same day as the bank date), or minutes for a bank whose
    SLA_UNIT is "minutes" (BPR: 1 by default); 0 for a bank with a fixed rule."""
    bank = engine.get_bank(bank)
    if not engine.option(bank, "SLA_DAYS_SETTING", True):
        return 0
    return int((data["banks"].get(bank.CODE) or {}).get("sla_days", engine.option(bank, "SLA_DEFAULT", 0)))


def sla_unit(bank) -> str:
    """'days' or 'minutes': what the SLA setting of a bank counts."""
    return engine.option(engine.get_bank(bank), "SLA_UNIT", "days")


def sla_max(bank) -> int:
    return int(engine.option(engine.get_bank(bank), "SLA_MAX", SLA_DAYS_MAX))


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
    if not 0 <= int(days) <= sla_max(bank):
        raise SettingsError(f"The SLA {sla_unit(bank)} must be between 0 and {sla_max(bank)}.")
    data = load(path)
    old = sla_days(data, bank)
    if old != int(days):
        data["banks"].setdefault(bank.CODE, {})["sla_days"] = int(days)
        data = change(f"{bank.NAME}: SLA {sla_unit(bank)}", old, int(days), who, data, path)
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


def _clean_holidays(rows) -> list:
    """[{"date": "YYYY-MM-DD", "name": ...}] sorted, one per date; rows without a valid date are left out."""
    out = {}
    for r in rows or []:
        date = r.get("date") if isinstance(r, dict) else None
        name = str((r.get("name") if isinstance(r, dict) else "") or "").strip()
        try:
            day = datetime.strptime(str(date)[:10], "%Y-%m-%d").strftime("%Y-%m-%d")
        except (TypeError, ValueError):
            continue
        out[day] = name or "Holiday"
    return [{"date": d, "name": n} for d, n in sorted(out.items())]


def set_holidays(rows: list, who: str, path: Path | None = None) -> dict:
    """Save the public holidays as `who` (an editor or administrator); rows: [{"date", "name"}]."""
    for r in rows:
        if str(r.get("date") or "").strip() and not _clean_holidays([r]):
            raise SettingsError(f"Not a date: {r.get('date')} (use YYYY-MM-DD).")
    data = load(path)
    new = _clean_holidays(rows)
    if new != data["holidays"]:
        before = {h["date"]: h["name"] for h in data["holidays"]}
        after = {h["date"]: h["name"] for h in new}
        added = [f"{d} {after[d]}" for d in after if d not in before or before[d] != after[d]]
        removed = [f"{d} {before[d]}" for d in before if d not in after]
        data["holidays"] = new
        data = change("Public holidays", "removed: " + (", ".join(removed) or "none"),
                      "added/changed: " + (", ".join(added) or "none"), who, data, path)
    return data


def set_followup(start, days: int, who: str, path: Path | None = None) -> dict:
    """Save the Follow-up settings as `who`: the first period followed up (a date) and the days before overdue."""
    try:
        start = datetime.strptime(str(start)[:10], "%Y-%m-%d").strftime("%Y-%m-%d")
    except ValueError:
        raise SettingsError("The follow-up start must be a date.") from None
    if not 1 <= int(days) <= 90:
        raise SettingsError("The days before an item is overdue must be between 1 and 90.")
    data = load(path)
    for key, label, v in (("from", "Follow-up: reports from", start), ("days", "Follow-up: overdue after (days)",
                                                                      int(days))):
        old = data["followup"][key]
        if old != v:
            data["followup"][key] = v
            data = change(label, old, v, who, data, path)
    return data
