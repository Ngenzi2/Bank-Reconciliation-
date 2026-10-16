"""Settings changed in the app (targets, SLA days per bank, editors): saved in a file, kept after a restart, every
change recorded with who made it; administrators come from drive_settings.toml [access] (or BANK_RECON_ADMINS)."""
import json

import pytest

import config
from services import app_settings as s

_REAL_LOAD_ADMINS = config.load_admins


def test_defaults_when_there_is_no_file(tmp_path):
    data = s.load(tmp_path / "none.json")
    assert data["targets"] == {"receipt": 0.99, "ontime": 0.90} and data["editors"] == []
    assert s.sla_days(data, "IM") == 0 and s.sla_days(data, "ECO") == 0


def test_targets_are_saved_and_read_back_after_a_restart(tmp_path):
    f = tmp_path / "app_settings.json"
    s.set_targets(0.995, 0.85, "boss@irembo.com", f)
    again = s.load(f)                                            # a new start of the app reads the file
    assert again["targets"] == {"receipt": 0.995, "ontime": 0.85}
    assert [(h["by"], h["what"], h["before"], h["after"]) for h in again["history"]] == [
        ("boss@irembo.com", "Target: notifications received", "99.00%", "99.50%"),
        ("boss@irembo.com", "Target: on time", "90.00%", "85.00%")]
    with pytest.raises(s.SettingsError):
        s.set_targets(1.2, 0.9, "x@irembo.com", f)


def test_sla_days_per_bank_and_a_fixed_rule_bank_has_none(tmp_path):
    f = tmp_path / "app_settings.json"
    s.set_sla_days("BOA", 2, "ed@irembo.com", f)
    data = s.load(f)
    assert s.sla_days(data, "BOA") == 2 and s.sla_days(data, "IM") == 0
    assert json.loads(f.read_text())["banks"] == {"BOA": {"sla_days": 2}}
    with pytest.raises(s.SettingsError):
        s.set_sla_days("ECO", 1, "ed@irembo.com", f)              # Ecobank: same minute, no days
    with pytest.raises(s.SettingsError):
        s.set_sla_days("IM", 99, "ed@irembo.com", f)


def test_roles_from_the_admin_list_and_the_editors(tmp_path, monkeypatch):
    toml = tmp_path / "drive_settings.toml"
    toml.write_text('[access]\nadmins = ["Admin@Irembo.com"]\n')
    monkeypatch.setattr(config, "load_admins", lambda: _REAL_LOAD_ADMINS(toml, env={}))
    f = tmp_path / "app_settings.json"
    s.set_editors(["Ed@irembo.com", ""], "admin@irembo.com", f)
    data = s.load(f)
    assert s.role("ADMIN@irembo.com", data) == "admin" and s.role("ed@irembo.com", data) == "editor"
    assert s.role("someone@irembo.com", data) == "viewer" and s.role("", data) == "viewer"
    assert s.can_edit("editor") and not s.can_edit("viewer")
    with pytest.raises(s.SettingsError):
        s.set_editors(["not-an-email"], "admin@irembo.com", f)


def test_admins_are_read_from_drive_settings_and_the_environment(tmp_path):
    toml = tmp_path / "drive_settings.toml"
    toml.write_text('[google_drive]\nauth = "oauth"\n\n[access]\nadmins = ["A@irembo.com"]\n')
    assert config.load_admins(toml, env={}) == {"a@irembo.com"}
    assert config.load_admins(toml, env={"BANK_RECON_ADMINS": "b@irembo.com, "}) == {"a@irembo.com", "b@irembo.com"}
    assert config.load_admins(tmp_path / "missing.toml", env={}) == set()
