"""Shared test setup: the app's folder on the import path; no test may read or change the real Drive settings,
OAuth client, token, service-account key or Slack webhook of this computer."""
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
APP_DIR = ROOT / "reconciliation_app"
sys.path.insert(0, str(APP_DIR))
sys.path.insert(0, str(Path(__file__).resolve().parent))


@pytest.fixture(autouse=True)
def isolated_secrets(tmp_path, monkeypatch):
    secrets = tmp_path / "secrets"
    monkeypatch.setenv("BANK_RECON_OAUTH_CLIENT", str(secrets / "oauth-client.json"))
    monkeypatch.setenv("BANK_RECON_TOKEN", str(secrets / "token.json"))
    monkeypatch.setenv("BANK_RECON_CREDENTIALS", str(secrets / "service-account.json"))
    monkeypatch.setenv("BANK_RECON_TMP", str(tmp_path / "session-tmp"))
    monkeypatch.setenv("BANK_RECON_SLACK_WEBHOOK_FILE", str(secrets / "slack-webhook.txt"))   # never the real channel
    monkeypatch.setenv("IM_RECON_BASE", str(tmp_path / "base"))
    import config
    monkeypatch.setattr(config, "SETTINGS_FILE", tmp_path / "no-drive-settings.toml")
    monkeypatch.setattr(config.load, "__defaults__", (tmp_path / "no-drive-settings.toml", None))
    return secrets
