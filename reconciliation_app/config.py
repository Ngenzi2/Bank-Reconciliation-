"""
GOOGLE DRIVE SETTINGS
=====================
Where the app finds its Google Drive settings. Nothing secret is written in the code or read into it here: the
OAuth client, the sign-in token and the service-account key stay in their own files, outside the project folder;
this file only knows their paths.

How the app signs in to Google Drive (auth):
    oauth            (default) as you, with your Google account: reports are created in your My Drive
    service_account  as a service account (fallback): it can read shared folders, but it cannot create files in
                     a My Drive folder (it has no storage of its own)

Each setting is taken from (first found):
    1. an environment variable
    2. drive_settings.toml in the project folder (ignored by Git; drive_settings.example.toml shows the format)
    3. the default

    setting                   environment variable             default
    auth                      BANK_RECON_AUTH                  oauth
    oauth_client_file         BANK_RECON_OAUTH_CLIENT          ~/.config/bank-reconciliation/oauth-client.json
    token_file                BANK_RECON_TOKEN                 ~/.config/bank-reconciliation/token.json
    credentials_file          BANK_RECON_CREDENTIALS           ~/.config/bank-reconciliation/service-account.json
    reports_root_folder_id    BANK_RECON_REPORTS_FOLDER_ID     (none) the "Reconciliation Reports" folder
    irembopay_folder_id       BANK_RECON_IREMBOPAY_FOLDER_ID   (none) the folder of the IremboPay payment files
    ecobank_statements_folder_id  BANK_RECON_ECOBANK_FOLDER_ID (none) the Ecobank "Statements" folder (IREMBO PAY files)
    bpr_statements_folder_id  BANK_RECON_BPR_FOLDER_ID         (none) the BPR "Statements" folder
"""
from __future__ import annotations

import os
import tomllib
from dataclasses import dataclass
from pathlib import Path

PROJECT_DIR = Path(__file__).resolve().parent.parent
SETTINGS_FILE = PROJECT_DIR / "drive_settings.toml"
SECRETS_DIR = Path.home() / ".config" / "bank-reconciliation"
DEFAULT_CREDENTIALS = SECRETS_DIR / "service-account.json"
DEFAULT_OAUTH_CLIENT = SECRETS_DIR / "oauth-client.json"
DEFAULT_TOKEN = SECRETS_DIR / "token.json"
AUTH_MODES = ("oauth", "service_account")
ENV = {"auth": "BANK_RECON_AUTH",
       "oauth_client_file": "BANK_RECON_OAUTH_CLIENT",
       "token_file": "BANK_RECON_TOKEN",
       "credentials_file": "BANK_RECON_CREDENTIALS",
       "reports_root_folder_id": "BANK_RECON_REPORTS_FOLDER_ID",
       "irembopay_folder_id": "BANK_RECON_IREMBOPAY_FOLDER_ID",
       "ecobank_statements_folder_id": "BANK_RECON_ECOBANK_FOLDER_ID",
       "bpr_statements_folder_id": "BANK_RECON_BPR_FOLDER_ID"}


@dataclass(frozen=True)
class DriveSettings:
    credentials_file: Path                           # the service-account key (auth = service_account)
    reports_root_folder_id: str = ""
    irembopay_folder_id: str = ""
    ecobank_statements_folder_id: str = ""           # step 2: the Ecobank statements are read from there
    bpr_statements_folder_id: str = ""               # step 2: the BPR statements are read from there
    error: str = ""                                  # drive_settings.toml could not be read
    auth: str = "oauth"
    oauth_client_file: Path = DEFAULT_OAUTH_CLIENT   # the OAuth "Desktop app" client (auth = oauth)
    token_file: Path = DEFAULT_TOKEN                 # your sign-in, written by "Sign in with Google"

    @property
    def is_oauth(self) -> bool:
        return self.auth == "oauth"

    def problems(self) -> list:
        """What is missing in the settings, in plain words (empty = ready). Not being signed in yet is not a
        settings problem (see integrations/google_drive.py). The secret files are never opened here."""
        out = [self.error] if self.error else []
        if self.auth not in AUTH_MODES:
            out.append(f"Unknown auth '{self.auth}' (use 'oauth' or 'service_account').")
        elif self.is_oauth and not self.oauth_client_file.is_file():
            out.append(f"No OAuth client file at {self.oauth_client_file}.")
        elif not self.is_oauth and not self.credentials_file.is_file():
            out.append(f"No credentials file at {self.credentials_file}.")
        if not self.reports_root_folder_id and not self.irembopay_folder_id:
            out.append("No Drive folder set (reports_root_folder_id / irembopay_folder_id).")
        return out


def load(settings_file: Path = SETTINGS_FILE, env=None) -> DriveSettings:
    """The Drive settings (environment variables first, then drive_settings.toml, then the defaults)."""
    env = os.environ if env is None else env
    values, error = {}, ""
    if Path(settings_file).is_file():
        try:
            values = tomllib.loads(Path(settings_file).read_text()).get("google_drive", {})
        except (tomllib.TOMLDecodeError, OSError) as err:
            error = f"{Path(settings_file).name} could not be read ({err})."

    def get(key, default=""):
        return str(env.get(ENV[key]) or values.get(key) or default).strip()

    def path(key, default):
        return Path(get(key, default)).expanduser()

    return DriveSettings(credentials_file=path("credentials_file", DEFAULT_CREDENTIALS),
                         reports_root_folder_id=get("reports_root_folder_id"),
                         irembopay_folder_id=get("irembopay_folder_id"),
                         ecobank_statements_folder_id=get("ecobank_statements_folder_id"),
                         bpr_statements_folder_id=get("bpr_statements_folder_id"), error=error,
                         auth=get("auth", "oauth").lower(),
                         oauth_client_file=path("oauth_client_file", DEFAULT_OAUTH_CLIENT),
                         token_file=path("token_file", DEFAULT_TOKEN))


def load_admins(settings_file: Path = SETTINGS_FILE, env=None) -> set:
    """The administrators of the app's settings (lower-case e-mail addresses): drive_settings.toml, section [access],
    admins = ["name@irembo.com", ...], and BANK_RECON_ADMINS (comma-separated). Set up by a developer on this
    computer; nothing else can make someone an administrator."""
    env = os.environ if env is None else env
    found = set()
    if Path(settings_file).is_file():
        try:
            found |= {str(a) for a in tomllib.loads(Path(settings_file).read_text()).get("access", {}).get("admins", [])}
        except (tomllib.TOMLDecodeError, OSError):
            pass
    found |= {a for a in str(env.get("BANK_RECON_ADMINS", "")).split(",")}
    return {a.strip().lower() for a in found if a.strip()}
