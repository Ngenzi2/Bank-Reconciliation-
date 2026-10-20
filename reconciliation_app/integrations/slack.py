"""
SLACK (Results, "Send to Slack")
===============================
Posts the report message to a Slack channel through an Incoming Webhook. The webhook URL is a secret (whoever has
it can post to the channel): it is kept in its own file outside the project folder and Git, readable only by this
computer's user, and it is never shown in full (masked()).

    webhook file   ~/.config/bank-reconciliation/slack-webhook.txt   (or the path in BANK_RECON_SLACK_WEBHOOK_FILE)

    save_webhook()  keep a new webhook URL          load_webhook()  the URL, or None
    remove_webhook() forget it (Settings → Slack in the side panel)
    masked()        'hooks.slack.com/services/T9S…/…cqT'            send()  post a message (raises SlackError)
    send_payload()  post a message made of Slack blocks (the Results report message)
"""
from __future__ import annotations

import json
import os
import urllib.error
import urllib.request
from pathlib import Path

import config

PREFIX = "https://hooks.slack.com/services/"


class SlackError(Exception):
    pass


def webhook_file() -> Path:
    return Path(os.environ.get("BANK_RECON_SLACK_WEBHOOK_FILE") or config.SECRETS_DIR / "slack-webhook.txt")


def load_webhook() -> str | None:
    path = webhook_file()
    try:
        url = path.read_text(encoding="utf-8").strip()
    except OSError:
        return None
    return url if url.startswith(PREFIX) else None


def save_webhook(url: str) -> None:
    """Keep the webhook URL in its file (created readable by this user only)."""
    url = (url or "").strip()
    if not url.startswith(PREFIX) or len(url) <= len(PREFIX) + 10:
        raise SlackError("This is not a Slack Incoming Webhook URL (it starts with https://hooks.slack.com/services/).")
    path = webhook_file()
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(".tmp")
    if tmp.exists():
        tmp.unlink()
    fd = os.open(tmp, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600)
    with os.fdopen(fd, "w", encoding="utf-8") as fh:
        fh.write(url + "\n")
    os.replace(tmp, path)
    os.chmod(path, 0o600)


def remove_webhook() -> None:
    """Forget the webhook URL (its file is deleted; nothing else is touched)."""
    webhook_file().unlink(missing_ok=True)


def masked(url: str | None) -> str:
    """'hooks.slack.com/services/T9S…/…cqT': enough to recognise it, never the whole secret."""
    if not url:
        return ""
    parts = url[len(PREFIX):].split("/")
    return f"hooks.slack.com/services/{parts[0][:3]}…/…{parts[-1][-3:]}" if parts else "hooks.slack.com/…"


def send(text: str, url: str | None = None, timeout: int = 15) -> None:
    """Post a plain-text message; raises SlackError with a plain message when Slack does not accept it."""
    send_payload({"text": text}, url, timeout)


def send_payload(payload: dict, url: str | None = None, timeout: int = 15) -> None:
    """Post a message made of Slack blocks ({"text": the notification line, "blocks": [...]})."""
    url = url or load_webhook()
    if not url:
        raise SlackError("No Slack webhook is set.")
    req = urllib.request.Request(url, data=json.dumps(payload).encode("utf-8"),
                                 headers={"Content-Type": "application/json"}, method="POST")
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            body = r.read().decode("utf-8", "replace").strip()
            if r.status != 200 or body != "ok":
                raise SlackError(f"Slack answered: {body or r.status}")
    except urllib.error.HTTPError as err:
        detail = err.read().decode("utf-8", "replace").strip()
        hint = " The webhook may have been removed or regenerated: paste the new one." if err.code in (403, 404) else ""
        raise SlackError(f"Slack refused the message ({err.code}: {detail or err.reason}).{hint}") from None
    except (urllib.error.URLError, TimeoutError, OSError) as err:
        raise SlackError(f"Slack could not be reached ({getattr(err, 'reason', err)}). Check the internet "
                         "connection and try again.") from None
