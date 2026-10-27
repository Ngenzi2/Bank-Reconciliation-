"""Step 6, automatic: the report is saved on Google Drive (replacing a report of the same period) and, once Drive has
confirmed it, the Slack message (with the report's link) is sent - once per report. Drive and Slack are faked: no
test reaches them."""
import json
from pathlib import Path

import pytest

import config
import engine as e
from integrations import google_drive, slack
from integrations.google_drive import DriveFile
from services import report_history
from test_boa import APP_DIR, SEP, run_september


def _same_code_time():
    """The app reloads its modules when the code changed since the last load: mark them as loaded, so the fakes
    below are not replaced by a reload."""
    files = [APP_DIR / "engine.py", APP_DIR / "pdf_report.py", APP_DIR / "integrations" / "irembopay_checkout.py",
             APP_DIR / "integrations" / "slack.py", *APP_DIR.glob("banks/*.py"), *APP_DIR.glob("services/*.py")]
    e._loaded_mtime = max(p.stat().st_mtime for p in files)


@pytest.fixture
def fakes(tmp_path, monkeypatch):
    token = tmp_path / "token.json"
    token.write_text("{}")
    settings = config.DriveSettings(credentials_file=tmp_path / "none.json", reports_root_folder_id="ROOT",
                                    oauth_client_file=token, token_file=token)
    monkeypatch.setattr(config, "load", lambda *a, **k: settings)
    monkeypatch.setattr(google_drive.DriveClient, "connect", classmethod(lambda cls, s: cls(object(), "oauth")))
    calls = {"saves": [], "sent": [], "existing": []}

    def save_report(client, root, bank, period, files, mode="new", existing=None, only=None):
        calls["saves"].append((mode, sorted(files)))
        if calls.get("drive_down"):
            return [report_history.SaveResult(f"x.{ext}", False, error="Drive is down") for ext in files]
        return [report_history.SaveResult(f"{report_history.report_stem(bank, period)}.{ext}", True,
                                          DriveFile(id=f"ID-{ext}", name=f"r.{ext}", mime="", size=1, md5="",
                                                    modified=""), replaced=mode == "replace") for ext in files]

    monkeypatch.setattr(report_history, "existing_reports", lambda *a: calls["existing"])
    monkeypatch.setattr(report_history, "save_report", save_report)
    monkeypatch.setattr(slack, "load_webhook", lambda: "https://hooks.slack.com/services/T/B/fake")
    monkeypatch.setattr(slack, "send_payload", lambda payload, url=None, timeout=15: calls["sent"].append(payload))
    return calls


def step6(tmp_path):
    from streamlit.testing.v1 import AppTest
    res, _ = run_september(tmp_path)
    at = AppTest.from_file(str(APP_DIR / "app.py"), default_timeout=120)
    at.session_state["goto"] = "reconciliation"
    at.session_state["bank_code"] = "BOA"
    at.session_state["report_label"] = "Month"
    at.session_state["picked"] = SEP.start.date()
    at.session_state["step"] = 5        # Report & Export (27/10/2026: 5 steps)
    at.session_state["result"] = res
    at.session_state["report_mem"] = {"name": res.report_file.name, "xlsx": res.report_file.read_bytes(),
                                      "pdf": e.report_pdf(res.report_file)}
    return at


def test_report_is_saved_then_sent_once_with_its_link(tmp_path, fakes):
    at = step6(tmp_path)
    _same_code_time()
    at.run()
    assert not at.exception, [x.value for x in at.exception]
    assert fakes["saves"] == [("new", ["pdf", "xlsx"])]
    assert len(fakes["sent"]) == 1
    sent = json.dumps(fakes["sent"][0], ensure_ascii=False)
    assert "<https://drive.google.com/file/d/ID-xlsx/view|📊 Excel report>" in sent
    assert "<https://drive.google.com/file/d/ID-pdf/view|📄 PDF report>" in sent
    assert "Bank of Africa Rwanda — Monthly Reconciliation | September 2026" in sent
    text = " ".join(str(m.value) for m in at.markdown)
    assert "Saved in" in text and "Sent at" in text
    at.run()                                             # step 6 drawn again: nothing is saved or sent twice
    assert len(fakes["saves"]) == 1 and len(fakes["sent"]) == 1


def test_an_existing_report_is_replaced(tmp_path, fakes):
    fakes["existing"] = [DriveFile(id="OLD", name="old.xlsx", mime="", size=1, md5="", modified="")]
    at = step6(tmp_path)
    _same_code_time()
    at.run()
    assert fakes["saves"][0][0] == "replace" and len(fakes["sent"]) == 1


def test_nothing_is_sent_when_drive_fails(tmp_path, fakes):
    fakes["drive_down"] = True
    at = step6(tmp_path)
    _same_code_time()
    at.run()
    assert not at.exception and fakes["sent"] == []
    assert any("Drive is down" in str(x.value) for x in at.error)
