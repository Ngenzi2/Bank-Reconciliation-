"""BPR: reading the statement (currency from the file name, invoice from the DOC column or the Description,
reversals excluded), the SLA (at most 1 minute after VALUE_DATE, or same day with no VALUE_DATE), matching, and
the automatic classification invariant. Made-up files with the real structure (bpr_fixtures.py), plus the real
August 2026 samples when they are on this computer (skipped otherwise)."""
import os
from pathlib import Path

import pandas as pd
import pytest

import config
import engine as e
from bpr_fixtures import payments, statement
from services import auto_review, folder_files

SEP = e.make_period("month", "2026-09-01")
SAMPLES = os.environ.get("BANK_RECON_SAMPLES")


def read(content, name="statement.xlsx"):
    return e.read_bank_file(name, content, "BPR")


def test_currency_comes_from_the_file_name():
    rwf = read(statement([("05/09/2026", "R1", "x", None, "2026-09-05 10:00:00", None, 1000)]), "sept 10 2026.xlsx")
    usd = read(statement([("05/09/2026", "R1", "x", None, "2026-09-05 10:00:00", None, 1000)]), "aug 20 usd.xlsx")
    assert rwf.currency == "RWF" and usd.currency == "USD"


def test_invoice_from_doc_column_float_like_or_from_description_or_missing():
    rows = [
        ("05/09/2026", "R1", "FTCM something", None, "2026-09-05 10:00:00", None, 5000),           # no invoice
        ("05/09/2026", "R2", "20144909967", "886218433552.0", "2026-09-05 10:01:00", None, 6000),  # DOC column
        ("06/09/2026", "R3", "886218433560 direct", None, "2026-09-06 09:00:00", None, 7000),      # in Description
    ]
    f = read(statement(rows))
    data = f.data.set_index("Bank Reference")
    assert pd.isna(data.loc["R1", "invoice_number"])
    assert data.loc["R2", "invoice_number"] == "886218433552"
    assert data.loc["R3", "invoice_number"] == "886218433560"
    assert len(f.data) == 3                                           # the no-invoice credit is kept too


def test_a_credit_cancelled_by_a_debit_with_the_same_reference_is_not_a_payment():
    rows = [
        ("08/09/2026", "R5", "886218433580", None, "2026-09-08 09:00:00", None, 9000),
        ("08/09/2026", "R5", "Cancel 886218433580", None, "2026-09-08 09:05:00", -9000, None),
        ("08/09/2026", "R6", "886218433581", None, "2026-09-08 09:10:00", None, 1000),   # a normal credit, kept
    ]
    f = read(statement(rows))
    assert list(f.data["Bank Reference"]) == ["R6"]


def test_no_value_date_column_falls_back_to_booking_date():
    rows = [("09/09/2026", "R6", "886218433590", None, None, None, 1000)]
    f = read(statement(rows, with_value_date=False), "no_value_date.xlsx")
    assert f.data["Has time"].eq(False).all()
    assert f.data["Value Date"].iloc[0] == pd.Timestamp("2026-09-09")


def run_september(rows, pays, tmp_path):
    files = [read(statement(rows)), e.read_payment_file("p.csv", payments(pays), "BPR")]
    return e.run(tmp_path / "base", SEP, files, require_review=False, bank="BPR", report_dir=tmp_path / "r")


def test_sla_on_time_violated_and_by_day_without_a_value_date(tmp_path):
    rows = [
        ("05/09/2026", "R2", "20144909967", "886218433552.0", "2026-09-05 10:01:00", None, 6000),   # on time (<=60s)
        ("06/09/2026", "R3", "886218433560 direct", None, "2026-09-06 09:00:00", None, 7000),       # violated (5 min)
    ]
    pays = [
        dict(invoice="886218433552", amount=6000, paid_at="2026-09-05 10:01:30", ref="R2"),
        dict(invoice="886218433560", amount=7000, paid_at="2026-09-06 09:05:00", ref="R3"),
    ]
    res = run_september(rows, pays, tmp_path)
    details = res.sheets["Payment details"].set_index("invoice_number")
    assert details.loc["886218433552", "sla_status"] == "On time"
    assert details.loc["886218433560", "sla_status"] == "SLA violated"


def test_sla_checked_by_day_when_the_file_has_no_value_date(tmp_path):
    rows = [("09/09/2026", "R6", "886218433590", None, None, None, 1000)]
    pays = [dict(invoice="886218433590", amount=1000, paid_at="2026-09-09 15:00:00", ref="R6")]
    files = [read(statement(rows, with_value_date=False), "no_vd.xlsx"),
            e.read_payment_file("p.csv", payments(pays), "BPR")]
    res = e.run(tmp_path / "base", SEP, files, require_review=False, bank="BPR", report_dir=tmp_path / "r")
    details = res.sheets["Payment details"].set_index("invoice_number")
    assert details.loc["886218433590", "sla_status"] == "On time"


def test_bpr_rw_counts_as_bpr(tmp_path):
    rows = [("05/09/2026", "R2", "886218433552", None, "2026-09-05 10:00:00", None, 6000)]
    pays = [dict(invoice="886218433552", amount=6000, paid_at="2026-09-05 10:00:30", ref="R2", provider="BPR_RW")]
    res = run_september(rows, pays, tmp_path)
    details = res.sheets["Payment details"].set_index("invoice_number")
    assert bool(details.loc["886218433552", "notification_received"])


def test_same_day_saved_in_two_files_is_deduplicated(tmp_path):
    rows = [("05/09/2026", "R2", "886218433552", None, "2026-09-05 10:00:00", None, 6000)]
    pays = [dict(invoice="886218433552", amount=6000, paid_at="2026-09-05 10:00:30", ref="R2")]
    files = [read(statement(rows), "a.xlsx"), read(statement(rows), "b.xlsx"),
            e.read_payment_file("p.csv", payments(pays), "BPR")]
    res = e.run(tmp_path / "base", SEP, files, require_review=False, bank="BPR", report_dir=tmp_path / "r")
    assert res.bank_rows == 1


def test_processed_equals_notified_plus_pending_plus_duplicate(tmp_path):
    rows = [
        ("05/09/2026", "R1", "886218433501", None, "2026-09-05 10:00:00", None, 1000),   # notified
        ("05/09/2026", "R2", "886218433502", None, "2026-09-05 10:00:00", None, 2000),   # duplicate (other channel)
        ("06/09/2026", "R3", "FTCM no invoice", None, "2026-09-06 09:00:00", None, 3000),  # no invoice: pending
        ("07/09/2026", "R4", "886218433504", None, "2026-09-07 09:00:00", None, 4000),   # not found: pending
    ]
    pays = [
        dict(invoice="886218433501", amount=1000, paid_at="2026-09-05 10:00:10", ref="R1"),
        dict(invoice="886218433502", amount=2000, paid_at="2026-09-05 10:00:10", ref="R2", provider="MTN_MOMO"),
    ]
    files = [read(statement(rows)), e.read_payment_file("p.csv", payments(pays), "BPR")]
    first = e.run(tmp_path / "base", SEP, files, require_review=True, bank="BPR", report_dir=tmp_path / "r1")
    filled, needs_a_look = auto_review.classify(first.review, {"886218433504": "Not an IremboPay invoice"})
    e.save_decisions(tmp_path / "base", SEP, filled, "BPR")
    res = e.run(tmp_path / "base", SEP, files, require_review=True, bank="BPR", report_dir=tmp_path / "r2")
    assert not res.stopped_for_review

    monthly = next(t for sheet, _, t in res.blocks if sheet == "Summary-Monthly")

    def total(title):
        t = monthly[title]
        return int(t.loc[t["Currency"].isna(), "Number of Transactions"].iloc[0])

    processed, notified = total("Overall processed payments"), total("Overall notifications received by Irembo")
    pending, duplicate = total("Pending notifications"), total("Duplicate Payment")
    assert (processed, notified, pending, duplicate) == (4, 1, 2, 1)
    assert processed == notified + pending + duplicate


def test_august_2026_matches_the_known_report(tmp_path):
    """The agreed numbers for August 2026: 537 processed, 512 notified, 13 SLA violated, 25 pending, 0 duplicates.
    Needs the real samples: set BANK_RECON_SAMPLES to a folder with Bpr-irembo/ and Bpr/ in it (skipped otherwise).
    The 5 invoices the checkout check still had no answer for are fixed here (not looked up live): the checkout
    page answers change over time, so a live call would not keep matching a report made in the past."""
    if not SAMPLES:
        pytest.skip("BANK_RECON_SAMPLES not set")
    data = Path(SAMPLES)
    folders = {"statements": config.FolderSpec("local", str(data / "Bpr-irembo")),
              "exports": config.FolderSpec("local", str(data / "Bpr"))}
    if not (data / "Bpr-irembo").is_dir() or not (data / "Bpr").is_dir():
        pytest.skip("real samples not found under BANK_RECON_SAMPLES")
    period = e.make_period("month", "2026-08-01")
    found = folder_files.find_files(period, "BPR", {"BPR": folders})
    if not found.statements or not found.exports:
        pytest.skip("real samples not on this computer")
    acc = folder_files.all_files(found)
    first = e.run(tmp_path / "base", period, acc, require_review=True, bank="BPR", report_dir=tmp_path / "r1")
    answers = {inv: "Open, not paid yet" for inv in
              ["876239124919", "886233705544", "886236247937", "886237258713", "886237993535"]}
    filled, needs_a_look = auto_review.classify(first.review, answers)
    e.save_decisions(tmp_path / "base", period, filled, "BPR")
    res = e.run(tmp_path / "base", period, acc, require_review=True, bank="BPR", report_dir=tmp_path / "r2")
    assert not res.stopped_for_review and needs_a_look == []

    monthly = next(t for sheet, _, t in res.blocks if sheet == "Summary-Monthly")

    def total(title):
        t = monthly[title]
        return int(t.loc[t["Currency"].isna(), "Number of Transactions"].iloc[0])

    processed, notified = total("Overall processed payments"), total("Overall notifications received by Irembo")
    sla, pending = total("SLA Violated"), total("Pending notifications")
    duplicate = total("Duplicate Payment")
    assert (processed, notified, sla, pending, duplicate) == (537, 512, 13, 25, 0)
    assert processed == notified + pending + duplicate
