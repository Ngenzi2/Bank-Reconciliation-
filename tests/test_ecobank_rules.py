"""Ecobank (26/10/2026): processed = the requests Ecobank answered SUCCESS (Cba Rsp Code 000), by Request Date,
each invoice once - the first SUCCESS counts, a second SUCCESS (debited twice) goes to Exception Resolution;
received = paid_at filled through ECOBANK; another provider = Duplicate Payment."""
import io

import pandas as pd

from banks import ecobank


def test_success_requests_kept_first_counts_second_is_reviewed():
    rows = pd.DataFrame({
        "invoice_number": ["A", "A", "A", "B", "C"],
        "Cba Rsp Code": ["E92", "000", "000", "E92", "000"],
        "Request Date": pd.to_datetime(["2026-08-03 08:00", "2026-08-03 09:00", "2026-08-03 10:00",
                                        "2026-08-04 08:00", "2026-08-05 08:00"]),
        "Amount": [100, 100, 100, 50, 70], "Cba Ref No": ["r0", "r1", "r2", "r3", "r4"]})
    pays = pd.DataFrame({"invoice_number": ["A", "B", "C"], "amount": [100, 50, 70]})
    kept = ecobank.one_row_per_invoice(rows, pays)
    assert kept["Cba Ref No"].tolist() == ["r1", "r2", "r3", "r4"]      # A: its 2 SUCCESS; B: its only try
    kept = kept.assign(found_in_irembopay=True, payment_provider=["ECOBANK", "ECOBANK", "ECOBANK", "MTN"],
                       paid_at=pd.to_datetime(["2026-08-03 09:00"] * 3 + ["2026-08-05 08:00"]),
                       payment_status="PAID", sla_status="On time")
    reasons = ecobank.review_reasons(kept)
    assert reasons.tolist() == ["", "Invoice credited again: debited again by Ecobank; the request counted is r1", "", ""]
    kept["is_extra_credit"] = reasons.ne("").values
    t = ecobank.tables(kept, rows, pd.Timestamp("2026-08-03"), pd.Timestamp("2026-08-10"))
    assert t["processed"].tolist() == [True, False, False, True]          # A once; B was not debited (E92)
    assert t["received"].tolist() == [True, False, False, False]
    assert t["duplicate"].tolist() == [False, False, False, True]         # C: paid through MTN


def test_the_period_is_the_request_date():
    rows = pd.DataFrame({"Request Date": pd.to_datetime(["2026-08-02 23:59", "2026-08-03 00:01"]),
                         "invoice_number": ["A", "B"], "Cba Rsp Code": ["000", "000"], "Cba Ref No": ["r1", "r2"]})
    assert ecobank.in_period(rows, None, pd.Timestamp("2026-08-03"), pd.Timestamp("2026-08-10")).tolist() == [False, True]


def test_a_resend_with_the_same_cba_ref_is_one_debit():
    """26/10/2026: the same invoice and the same Cba Ref No again (another Request Date) = Ecobank sent the same
    debit again, not a second debit; a different Cba Ref No is still an extra debit.
    13/11/2026: the re-send counts once in the period - only the rows of the period are compared."""
    rows = pd.DataFrame({
        "invoice_number": ["A", "A", "A", "B", "B"],
        "Cba Rsp Code": ["000"] * 5,
        "Request Date": pd.to_datetime(["2026-10-19 16:02", "2026-10-20 10:00", "2026-10-21 02:00",
                                        "2026-10-20 14:18", "2026-10-20 14:22"]),
        "Amount": [500] * 3 + [1200] * 2, "Cba Ref No": ["ITR1", "ITR1", "ITR1", "D81", "D13"]})
    pays = pd.DataFrame({"invoice_number": ["A", "B"], "amount": [500, 1200]})
    kept = ecobank.one_row_per_invoice(rows, pays).reset_index(drop=True)
    assert kept["Cba Ref No"].tolist() == ["ITR1", "ITR1", "D81", "D13", "ITR1"]   # every SUCCESS, by Request Date
    week = kept[ecobank.in_period(kept, kept, pd.Timestamp("2026-10-19"), pd.Timestamp("2026-10-26"))]
    assert week["Cba Ref No"].tolist() == ["ITR1", "D81", "D13"]          # A once; B: two different debits
    assert week["Attempts"].tolist() == [3, 2, 2]                          # the resends are still counted as tries
    reasons = ecobank.review_reasons(week.assign(found_in_irembopay=True))
    assert reasons.tolist() == ["", "", "Invoice credited again: debited again by Ecobank; the request counted is D81"]
    # A was first sent on 19/10: a period that starts on 20/10 still counts it, once (the 20/10 row, not 21/10)
    later = kept[ecobank.in_period(kept, kept, pd.Timestamp("2026-10-20"), pd.Timestamp("2026-10-27"))]
    assert later["Request Date"].dt.day.tolist() == [20, 20, 20]


def test_the_table_decided_for_an_extra_debit_is_read():
    """26/10/2026: an extra debit is not in Processed, but the table chosen for it in Exception Resolution counts."""
    import engine
    rows = pd.DataFrame({"is_extra_credit": [False, True, True, True, True],
                         "invoice_number": ["A", "A", "A", "A", "A"],
                         "review_table": ["", "Paid with other channel", "Pending notification", "Exclude", ""]})
    none = pd.Series(False, index=rows.index)
    m = {"processed": rows["is_extra_credit"].eq(False), "received": rows["is_extra_credit"].eq(False),
         "sla": none, "pending": none, "duplicate": none}
    out = engine.with_extra_decisions(rows, m, pd.Series(True, index=rows.index))
    assert out["processed"].tolist() == [True, False, False, False, False]      # each invoice once
    assert out["duplicate"].tolist() == [False, True, False, False, False]      # a double payment
    assert out["pending"].tolist() == [False, False, True, False, False]
    assert out["received"].tolist() == [True, False, False, False, False]
    outside = engine.with_extra_decisions(rows, m, pd.Series(False, index=rows.index))
    assert not outside["duplicate"].any()                                       # only the period's rows


def test_an_invoice_counts_once_per_table_and_the_report_says_where_each_credit_went():
    """26/10/2026: a second debit put in a table that already has its invoice is not counted again; each credit of
    the invoice gets a note: why it is in its table and where the other one went."""
    import engine
    rows = pd.DataFrame({"invoice_number": ["X", "X", "Y", "Y", "Y"],
                         "is_extra_credit": [False, True, False, True, True],
                         "review_table": ["Pending notification", "Pending notification", "Paid with other channel",
                                          "Paid with other channel", "Paid with other channel"],
                         "Cba Ref No": ["D412", "D838", "D660", "D663", "D089"],
                         "Request Date": pd.to_datetime(["2026-10-20 19:47", "2026-10-21 19:03", "2026-10-21 18:47",
                                                         "2026-10-21 19:29", "2026-10-22 00:12"])})
    none = pd.Series(False, index=rows.index)
    m = {"processed": ~rows["is_extra_credit"], "received": none, "sla": none,
         "pending": pd.Series([True, False, False, False, False]),
         "duplicate": pd.Series([False, False, True, False, False])}
    out = engine.with_extra_decisions(rows, m, pd.Series(True, index=rows.index))
    assert out["pending"].tolist() == [True, False, False, False, False]          # X once in Pending
    assert out["duplicate"].tolist() == [False, False, True, False, False]        # Y once in Duplicate
    notes = engine.table_notes(rows, out, "ECO")
    assert notes[0].startswith("In Pending notification: the credit counted") and "D838" in notes[0]
    assert notes[1].startswith("Not counted: the invoice is already in Pending notification")
    assert notes[4].startswith("Not counted: the invoice is already in Duplicate payment")

    # X's second debit decided "Paid with other channel": X in Pending once AND Duplicate once, each says why
    rows.loc[1, "review_table"] = "Paid with other channel"
    out = engine.with_extra_decisions(rows, m, pd.Series(True, index=rows.index))
    assert out["duplicate"].tolist() == [False, True, True, False, False]
    notes = engine.table_notes(rows, out, "ECO")
    assert "D838 of 21/10 19:03 -> Duplicate payment" in notes[0]
    assert notes[1].startswith("In Duplicate payment: second credit") and "D412 of 20/10 19:47 -> Pending" in notes[1]


def test_review_reasons_prefers_the_credit_irembopay_recorded():
    """2026-11-10 (rule 4, PREFER_RECORDED_REFERENCE removed again - fixed directly in review_reasons): a second
    debit with a different Cba Ref No - the one matching IremboPay's own provider_reference is kept, not just
    whichever came in first. No match for either: falls back to the earliest, as before."""
    rows = pd.DataFrame({
        "invoice_number": ["A", "A", "A", "B", "B"],
        "Cba Rsp Code": ["000"] * 5,
        "Request Date": pd.to_datetime(["2026-10-26 16:19", "2026-10-28 09:05", "2026-10-29 12:26",
                                        "2026-10-20 10:00", "2026-10-21 10:00"]),
        "Cba Ref No": ["RW1271529910", "RW1271705012", "ITHERPPGTR261029", "X1", "X2"],
        "provider_reference": ["ITHERPPGTR261029"] * 3 + ["", ""]})
    reasons = ecobank.review_reasons(rows)
    # A: the third row (matches provider_reference) is kept; the other two are "credited again"
    assert reasons[0] != "" and "ITHERPPGTR261029" in reasons[0]
    assert reasons[1] != "" and "ITHERPPGTR261029" in reasons[1]
    assert reasons[2] == ""
    # B: IremboPay recorded no reference for either credit - falls back to the earliest (X1)
    assert reasons[3] == ""
    assert reasons[4] != "" and "X1" in reasons[4]


def test_no_invoice_credit_is_pending_automatically_but_stays_in_review():
    """2026-11-10 (NO_INVOICE_PENDING_REVIEWED): a bank credit with no real invoice number is counted as Pending
    straight away - but, unlike BPR's NO_INVOICE_PENDING, it is not removed from Exception Resolution, so a
    person can still see and override it."""
    import engine
    assert engine.has_invoice_number("886299919186") is True
    assert engine.has_invoice_number("IREMBO|27-Oct-2026|000162") is False
    assert engine.has_invoice_number(None) is False

    merged = pd.DataFrame({
        "period": ["current", "current"],
        "invoice_number": ["IREMBO|27-Oct-2026|000162", "886200000000"],
        "found_in_irembopay": [False, False], "payment_status": [None, None],
        "paid_at": pd.to_datetime([None, None]),
        "Currency": ["RWF", "RWF"], "Request Date": pd.to_datetime(["2026-10-27", "2026-10-27"])})
    merged["_key"] = engine.key_of(merged, "ECO")
    merged["_to_review"] = True
    review = pd.DataFrame({
        "_key": merged["_key"], "invoice_number": merged["invoice_number"],
        "Reviewed (Yes/No)": ["", ""], "Report table": ["", ""], "Correct invoice number": ["", ""],
        "payment_status": ["", ""], "payment_provider": ["", ""],
        "paid_at (YYYY-MM-DD HH:MM:SS)": ["", ""], "Merchant name": ["", ""]})
    payments = pd.DataFrame({"invoice_number": pd.Series([], dtype=str), "amount": pd.Series([], dtype=float)})
    out, _, counts, _ = engine.apply_review(merged, merged.copy(), payments, review, "ECO")

    no_inv = out[out["invoice_number"] == "IREMBO|27-Oct-2026|000162"].iloc[0]
    assert no_inv["review_status"] == "Pending notification (no invoice number)"
    assert no_inv["payment_status"] == "NEW"
    real = out[out["invoice_number"] == "886200000000"].iloc[0]
    assert real["review_status"] == "Not reviewed"              # a real, unresolved invoice: unaffected
    assert counts["Not reviewed"] == 1                           # only the real invoice blocks require_review


def test_pending_but_paid_in_irembopay_since_gets_a_note_and_stays_pending():
    """28/10/2026: NEW in the payment file (Pending) but the checkout now says Paid: a note, still in Pending."""
    import engine
    rows = pd.DataFrame({"IremboPay check": ["Paid", "Expired", "Paid"], "Report note": ["", "", "earlier note"]})
    whole = {"pending": pd.Series([True, True, False]), "received": pd.Series([False, False, True])}
    notes = engine.checkout_notes(rows, whole)
    assert notes[0] == engine.PAID_SINCE and notes[1] == "" and notes[2] == "earlier note"   # row 2: not in Pending
    assert whole["pending"].tolist() == [True, True, False]                                    # nothing moved


def test_currency_from_checkout_only_for_a_credit_not_paid_in_the_file():
    """13/11/2026 (CURRENCY_FROM_CHECKOUT): the invoice's own currency replaces the statement's only when the
    payment file does not show the credit as paid and the checkout gave one."""
    import engine
    rows = pd.DataFrame({"invoice_number": ["A", "B", "C", None], "Currency": ["RWF"] * 4,
                         "paid_at": pd.to_datetime(["2026-10-27 10:00", None, None, None])})
    assert engine.currency_from_checkout(rows, {"A": "USD", "B": "USD"}).tolist() == ["RWF", "USD", "RWF", "RWF"]


def _made_up_week(tmp_path, more=(), **kw):
    """One made-up Ecobank week (26/10 - 01/11/2026): A paid through Ecobank, B still NEW (USD 1 in IremboPay, sent
    as RWF 1 by the bank), C paid through MTN. more: other statement rows (invoice, amount, day of October, ref)."""
    import engine
    cols = ["Source Channel", "Biller Code", "Bill Ref No", "Ccy", "Amount", "Request Date", "Cba Rsp Code",
            "Cba Rsp Msg", "Biller Rsp Code", "Biller Rsp Msg", "Cba Ref No", "Customer Ref No"]
    row = lambda inv, amt, day, ref: ["AGENCYBANKING", "IREMBO_PAY", inv, "RWF", amt, f"2026-10-{day} 08:00:00",
                                      "000", "SUCCESS", "000", "Success", ref, "250700000000"]
    buf = io.BytesIO()
    pd.DataFrame([[None] * len(cols)] * 3 + [cols, row("886200000001", 5000, 27, "R1"),
                                             row("886200000002", 1, 28, "R2"),
                                             row("886200000003", 3000, 29, "R3")]
                 + [row(*m) for m in more]).to_excel(buf, header=False, index=False)
    pays = pd.DataFrame({
        "invoice_number": ["886200000001", "886200000002", "886200000003"], "amount": [5000, 1, 3000],
        "currency": ["RWF", "USD", "RWF"], "payment_status": ["PAID", "NEW", "PAID"],
        "paid_at": ["2026-10-27 10:00:05", None, "2026-10-29 09:30:00"],
        "payment_provider": ["ECOBANK", None, "MTN"], "transaction_status": ["SUCCESSFUL", None, "SUCCESSFUL"],
        "created_at": ["2026-10-27 09:00:00", "2026-10-20 09:00:00", "2026-10-29 09:00:00"],
        "name": ["RRA", "GOR", "RSSB"]})
    files = [engine.read_bank_file("IREMBO PAY2026-10-30.xlsx", buf.getvalue(), "ECO"),
             engine.read_payment_file("payments.csv", pays.to_csv(index=False).encode(), "ECO")]
    res = engine.run(tmp_path, engine.make_period("week", "2026-10-26"), files, bank="ECO", report_dir=tmp_path / "r",
                     checkout={"886200000002": "Open, not paid yet"}, **kw)
    tables = next(b for b in res.blocks if b[0] == "Summary-Weekly")[2]

    def count(table, currency=None):
        t = tables[table]
        return int(t[t["Currency"].isna() if currency is None else t["Currency"].eq(currency)]
                   ["Number of Transactions"].iloc[0])
    return res, count


def test_a_pending_invoice_takes_its_currency_from_the_checkout(tmp_path):
    """13/11/2026: B is USD 1 in IremboPay but RWF 1 on the Ecobank statement: with the checkout's currency it is
    counted under USD; the statement sheet keeps what the bank wrote; Processed = Notified + Pending + Duplicate."""
    res, count = _made_up_week(tmp_path / "a", checkout_currency={"886200000002": "USD"})
    assert count("Pending notifications", "USD") == 1 and count("Pending notifications", "RWF") == 0
    assert count("Overall processed payments", "USD") == 1 and count("Overall processed payments", "RWF") == 2
    assert count("Overall processed payments") == 3 == (count("Overall notifications received by Irembo")
                                                        + count("Pending notifications") + count("Duplicate Payment"))
    pending = res.sheets["Pending notification"]
    assert pending["Currency"].tolist() == ["USD"] and pending["Amount"].tolist() == [1]
    assert pending["Report note"].iloc[0].startswith("Currency from IremboPay's checkout (USD); the bank statement "
                                                     "says RWF")
    statement = res.sheets["Ecobank statement"].set_index("invoice_number")
    assert statement["Currency"]["886200000002"] == "RWF"

    res, count = _made_up_week(tmp_path / "b")                      # no currency from the checkout: as before
    assert count("Pending notifications", "RWF") == 1 and count("Pending notifications", "USD") == 0
    assert count("Overall processed payments") == 3 == (count("Overall notifications received by Irembo")
                                                        + count("Pending notifications") + count("Duplicate Payment"))


def test_a_resend_in_the_week_counts_whichever_earlier_statements_are_loaded(tmp_path):
    """13/11/2026: B was first sent on 20/10 and sent again (same Cba Ref No) on 28/10 and 30/10. The week of 26/10
    counts it once - the same figures with or without the statement of the earlier week."""
    again = [("886200000002", 1, 30, "R2")]
    _, alone = _made_up_week(tmp_path / "a", more=again)
    _, with_earlier = _made_up_week(tmp_path / "b", more=again + [("886200000002", 1, 20, "R2")])
    for count in (alone, with_earlier):
        assert count("Overall processed payments") == 3 and count("Pending notifications") == 1
        assert count("Overall processed payments") == (count("Overall notifications received by Irembo")
                                                       + count("Pending notifications") + count("Duplicate Payment"))
