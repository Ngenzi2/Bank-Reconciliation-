"""Layout tests of the report PDF (reconciliation_app/pdf_report.py). The PDF only presents the saved numbers, so the
reports here are built in the shape engine.read_report() returns; no reconciliation runs."""
import re

import numpy as np
import pandas as pd
import pytest
from reportlab import rl_config
from reportlab.pdfbase.pdfmetrics import stringWidth

import pdf_report as P

TITLES = P.TABLE_ORDER


def table(cur, counts, amounts, processed=None, vs=None, unit="Week"):
    rows = [[c, n, a] for c, n, a in zip(cur, counts, amounts)] + [[np.nan, sum(counts), np.nan]]
    t = pd.DataFrame(rows, columns=["Currency", "Number of Transactions", "Total Amount"])
    if processed is not None:
        tot = [*processed, sum(processed)]
        t["%"] = [n / p if p else 0.0 for n, p in zip(t["Number of Transactions"], tot)]
    if vs is not None:
        t[f"VS prev {unit} (%)"] = vs
    return t


def block(cur=("RWF", "USD"), proc=(5342, 77), amt=(108267178, 164828), sla=(19, 0), pend=(0, 0), dup=(0, 0),
          vs=None, unit="Week"):
    rec = [p - x for p, x in zip(proc, pend)]
    return {TITLES[0]: table(cur, proc, amt, vs=vs, unit=unit),
            TITLES[1]: table(cur, rec, amt, proc, vs, unit),
            TITLES[2]: table(cur, sla, [15789 * s for s in sla], proc, vs, unit),
            TITLES[3]: table(cur, pend, [0] * len(cur), proc, vs, unit),
            TITLES[4]: table(cur, dup, [0] * len(cur), proc, vs, unit)}


def report(main=None, weeks=0, name="I&M Bank", not_found=10, week_block=None):
    blocks = [("Summary-Weekly", f"Week {i + 1}: {i * 7 + 1:02d}/08/2026 - {i * 7 + 7:02d}/08/2026",
               week_block or block(proc=(274, 4), amt=(3736768, 400), sla=(1, 0))) for i in range(weeks)]
    blocks.append(("Summary-Monthly", "Month: August 2026 (01/08/2026 - 31/08/2026)", main or block()))
    return {"report_type": "month", "bank": name, "full_name": name, "title": "Month · August 2026",
            "blocks": blocks, "not_found": not_found, "reconciled_on": "Sunday 11/10/2026"}


@pytest.fixture
def pdf_text(monkeypatch):
    """Build a PDF with uncompressed page streams and return its raw text, so the drawn strings can be searched."""
    monkeypatch.setattr(rl_config, "pageCompression", 0)
    monkeypatch.setattr(rl_config, "invariant", 1)

    def build(rep):
        return P.build_pdf(rep).decode("latin-1")
    return build


def pages(text):
    return len(re.findall(r"/Type /Page\b", text))


# ---- header, footer, pages ----------------------------------------------------------------------------------------
def test_header_shows_bank_title_period_and_reconciliation_day(pdf_text):
    text = pdf_text(report())
    drawn = re.findall(r"\(((?:[^()\\]|\\.)*)\) Tj", text)               # the strings drawn on the pages
    assert "M BANK" in drawn                                         # reportlab draws "I&" and "M BANK" as two runs
    for s in ("Reconciliation Report", "Sunday 11/10/2026", "Monthly report \\267 August 2026"):
        assert s in drawn, s
    assert "01/08/2026 \\226 31/08/2026" in drawn                     # reporting period (en dash)


def test_page_x_of_y_on_every_page_of_a_multi_page_report(pdf_text):
    text = pdf_text(report(weeks=6))
    n = pages(text)
    assert n >= 3
    for i in range(1, n + 1):
        assert f"(Page {i} of {n})" in text


def test_two_weeks_fit_on_one_page():
    """A normal 2-currency week block takes at most half of a later page, so weeks are packed two per page."""
    rep = report(weeks=1)
    parts = [P.section_heading("Week 1", "01/08/2026 – 07/08/2026"), P.Spacer(1, 3 * P.mm),
             *P.tables_block(rep["blocks"][0][2])]
    h = sum(f.wrap(P.USABLE, 10_000)[1] for f in parts)
    frame_h = P.PAGE_H - P.LATER_TOP - P.MARGIN - 2 * P.mm
    assert 2 * (h + 5 * P.mm) <= frame_h


def test_long_bank_name_is_cut_in_the_footer_but_the_period_stays():
    long = "Bank of Kigali Commercial and Investment Banking Corporation Rwanda Public Limited Company"
    cut = P.fit_text(long, "Helvetica-Bold", 8.5, P.USABLE * 0.5)
    assert cut.endswith("…") and stringWidth(cut, "Helvetica-Bold", 8.5) <= P.USABLE * 0.5


def test_long_bank_name_report_builds(pdf_text):
    text = pdf_text(report(name="Bank of Kigali Commercial and Investment Banking Corporation Rwanda Plc", weeks=2))
    assert "BANK OF KIGALI COMMERCIAL" in text and "(Page 1 of" in text


# ---- tables -------------------------------------------------------------------------------------------------------
def test_numbers_formatted_with_thousands_and_total_row_has_no_amount():
    cols, head, rows, totals = P._cells(block()[TITLES[1]])
    assert head == ["Currency", "Transactions", "Amount", "% of processed"]
    assert rows[0] == ["RWF", "5,342", "108,267,178", "100.00%"]
    assert totals == [2] and rows[2][0] == "Total" and rows[2][2] == P.NO_AMOUNT   # currencies never added up


def test_zero_values_are_shown_not_blank():
    rep = report(main=block(proc=(0, 0), amt=(0, 0), sla=(0, 0)), not_found=0)
    _, _, rows, _ = P._cells(rep["blocks"][-1][2][TITLES[2]])
    assert rows[0][1:] == ["0", "0", "0.00%"]
    values = {label: value for label, value, _, _ in P.kpi_values(rep["blocks"][-1][2])}
    assert values["PAYMENTS PROCESSED"] == "0" and values["PENDING NOTIFICATIONS"] == "0"


@pytest.mark.parametrize("main", [
    block(),
    block(cur=("RWF", "USD", "EUR"), proc=(1234567, 98765, 4321), amt=(987654321098765.0, 12345678901.0, 99999999.0),
          sla=(12345, 678, 9), pend=(1500, 3, 0), dup=(77, 0, 1), vs=["12.50%", "-3.25%", "N/A", "100.00%"]),
])
def test_no_cell_is_clipped(main):
    """Every value fits inside its column (font width + padding), in two-column or full-width layout."""
    half = (P.USABLE - P.GAP) / 2
    for title in TITLES:
        width = half if P.fits(main[title], half) else P.USABLE
        tb = P.summary_table(title, main[title], width)
        assert sum(tb._colWidths) <= width + 0.01
        _, _, rows, _ = P._cells(main[title])
        for r in rows:
            for text, w in zip(r, tb._colWidths):
                assert stringWidth(text, "Helvetica-Bold", P.BODY) + 2 * P.PAD <= w, (title, text)


def test_large_amounts_switch_to_full_width_tables():
    big = block(amt=(987654321098765.0, 12345678901.0), vs=["12.50%", "-3.25%", "N/A"])
    half = (P.USABLE - P.GAP) / 2
    assert not P.fits(big[TITLES[1]], half)
    assert P.fits(block()[TITLES[1]], half)


def test_number_columns_are_right_aligned_and_total_bold():
    tb = P.summary_table(TITLES[0], block()[TITLES[0]], 250)
    style = tb._cellStyles
    assert style[2][1].alignment == "RIGHT" and style[2][2].alignment == "RIGHT" and style[2][0].alignment == "LEFT"
    assert style[-1][0].fontname == "Helvetica-Bold" and style[2][0].fontname == "Helvetica"


def test_title_and_header_rows_repeat_on_a_new_page():
    tb = P.summary_table(TITLES[0], block()[TITLES[0]], P.USABLE)
    assert tb.repeatRows == 2


# ---- key figures --------------------------------------------------------------------------------------------------
def test_notified_and_sla_are_separate_figures():
    values = {label: (value, sub, alert) for label, value, sub, alert in P.kpi_values(block())}
    assert values["NOTIFIED"][0] == "100.00%"        # received / processed, as saved
    assert values["SLA VIOLATIONS"][0] == "19"
    assert values["SLA VIOLATIONS"][1].startswith("0.35% of processed")
    assert values["SLA VIOLATIONS"][2] is True                         # highlighted


def test_sla_card_not_highlighted_without_violations():
    values = {label: alert for label, _, _, alert in P.kpi_values(block(sla=(0, 0)))}
    assert values["SLA VIOLATIONS"] is False


def test_heading_parsing():
    assert P.split_heading("Month: August 2026 (01/08/2026 - 31/08/2026)") == ("Month", "August 2026",
                                                                               "01/08/2026 – 31/08/2026")
    assert P.split_heading("Week 1: 01/08/2026 - 02/08/2026") == ("", "Week 1", "01/08/2026 – 02/08/2026")


def test_report_pdf_reads_the_saved_report(monkeypatch, tmp_path):
    path = tmp_path / "IM_recon_summary_01-08-2026_to_31-08-2026.xlsx"
    path.write_bytes(b"x")
    monkeypatch.setattr(P.engine, "read_report", lambda p: report())
    assert P.report_pdf(path).startswith(b"%PDF")


# ---- the custom report (step 6) -----------------------------------------------------------------------------------
def custom_content(orientation="Portrait", density="Comfortable", invoices=3):
    """The shape engine.custom_content() returns: overview + a table section, a chart section and a list section."""
    summary = pd.DataFrame({"Currency": ["RWF", "USD", "Total"], "Transactions": [5342, 77, 5419],
                            "Total amount": [108267178.0, 164828.0, np.nan], "% of processed": [0.0036, 0.0, 0.0035]})
    inv = pd.DataFrame({"Value Date": pd.to_datetime(["2026-08-03"] * invoices), "Currency": ["RWF"] * invoices,
                        "Invoice": [f"88621568{i:04d}" for i in range(invoices)], "Amount": [500.0] * invoices,
                        "Decision": ["Not reviewed"] * invoices, "Comment": ["a long comment " * 4] * invoices})
    base = {"table": None, "series": [], "group": "", "currency": "All", "list": None, "list_title": ""}
    return {"title": "I&M Bank Reconciliation Report", "period": "August 2026", "type": "Full Report",
            "orientation": orientation, "density": density, "stamp": "Reconciled on 11 Oct 2026 · 14:03",
            "kpis": [("Processed", "5,419"), ("Matched", "5,409"), ("Exceptions", "10"), ("Match rate", "99.82%")],
            "currency_lines": ["RWF · 5,342 rows · 5,332 matched"], "exceptions_line": "",
            "file_stem": "IM_full_report",
            "sections": [{**base, "key": "sla", "title": "SLA Violations", "as": "Table", "table": summary},
                         {**base, "key": "received", "title": "Notifications Received", "as": "Bar chart",
                          "series": [("Week 1", 278), ("Week 2", 1581)], "group": "Week"},
                         {**base, "key": "pending", "title": "Pending Notifications", "as": "Table", "list": inv,
                          "list_title": "Pending invoices"}]}


@pytest.mark.parametrize("orientation", ["Portrait", "Landscape"])
@pytest.mark.parametrize("density", ["Comfortable", "Compact"])
def test_custom_report_builds_in_every_layout(pdf_text, monkeypatch, orientation, density):
    monkeypatch.setattr(rl_config, "pageCompression", 0)
    text = P.custom_pdf(custom_content(orientation, density)).decode("latin-1")
    drawn = re.findall(r"\(((?:[^()\\]|\\.)*)\) Tj", text)
    assert "Reconciliation Report" in drawn and "11 Oct 2026 \\267 14:03" in drawn
    n = pages(text)
    assert all(f"(Page {i} of {n})" in text for i in range(1, n + 1))
    box = re.search(r"/MediaBox \[ 0 0 ([\d.]+) ([\d.]+) \]", text)
    assert (float(box.group(1)) > float(box.group(2))) == (orientation == "Landscape")


def test_custom_report_long_invoice_list_spans_pages_with_page_numbers(monkeypatch):
    monkeypatch.setattr(rl_config, "pageCompression", 0)
    text = P.custom_pdf(custom_content(invoices=120)).decode("latin-1")
    n = pages(text)
    assert n >= 3 and all(f"(Page {i} of {n})" in text for i in range(1, n + 1))
    assert text.count("(Value Date) Tj") >= n - 1 >= 2              # the list's header row repeats on each of its pages


def test_custom_table_total_row_keeps_currencies_apart():
    content = custom_content()
    tb = P.data_table(content["sections"][0]["table"], 400, 8, 4, total_row=True)
    assert tb._cellvalues[-1][0] == "Total" and tb._cellvalues[-1][2] == P.NO_AMOUNT
    assert tb._cellvalues[1][1] == "5,342" and tb._cellvalues[1][3] == "0.36%"
    assert tb._cellStyles[-1][0].fontname == "Helvetica-Bold"
    assert tb.repeatRows == 1


def test_after_the_report_reviewed_rows_are_not_counted_again(pdf_text):
    """The rows reviewed in Exception Resolution are in the table their decision chose: the report PDF has no
    'not found' figure and never says 'exception'."""
    labels = [label for label, _, _, _ in P.kpi_values(block())]
    assert labels == ["PAYMENTS PROCESSED", "NOTIFIED", "SLA VIOLATIONS", "PENDING NOTIFICATIONS",
                      "DUPLICATE PAYMENTS"]
    text = pdf_text(report(weeks=2)).lower()
    assert "exception" not in text and "not found" not in text
