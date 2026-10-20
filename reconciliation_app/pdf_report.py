"""PDF of a saved reconciliation report (presentation only).

The numbers come from engine.read_report() exactly as they are saved in the Excel report; nothing is recalculated
here, the tables are only laid out for A4 printing:
  page 1  = header band, key figures, the main summary (5 tables + how to read them)
  then    = one block per week (as many as fit on a page; a block is never split over two pages)
Every page has a footer with "Page X of Y"; pages after the first repeat a slim header.
"""
import io
import re
from pathlib import Path
from xml.sax.saxutils import escape

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfgen import canvas as rl_canvas
from reportlab.platypus import (BaseDocTemplate, Frame, KeepTogether, LongTable, NextPageTemplate, PageTemplate,
                                Paragraph, Spacer, Table, TableStyle)

import engine

# ---- look ---------------------------------------------------------------------------------------------------------
TEAL, GREEN = colors.HexColor("#123239"), colors.HexColor("#99E378")
INK, MUTED = colors.HexColor("#1D2B2F"), colors.HexColor("#5B6B70")
RULE, HAIRLINE = colors.HexColor("#9AA8AD"), colors.HexColor("#E1E8EA")
TITLE_FILL, TOTAL_FILL = colors.HexColor("#E8F0F2"), colors.HexColor("#F2F6F7")
WARN_FILL, WARN_STROKE, WARN_INK = colors.HexColor("#FFF3D6"), colors.HexColor("#D4A72C"), colors.HexColor("#7A4E00")
BAND_SUB = colors.HexColor("#C9D9DC")

PAGE_W, PAGE_H = A4
MARGIN = 15 * mm
USABLE = PAGE_W - 2 * MARGIN
FOOTER_Y = 9 * mm
LATER_TOP = 24 * mm                       # room for the slim header on pages 2+
GAP = 6 * mm                              # between the two table columns
BODY = 8.0                                # table font size
PAD = 5                                   # left/right cell padding

TABLE_ORDER = ["Overall processed payments", "Overall notifications received by Irembo", "SLA Violated",
               "Pending notifications", "Duplicate Payment"]
HEADERS = {"Currency": "Currency", "Number of Transactions": "Transactions", "Total Amount": "Amount",
           "%": "% of processed"}
NO_AMOUNT = "—"                           # the total row has no amount: currencies are never added together


def _style(name, **kw):
    base = {} if "parent" in kw else dict(fontName="Helvetica", fontSize=BODY, leading=BODY + 2, textColor=INK)
    base.update(kw)
    return ParagraphStyle(name, **base)


S = {
    "bank": _style("bank", fontName="Helvetica-Bold", fontSize=8.5, leading=11, textColor=GREEN),
    "title": _style("title", fontName="Helvetica-Bold", fontSize=17, leading=21, textColor=colors.white),
    "band_sub": _style("band_sub", fontSize=8.5, leading=11, textColor=BAND_SUB),
    "meta_k": _style("meta_k", fontSize=7, leading=9, textColor=BAND_SUB),
    "meta_v": _style("meta_v", fontName="Helvetica-Bold", fontSize=8.5, leading=11, textColor=colors.white),
    "h2": _style("h2", fontName="Helvetica-Bold", fontSize=11, leading=14, textColor=TEAL),
    "h2_note": _style("h2_note", fontSize=8, leading=10, textColor=MUTED),
    "kpi_label": _style("kpi_label", fontName="Helvetica-Bold", fontSize=7.2, leading=9, textColor=MUTED),
    "kpi_value": _style("kpi_value", fontName="Helvetica-Bold", fontSize=17, leading=20, textColor=TEAL),
    "kpi_sub": _style("kpi_sub", fontSize=7.4, leading=9.2, textColor=MUTED),
    "t_title": _style("t_title", fontName="Helvetica-Bold", fontSize=8.6, leading=10.5, textColor=TEAL),
    "th": _style("th", fontName="Helvetica-Bold", fontSize=7.2, leading=8.6, textColor=MUTED),
    "th_r": _style("th_r", fontName="Helvetica-Bold", fontSize=7.2, leading=8.6, textColor=MUTED, alignment=2),
    "note_h": _style("note_h", fontName="Helvetica-Bold", fontSize=8.2, leading=10.5, textColor=TEAL),
    "note": _style("note", fontSize=7.3, leading=9.4, textColor=MUTED),
}


# ---- small helpers ------------------------------------------------------------------------------------------------
def _missing(v) -> bool:
    return v is None or (not isinstance(v, str) and pd.isna(v))


def fmt_count(v) -> str:
    return f"{int(v):,}"


def fmt_amount(v) -> str:
    return NO_AMOUNT if _missing(v) else f"{v:,.0f}"


def fmt_pct(v) -> str:
    return f"{float(v):.2%}"


def _total(t: pd.DataFrame):
    """(transactions, % or None) of the Total row of one summary table."""
    row = t[t["Currency"].isna()]
    if row.empty:
        return 0, None
    pct = row["%"].iloc[0] if "%" in t else None
    return int(row["Number of Transactions"].iloc[0]), (None if _missing(pct) else float(pct))


def _currency_rows(t: pd.DataFrame):
    return [(c, int(n)) for c, n in zip(t["Currency"], t["Number of Transactions"]) if not _missing(c)]


def split_heading(heading: str):
    """'Month: August 2026 (01/08/2026 - 31/08/2026)' -> ('Month', 'August 2026', '01/08/2026 – 31/08/2026')."""
    m = re.match(r"^\s*([^:]+):\s*(.*?)\s*\(([^()]*)\)\s*$", str(heading))
    if not m:                                                 # 'Week 1: 01/08/2026 - 02/08/2026'
        head, _, rest = str(heading).partition(":")
        if rest and re.search(r"\d{2}/\d{2}/\d{4}", rest):
            return "", head.strip(), rest.replace(" - ", " – ").strip()
        return "", str(heading), ""
    return m.group(1).strip(), m.group(2).strip(), m.group(3).replace(" - ", " – ").strip()


def fit_text(text: str, font: str, size: float, width: float) -> str:
    """Cut text with '…' so it fits in width (the footer never runs into the page number)."""
    if stringWidth(text, font, size) <= width:
        return text
    while text and stringWidth(text + "…", font, size) > width:
        text = text[:-1]
    return text.rstrip() + "…"


# ---- one summary table --------------------------------------------------------------------------------------------
def _cells(t: pd.DataFrame):
    """Header labels and text rows of one table (numbers formatted, Total row last as it is saved)."""
    cols = list(t.columns)
    head = [HEADERS.get(c, c.replace("VS prev", "vs prev.").replace(" (%)", "")) for c in cols]
    rows, total_rows = [], []
    for i, rec in enumerate(t.itertuples(index=False)):
        out = []
        for c, v in zip(cols, rec):
            if c == "Currency":
                out.append("Total" if _missing(v) else str(v))
                if _missing(v):
                    total_rows.append(i)
            elif c == "Number of Transactions":
                out.append(fmt_count(v))
            elif c == "Total Amount":
                out.append(fmt_amount(v))
            elif c == "%":
                out.append(fmt_pct(v))
            else:
                out.append("" if _missing(v) else str(v))
        rows.append(out)
    return cols, head, rows, total_rows


def needed_widths(t: pd.DataFrame):
    """Narrowest width of each column that shows every value on one line (headers may wrap on word breaks)."""
    cols, head, rows, _ = _cells(t)
    widths = []
    for j, h in enumerate(head):
        w_vals = max([stringWidth(r[j], "Helvetica-Bold", BODY) for r in rows] + [0])
        w_head = max(stringWidth(word, "Helvetica-Bold", 7.2) for word in h.split())
        widths.append(max(w_vals, w_head) + 2 * PAD + 2)
    return widths


def fits(t: pd.DataFrame, width: float) -> bool:
    return sum(needed_widths(t)) <= width


def summary_table(title: str, t: pd.DataFrame, width: float, warning: bool = False) -> LongTable:
    """One summary table: title row, header row (both repeat on a new page), currency rows, bold Total row."""
    cols, head, rows, total_rows = _cells(t)
    need = needed_widths(t)
    numeric = [j for j, c in enumerate(cols) if c != "Currency"] or list(range(len(cols)))
    first = max(need[0], 0.2 * width) if 0 not in numeric else 0
    even = (width - first) / len(numeric)
    if all(even >= need[j] for j in numeric):                 # balanced: every number column gets the same width
        widths = [even if j in numeric else first for j in range(len(cols))]
    else:                                                     # one column needs more room: give the rest evenly
        extra = max(0.0, width - sum(need))
        widths = [w + (extra / len(numeric) if j in numeric else 0) for j, w in enumerate(need)]
    if sum(widths) > width:                                   # cannot happen when fits() was checked; shrink evenly
        widths = [w * width / sum(widths) for w in widths]

    title_style = S["t_title"] if not warning else _style("t_title_w", parent=S["t_title"], textColor=WARN_INK)
    data = [[Paragraph(escape(title), title_style)] + [""] * (len(cols) - 1),
            [Paragraph(escape(h), S["th"] if c == "Currency" else S["th_r"]) for h, c in zip(head, cols)]]
    data += rows
    tb = LongTable(data, colWidths=widths, repeatRows=2, hAlign="LEFT")
    ink = WARN_INK if warning else INK
    style = [
        ("SPAN", (0, 0), (-1, 0)),
        ("BACKGROUND", (0, 0), (-1, 0), WARN_FILL if warning else TITLE_FILL),
        ("LINEBEFORE", (0, 0), (0, 0), 3, WARN_STROKE if warning else TEAL),
        ("BOX", (0, 0), (-1, -1), 0.5, WARN_STROKE if warning else RULE),
        ("LINEBELOW", (0, 1), (-1, 1), 0.8, WARN_STROKE if warning else TEAL),
        ("LINEBELOW", (0, 2), (-1, -2), 0.3, HAIRLINE),
        ("FONTNAME", (0, 2), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 2), (-1, -1), BODY),
        ("TEXTCOLOR", (0, 2), (-1, -1), ink),
        ("ALIGN", (1, 2), (-1, -1), "RIGHT"),
        ("ALIGN", (0, 2), (0, -1), "LEFT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), PAD),
        ("RIGHTPADDING", (0, 0), (-1, -1), PAD),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING", (0, 0), (-1, 0), 4),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 4),
    ]
    for i in total_rows:
        r = i + 2
        style += [("FONTNAME", (0, r), (-1, r), "Helvetica-Bold"), ("BACKGROUND", (0, r), (-1, r), TOTAL_FILL),
                  ("LINEABOVE", (0, r), (-1, r), 0.8, WARN_STROKE if warning else TEAL)]
    tb.setStyle(TableStyle(style))
    return tb


# ---- page 1 blocks ------------------------------------------------------------------------------------------------
def band(bank: str, subtitle: str, meta: list, width: float = USABLE, title: str = "Reconciliation Report") -> Table:
    """Dark teal band: bank, report title and subtitle on the left; [(label, value)] on the right."""
    left = [Paragraph(escape(bank.upper()), S["bank"]), Paragraph(escape(title), S["title"]),
            Paragraph(escape(subtitle), S["band_sub"])]
    right = []
    for k, v in meta:
        if v:
            right += [Paragraph(escape(k.upper()), S["meta_k"]), Paragraph(escape(v), S["meta_v"]), Spacer(1, 3)]
    right_w = 62 * mm
    t = Table([[left, right]], colWidths=[width - right_w, right_w])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), TEAL),
        ("LINEBEFORE", (0, 0), (0, -1), 4, GREEN),
        ("LINEBEFORE", (1, 0), (1, -1), 0.5, colors.HexColor("#2E5058")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 12),
        ("RIGHTPADDING", (0, 0), (-1, -1), 10),
        ("TOPPADDING", (0, 0), (-1, -1), 10),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
    ]))
    return t


def header_band(rep: dict, kind: str, period_name: str, dates: str, reconciled: str) -> Table:
    """The standard report's band: bank, title, period; period dates and reconciliation day on the right."""
    sub = " · ".join(x for x in (kind and f"{kind}ly report".replace("Dayly", "Daily"), period_name) if x)
    return band(rep.get("full_name") or rep.get("bank") or "Bank", sub,
                [("Reporting period", dates or period_name), ("Reconciled on", reconciled)])


def kpi_values(tables: dict) -> list:
    """[(label, value, sub-line, is_alert)] - every number is read from the saved tables, nothing is recomputed.
    The rows reviewed in Exception Resolution are already in the table their decision chose: no separate figure."""
    proc = tables["Overall processed payments"]
    processed, _ = _total(proc)
    received, rate = _total(tables["Overall notifications received by Irembo"])
    sla, sla_pct = _total(tables["SLA Violated"])
    pending, pending_pct = _total(tables["Pending notifications"])
    dup, dup_pct = _total(tables["Duplicate Payment"])
    per_currency = " · ".join(f"{c} {n:,}" for c, n in _currency_rows(proc)) or "No currency rows"

    def pct(p):
        return fmt_pct(p) if p is not None else "—"

    return [
        ("PAYMENTS PROCESSED", fmt_count(processed), per_currency, False),
        ("NOTIFIED", pct(rate), f"{received:,} of {processed:,} notifications received", False),
        ("SLA VIOLATIONS", fmt_count(sla), f"{pct(sla_pct)} of processed payments", sla > 0),
        ("PENDING NOTIFICATIONS", fmt_count(pending), f"{pct(pending_pct)} of processed", False),
        ("DUPLICATE PAYMENTS", fmt_count(dup), f"{pct(dup_pct)} of processed · other channel", False),
    ]


def kpi_cards(tables: dict) -> Table:
    """Five key figures: three on the first row, two (wider) on the second; the SLA card stands out when there
    are violations."""
    gap = 4 * mm

    def card(label, value, sub, alert, w):
        value_style = S["kpi_value"] if not alert else _style("kpi_value_w", parent=S["kpi_value"], textColor=WARN_INK)
        c = Table([[Paragraph(escape(label), S["kpi_label"])], [Paragraph(escape(value), value_style)],
                   [Paragraph(escape(sub), S["kpi_sub"])]], colWidths=[w], rowHeights=[12, 22, 20])
        c.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), WARN_FILL if alert else colors.white),
            ("BOX", (0, 0), (-1, -1), 0.5, WARN_STROKE if alert else RULE),
            ("LINEBEFORE", (0, 0), (0, -1), 3, WARN_STROKE if alert else GREEN),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 9),
            ("RIGHTPADDING", (0, 0), (-1, -1), 6),
            ("TOPPADDING", (0, 0), (-1, 0), 6),
            ("TOPPADDING", (0, 1), (-1, -1), 1),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
        ]))
        return c

    def row(items):
        w = (USABLE - gap * (len(items) - 1)) / len(items)
        cells = sum(([card(*it, w), ""] for it in items), [])[:-1]
        t = Table([cells], colWidths=sum(([w, gap] for _ in items), [])[:-1], hAlign="LEFT")
        t.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                               ("TOPPADDING", (0, 0), (-1, -1), 0), ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
                               ("VALIGN", (0, 0), (-1, -1), "TOP")]))
        return t

    items = kpi_values(tables)
    grid = Table([[row(items[:3])], [row(items[3:])]], colWidths=[USABLE], rowHeights=[None, None], hAlign="LEFT")
    grid.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                              ("TOPPADDING", (0, 0), (-1, -1), 0), ("BOTTOMPADDING", (0, 0), (-1, 0), gap),
                              ("BOTTOMPADDING", (0, 1), (-1, 1), 0)]))
    return grid


def definitions_box(width: float) -> Table:
    items = [
        ("Notified", "processed payments whose notification reached Irembo. It measures delivery, "
                     "not timeliness, and is not an SLA compliance rate."),
        ("SLA violations", "processed payments settled outside the agreed SLA window, as a share of processed."),
        ("Pending", "processed payments still waiting for their notification."),
        ("Duplicate payments", "paid at the bank while IremboPay records them as paid through another channel."),
        ("Amounts", "shown per currency; the Total row counts transactions only."),
    ]
    rows = [[Paragraph("How to read this report", S["note_h"])]]
    rows += [[Paragraph(f"<b>{escape(k)}</b>: {escape(v)}", S["note"])] for k, v in items]
    box = Table(rows, colWidths=[width])
    box.setStyle(TableStyle([("BOX", (0, 0), (-1, -1), 0.5, HAIRLINE), ("BACKGROUND", (0, 0), (-1, -1), TOTAL_FILL),
                             ("LEFTPADDING", (0, 0), (-1, -1), 8), ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                             ("TOPPADDING", (0, 0), (-1, -1), 1.5), ("BOTTOMPADDING", (0, 0), (-1, -1), 1.5),
                             ("TOPPADDING", (0, 0), (-1, 0), 6), ("BOTTOMPADDING", (0, -1), (-1, -1), 6)]))
    return box


def glance_box(tables: dict, width: float) -> Table:
    """The key figures of one week (read from its tables), in the free slot next to 'Duplicate Payment'."""
    processed, _ = _total(tables["Overall processed payments"])
    received, rate = _total(tables["Overall notifications received by Irembo"])
    items = [("Payments processed", fmt_count(processed)),
             ("Notified", fmt_pct(rate) if rate is not None else "—")]
    for label, title in (("SLA violations", "SLA Violated"), ("Pending notifications", "Pending notifications"),
                         ("Duplicate payments", "Duplicate Payment")):
        n, pct = _total(tables[title])
        items.append((label, f"{n:,}  ({fmt_pct(pct) if pct is not None else '—'})"))
    sla = _total(tables["SLA Violated"])[0]
    rows = [[Paragraph("At a glance", S["note_h"]), ""]]
    rows += [[Paragraph(escape(k), S["note"]), v] for k, v in items]
    box = Table(rows, colWidths=[width * 0.58, width * 0.42])
    style = [("BOX", (0, 0), (-1, -1), 0.5, HAIRLINE), ("BACKGROUND", (0, 0), (-1, -1), TOTAL_FILL),
             ("SPAN", (0, 0), (-1, 0)), ("FONTNAME", (1, 1), (1, -1), "Helvetica-Bold"),
             ("FONTSIZE", (1, 1), (1, -1), BODY), ("TEXTCOLOR", (1, 1), (1, -1), INK),
             ("ALIGN", (1, 1), (1, -1), "RIGHT"), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
             ("LINEBELOW", (0, 1), (-1, -2), 0.3, HAIRLINE),
             ("LEFTPADDING", (0, 0), (-1, -1), 8), ("RIGHTPADDING", (0, 0), (-1, -1), 8),
             ("TOPPADDING", (0, 0), (-1, -1), 1.6), ("BOTTOMPADDING", (0, 0), (-1, -1), 1.6),
             ("TOPPADDING", (0, 0), (-1, 0), 4), ("BOTTOMPADDING", (0, 0), (-1, 0), 2)]
    if sla:
        style.append(("TEXTCOLOR", (1, 3), (1, 3), WARN_INK))
    box.setStyle(TableStyle(style))
    return box


def merchant_block(table, width: float = USABLE) -> list:
    """Payment per merchant, under the block's tables at full width (it may run onto the next page, its header
    repeated): one row per merchant and currency, then "All" per currency; amounts are never added across
    currencies; % = share of the currency's notifications received. [] for reports made before 12/10/2026."""
    if table is None or not len(table):
        return []
    disp = pd.DataFrame({"Merchant": table["Merchant name"].astype(str), "Currency": table["Currency"].astype(str),
                         "Transactions": table["Number of Transactions"].astype(int),
                         "Amount": pd.to_numeric(table["Total Amount"], errors="coerce"),
                         "% of received": [fmt_pct(v) for v in table["%"]]})
    vs = [c for c in table.columns if str(c).startswith("VS prev")]
    if vs and table[vs[0]].notna().any():
        disp[str(vs[0]).replace("VS prev", "vs prev.").replace(" (%)", "")] = table[vs[0]].fillna("").astype(str)
    return [Spacer(1, 4 * mm), data_table(disp, width, 7.6, 4, total_row=True, title="Payment per merchant")]


def tables_block(tables: dict, with_notes: bool = False) -> list:
    """The 5 tables of one block: two columns when every value fits, otherwise one under the other at full width."""
    half = (USABLE - GAP) / 2
    if all(fits(tables[t], half) for t in TABLE_ORDER):
        cells = [summary_table(t, tables[t], half, warning=(t == "SLA Violated")) for t in TABLE_ORDER]
        cells.append(definitions_box(half) if with_notes else glance_box(tables, half))
        rows = [[cells[i], "", cells[i + 1]] for i in range(0, 6, 2)]
        g = Table(rows, colWidths=[half, GAP, half], hAlign="LEFT")
        g.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0),
                               ("RIGHTPADDING", (0, 0), (-1, -1), 0), ("TOPPADDING", (0, 0), (-1, -1), 0),
                               ("BOTTOMPADDING", (0, 0), (-1, -2), 4 * mm),
                               ("BOTTOMPADDING", (0, -1), (-1, -1), 0)]))
        return [g]
    out = []                                                  # wide values: full width, each table may span pages
    for t in TABLE_ORDER:
        out += [summary_table(t, tables[t], USABLE, warning=(t == "SLA Violated")), Spacer(1, 4 * mm)]
    if with_notes:
        out.append(KeepTogether([definitions_box(USABLE)]))
    return out


def section_heading(text: str, note: str = "", badge: str = "", alert: bool = False, width: float = USABLE) -> Table:
    left = [Paragraph(escape(text), S["h2"])] + ([Paragraph(escape(note), S["h2_note"])] if note else [])
    badge_style = _style("badge", fontName="Helvetica-Bold", fontSize=7.6, leading=9.5, alignment=2,
                         textColor=WARN_INK if alert else MUTED)
    t = Table([[left, Paragraph(escape(badge), badge_style)]], colWidths=[width * 0.66, width * 0.34])
    t.setStyle(TableStyle([("LINEBELOW", (0, 0), (-1, 0), 0.8, TEAL), ("VALIGN", (0, 0), (-1, -1), "BOTTOM"),
                           ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                           ("TOPPADDING", (0, 0), (-1, -1), 0), ("BOTTOMPADDING", (0, 0), (-1, -1), 4)]))
    return t


def sla_badge(tables: dict):
    n, p = _total(tables["SLA Violated"])
    if n == 0:
        return "No SLA violations", False
    return f"SLA violations: {n:,} ({fmt_pct(p) if p is not None else '—'} of processed)", True


# ---- page decoration (two passes so the footer can say "Page X of Y") ----------------------------------------------
def numbered_canvas(header_left: str, header_right: str, footer_left: str, pagesize=A4):
    page_w, page_h = pagesize
    usable = page_w - 2 * MARGIN

    class NumberedCanvas(rl_canvas.Canvas):
        def __init__(self, *args, **kwargs):
            super().__init__(*args, **kwargs)
            self._pages = []

        def showPage(self):
            self._pages.append(dict(self.__dict__))
            self._startPage()

        def save(self):
            total = len(self._pages)
            for state in self._pages:
                self.__dict__.update(state)
                self._decorate(total)
                super().showPage()
            super().save()

        def _decorate(self, total):
            page = self._pageNumber
            self.saveState()
            if page > 1:                                      # slim header on later pages
                y = page_h - 12 * mm
                self.setFillColor(TEAL)
                self.rect(MARGIN, y - 1.5, 2.2, 10, stroke=0, fill=1)
                self.setFont("Helvetica-Bold", 8.5)
                right_w = stringWidth(header_right, "Helvetica", 8)
                self.drawString(MARGIN + 7, y + 1, fit_text(header_left, "Helvetica-Bold", 8.5, usable - right_w - 20))
                self.setFont("Helvetica", 8)
                self.setFillColor(MUTED)
                self.drawRightString(page_w - MARGIN, y + 1, header_right)
                self.setStrokeColor(HAIRLINE)
                self.setLineWidth(0.6)
                self.line(MARGIN, y - 5, page_w - MARGIN, y - 5)
            self.setStrokeColor(HAIRLINE)
            self.setLineWidth(0.6)
            self.line(MARGIN, FOOTER_Y + 9, page_w - MARGIN, FOOTER_Y + 9)
            label = f"Page {page} of {total}"
            self.setFont("Helvetica", 7.5)
            self.setFillColor(MUTED)
            self.drawRightString(page_w - MARGIN, FOOTER_Y, label)
            room = usable - stringWidth(label, "Helvetica", 7.5) - 20
            self.drawString(MARGIN, FOOTER_Y, fit_text(footer_left, "Helvetica", 7.5, room))
            self.restoreState()

    return NumberedCanvas


# ---- the document -------------------------------------------------------------------------------------------------
def document(buf, pagesize, title: str, subject: str) -> BaseDocTemplate:
    """A4 document (portrait or landscape): page 1 starts at the top margin, later pages under the slim header."""
    w, h = pagesize
    first = Frame(MARGIN, MARGIN + 2 * mm, w - 2 * MARGIN, h - 2 * MARGIN - 2 * mm, id="first",
                  leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)
    later = Frame(MARGIN, MARGIN + 2 * mm, w - 2 * MARGIN, h - LATER_TOP - MARGIN - 2 * mm, id="later",
                  leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)
    doc = BaseDocTemplate(buf, pagesize=pagesize, leftMargin=MARGIN, rightMargin=MARGIN, topMargin=MARGIN,
                          bottomMargin=MARGIN, title=title, author="IremboPay Reconciliation", subject=subject)
    doc.addPageTemplates([PageTemplate(id="first", frames=[first]), PageTemplate(id="later", frames=[later])])
    return doc


def build_pdf(rep: dict, made: pd.Timestamp = None) -> bytes:
    """PDF bytes of a report read by engine.read_report() (tests may pass their own dict of the same shape)."""
    blocks = rep["blocks"]
    main = next((b for b in blocks if b[0] in ("Summary-Monthly", "Summary-Daily")), blocks[0])
    weeks = [b for b in blocks if b[0] == "Summary-Weekly" and b is not main]
    kind, period_name, dates = split_heading(main[1])
    title = rep.get("title") or period_name or "Report"
    if rep.get("reconciled_on"):
        reconciled = rep["reconciled_on"]
    elif made is not None:
        reconciled = f"{made:%A %d/%m/%Y} (file date)"
    else:
        reconciled = "—"
    bank = rep.get("full_name") or rep.get("bank") or "Bank"

    story = [header_band(rep, kind, period_name, dates, reconciled), NextPageTemplate("later"), Spacer(1, 5 * mm),
             kpi_cards(main[2]), Spacer(1, 6 * mm)]
    badge, alert = sla_badge(main[2])
    main_head = [section_heading(f"{kind or 'Period'} summary · {period_name}", dates, badge, alert), Spacer(1, 3 * mm)]
    merchants = rep.get("merchants") or {}
    main_tables = tables_block(main[2], with_notes=True)
    story += [KeepTogether(main_head + main_tables[:1]), *main_tables[1:], *merchant_block(merchants.get(main[1]))]
    for i, (_, heading, tables) in enumerate(weeks):
        w_kind, w_name, w_dates = split_heading(heading)
        badge, alert = sla_badge(tables)
        head = [section_heading(f"{w_kind} {w_name}".strip() if w_kind else w_name, w_dates, badge, alert),
                Spacer(1, 3 * mm)]
        if i == 0:                                            # the section title always travels with the first week
            head = [Paragraph("WEEKLY BREAKDOWN", S["kpi_label"]), Spacer(1, 2 * mm)] + head
        story += [Spacer(1, 5 * mm), KeepTogether(head + tables_block(tables)), *merchant_block(merchants.get(heading))]

    buf = io.BytesIO()
    doc = document(buf, A4, title=f"{bank} Reconciliation Report · {title}", subject=title)
    short = fit_text(bank, "Helvetica-Bold", 8.5, USABLE * 0.5)  # long names are cut so the period stays visible
    doc.build(story, canvasmaker=numbered_canvas(f"{short} · Reconciliation Report", title,
                                                 f"IremboPay · {short} reconciliation · {title}"))
    return buf.getvalue()


def report_pdf(path: Path) -> bytes:
    """Summary-only PDF of a saved report (same content as engine.report_pdf, A4 portrait management layout)."""
    path = Path(path)
    return build_pdf(engine.read_report(path), made=pd.Timestamp(path.stat().st_mtime, unit="s"))


# =====================================================================================================================
# The custom report (step 6, "Custom report"): the same look as the standard report. It draws the content built by
# engine.custom_content() as it is (sections, tables, charts, invoice lists, orientation and density chosen in the app);
# nothing is recalculated here.
# =====================================================================================================================
WRAP_COLUMNS = ("Merchant", "Comment", "Decision", "Provider", "Status")


def data_table(df: pd.DataFrame, width: float, fs: float, pad: float, total_row: bool = False,
               warning: bool = False, title: str | None = None) -> LongTable:
    """A table of the custom report: header repeated on every page, numbers on the right, long text wrapped,
    bold Total row (its amount is '—': currencies are never added together)."""
    cols = list(df.columns)
    numeric = [j for j, c in enumerate(cols) if pd.api.types.is_numeric_dtype(df[c]) or c == "% of processed"
               or (len(df) and df[c].astype(str).str.endswith("%").all())           # '99.82%' written as text
               or str(c).startswith("Amount")]                                     # amounts written as text
    cell = _style("c_cell", fontSize=fs, leading=fs + 2.2)
    th = _style("c_th", fontName="Helvetica-Bold", fontSize=fs - 0.6, leading=fs + 1.4, textColor=MUTED)
    th_r = _style("c_th_r", parent=th, alignment=2)

    def text(c, v, is_total):
        if is_total and c == "Total amount" and _missing(v):
            return NO_AMOUNT
        return engine.cell_text(c, v)

    body = []
    for rec in df.itertuples(index=False):
        is_total = total_row and str(rec[0]) in ("Total", "All")
        body.append([text(c, v, is_total) for c, v in zip(cols, rec)])
    # widths: every number and short value on one line; long text columns share what is left and wrap
    need = []
    for j, c in enumerate(cols):
        vals = [stringWidth(r[j], "Helvetica-Bold", fs) for r in body[:300]] or [0]
        head = max(stringWidth(w, "Helvetica-Bold", fs - 0.6) for w in str(c).split() or [""])
        need.append(max(head, min(max(vals), 40 * mm) if c in WRAP_COLUMNS else max(vals)) + 2 * pad + 2)
    if sum(need) > width:                                     # squeeze the wrapping columns first
        wrap = [j for j, c in enumerate(cols) if c in WRAP_COLUMNS]
        fixed = sum(w for j, w in enumerate(need) if j not in wrap)
        if wrap and fixed < width:
            for j in wrap:
                need[j] = (width - fixed) / len(wrap)
        else:
            need = [w * width / sum(need) for w in need]
    extra = width - sum(need)
    widths = [w + extra / len(need) for w in need]
    rows = [[Paragraph(escape(str(c)), th_r if j in numeric else th) for j, c in enumerate(cols)]]
    cell_bold = _style("c_cell_b", parent=cell, fontName="Helvetica-Bold")
    rows += [[Paragraph(escape(v), cell_bold if total_row and r[0] in ("Total", "All") else cell)
              if c in WRAP_COLUMNS else v for c, v in zip(cols, r)] for r in body]
    top = 0
    if title:                                       # a title row, repeated with the header on a new page
        rows.insert(0, [Paragraph(escape(title), S["t_title"])] + [""] * (len(cols) - 1))
        top = 1
    t = LongTable(rows, colWidths=widths, repeatRows=1 + top, hAlign="LEFT")
    line = WARN_STROKE if warning else TEAL
    h, b0 = top, top + 1                              # header row, first body row
    style = [("FONTNAME", (0, b0), (-1, -1), "Helvetica"), ("FONTSIZE", (0, b0), (-1, -1), fs),
             ("TEXTCOLOR", (0, b0), (-1, -1), WARN_INK if warning else INK),
             ("BACKGROUND", (0, h), (-1, h), WARN_FILL if warning else TITLE_FILL),
             ("LINEBELOW", (0, h), (-1, h), 0.8, line), ("LINEBELOW", (0, b0), (-1, -1), 0.3, HAIRLINE),
             ("BOX", (0, 0), (-1, -1), 0.5, WARN_STROKE if warning else RULE),
             ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
             ("LEFTPADDING", (0, 0), (-1, -1), pad), ("RIGHTPADDING", (0, 0), (-1, -1), pad),
             ("TOPPADDING", (0, 0), (-1, -1), pad * 0.7), ("BOTTOMPADDING", (0, 0), (-1, -1), pad * 0.7)]
    if top:
        style += [("SPAN", (0, 0), (-1, 0)), ("BACKGROUND", (0, 0), (-1, 0), TITLE_FILL),
                  ("LINEBEFORE", (0, 0), (0, 0), 3, TEAL), ("TOPPADDING", (0, 0), (-1, 0), 4),
                  ("BOTTOMPADDING", (0, 0), (-1, 0), 4), ("BACKGROUND", (0, h), (-1, h), colors.white)]
    style += [("ALIGN", (j, b0), (j, -1), "RIGHT") for j in numeric]
    for i, r in enumerate(body, b0):
        if total_row and r[0] in ("Total", "All"):
            style += [("FONTNAME", (0, i), (-1, i), "Helvetica-Bold"), ("BACKGROUND", (0, i), (-1, i), TOTAL_FILL),
                      ("LINEABOVE", (0, i), (-1, i), 0.8, line)]
    t.setStyle(TableStyle(style))
    return t


def custom_pdf(content: dict) -> bytes:
    """The custom report as a PDF, laid out like the standard report: band, key figures, then each section."""
    size = landscape(A4) if content["orientation"] == "Landscape" else A4
    usable = size[0] - 2 * MARGIN
    compact = content["density"] == "Compact"
    fs, pad, gap = (7.0, 3.0, 4 * mm) if compact else (8.0, 4.5, 6 * mm)
    bank = content["title"].removesuffix(" Reconciliation Report")
    stamp = content["stamp"].removeprefix("Reconciled on ")
    sub_h = _style("c_sub_h", parent=S["t_title"], keepWithNext=1, spaceBefore=gap / 2, spaceAfter=3)
    small = _style("c_small", fontSize=fs, leading=fs + 3, textColor=MUTED)

    story = [band(bank, content["type"], [("Reporting period", content["period"]), ("Reconciled on", stamp)],
                  usable), NextPageTemplate("later")]

    # ---- the overview: key figures in cards, then the lines per currency
    if content["kpis"]:
        n, space = len(content["kpis"]), 4 * mm
        w = (usable - space * (n - 1)) / n
        cards = []
        for label, value in content["kpis"]:
            c = Table([[Paragraph(escape(label.upper()), S["kpi_label"])], [Paragraph(escape(value), S["kpi_value"])]],
                      colWidths=[w])
            c.setStyle(TableStyle([("BOX", (0, 0), (-1, -1), 0.5, RULE), ("LINEBEFORE", (0, 0), (0, -1), 3, GREEN),
                                   ("LEFTPADDING", (0, 0), (-1, -1), 9), ("TOPPADDING", (0, 0), (-1, 0), 6),
                                   ("TOPPADDING", (0, 1), (-1, 1), 1), ("BOTTOMPADDING", (0, 0), (-1, 0), 0),
                                   ("BOTTOMPADDING", (0, 1), (-1, 1), 7 if not compact else 5)]))
            cards.append(c)
        row = Table([sum(([c, ""] for c in cards), [])[:-1]],
                    colWidths=sum(([w, space] for _ in cards), [])[:-1], hAlign="LEFT")
        row.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                                 ("TOPPADDING", (0, 0), (-1, -1), 0), ("BOTTOMPADDING", (0, 0), (-1, -1), 0)]))
        story += [Spacer(1, 5 * mm), row]
    lines = content["currency_lines"] + ([content["exceptions_line"]] if content["exceptions_line"] else [])
    if lines:
        box = Table([[Paragraph(escape(x), _style("c_line", fontSize=fs + 0.3, leading=fs + 4))] for x in lines],
                    colWidths=[usable], hAlign="LEFT")
        box.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), TOTAL_FILL), ("BOX", (0, 0), (-1, -1), 0.5, HAIRLINE),
                                 ("LEFTPADDING", (0, 0), (-1, -1), 9), ("TOPPADDING", (0, 0), (-1, -1), 2),
                                 ("BOTTOMPADDING", (0, 0), (-1, -1), 2), ("TOPPADDING", (0, 0), (-1, 0), 6),
                                 ("BOTTOMPADDING", (0, -1), (-1, -1), 6)]))
        story += [Spacer(1, 3 * mm), box]

    def invoice_list(sec, named=False):
        """The invoices of a section, under its table or its chart card (named: the section is said)."""
        if sec["list"] is None:
            return []
        title = (f"{sec['title']} · " if named else "") + f"{sec['list_title']} ({len(sec['list']):,})"
        return [Paragraph(escape(title), sub_h),
                data_table(sec["list"], usable, fs, pad) if len(sec["list"]) else Paragraph("None.", small)]

    def chart_card(sec, width):
        """One chart in its own card: title, what is counted, the chart, the values (for print)."""
        inner = width - 18
        parts = [Paragraph(escape(sec["title"]), S["t_title"]), Paragraph(escape(engine.series_label(sec)), small),
                 Spacer(1, 3), engine._pdf_chart(sec["as"], sec["series"], inner, (42 if compact else 55) * mm),
                 Paragraph(escape(" · ".join(f"{name} {v:,}" for name, v in sec["series"])), small)]
        card = Table([[parts]], colWidths=[width], hAlign="LEFT")
        card.setStyle(TableStyle([("BOX", (0, 0), (-1, -1), 0.5, RULE),
                                  ("LINEBEFORE", (0, 0), (0, -1), 3, WARN_STROKE if sec["key"] == "sla" else TEAL),
                                  ("VALIGN", (0, 0), (-1, -1), "TOP"),
                                  ("LEFTPADDING", (0, 0), (-1, -1), 9), ("RIGHTPADDING", (0, 0), (-1, -1), 9),
                                  ("TOPPADDING", (0, 0), (-1, -1), 7), ("BOTTOMPADDING", (0, 0), (-1, -1), 7)]))
        return card

    # ---- the sections: charts side by side (two per row, each in a card), tables and lists full width
    space = 5 * mm
    for kind, item in engine.layout_rows(content["sections"]):
        if kind == "section":
            sec = item
            head = section_heading(sec["title"], width=usable)
            head.keepWithNext = 1                             # a heading never ends a page alone
            story += [Spacer(1, gap), head, Spacer(1, 3 * mm)]
            story[-1].keepWithNext = 1
            if sec["table"] is not None:
                story.append(data_table(sec["table"], min(usable, 160 * mm), fs, pad, total_row=True,
                                        warning=sec["key"] == "sla"))
            story += invoice_list(sec)
            continue
        half = (usable - space) / 2
        cards = [chart_card(sec, half if len(item) == 2 else usable) for sec in item]
        row = Table([cards], colWidths=[half + space, half] if len(cards) == 2 else [usable], hAlign="LEFT")
        row.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0),
                                 ("RIGHTPADDING", (0, 0), (0, -1), space if len(cards) == 2 else 0),
                                 ("RIGHTPADDING", (1, 0), (-1, -1), 0),
                                 ("TOPPADDING", (0, 0), (-1, -1), 0), ("BOTTOMPADDING", (0, 0), (-1, -1), 0)]))
        story += [Spacer(1, gap), row]
        for sec in item:
            story += invoice_list(sec, named=True)

    buf = io.BytesIO()
    doc = document(buf, size, title=f"{content['title']} {content['period']}", subject=content["type"])
    short = fit_text(bank, "Helvetica-Bold", 8.5, (size[0] - 2 * MARGIN) * 0.5)
    doc.build(story, canvasmaker=numbered_canvas(f"{short} · Reconciliation Report", content["period"],
                                                 f"IremboPay · {short} · {content['type']} · {content['period']}",
                                                 pagesize=size))
    return buf.getvalue()


# =====================================================================================================================
# The Reports Dashboard (side panel "Dashboard") as a PDF: the view on screen, drawn with the same look. The app
# passes the figures it shows (services/dashboard.py read them from the saved monthly reports); nothing is computed.
# =====================================================================================================================
def _cards(items: list, per_row: int, width: float) -> Table:
    """Key figure cards [(label, value, sub, alert)], `per_row` per row."""
    gap = 4 * mm
    w = (width - gap * (per_row - 1)) / per_row
    cells = []
    multi = any(" · " in value for _, value, _, _ in items)       # 'RWF 196.0M · USD 634.9K': one line each
    for label, value, sub, alert in items:
        value_style = S["kpi_value"] if not alert else _style("kpi_value_w", parent=S["kpi_value"], textColor=WARN_INK)
        if " · " in value:
            value_style = _style("kpi_value_m", parent=value_style, fontSize=12.5, leading=14.5)
            text = "<br/>".join(escape(v) for v in value.split(" · "))
        else:
            text = escape(value)
        c = Table([[Paragraph(escape(label.upper()), S["kpi_label"])], [Paragraph(text, value_style)],
                   [Paragraph(escape(sub), S["kpi_sub"])]], colWidths=[w], rowHeights=[20, 32 if multi else 22, 22])
        c.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), WARN_FILL if alert else colors.white),
            ("BOX", (0, 0), (-1, -1), 0.5, WARN_STROKE if alert else RULE),
            ("LINEBEFORE", (0, 0), (0, -1), 3, WARN_STROKE if alert else GREEN),
            ("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 8),
            ("RIGHTPADDING", (0, 0), (-1, -1), 5), ("TOPPADDING", (0, 0), (-1, 0), 6),
            ("TOPPADDING", (0, 1), (-1, -1), 1), ("BOTTOMPADDING", (0, 0), (-1, -1), 0)]))
        cells.append(c)
    rows = []
    for i in range(0, len(cells), per_row):
        part = cells[i:i + per_row] + [""] * (per_row - len(cells[i:i + per_row]))
        rows.append(sum(([c, ""] for c in part), [])[:-1])
        rows.append([""] * (2 * per_row - 1))
    grid = Table(rows[:-1], colWidths=sum(([w, gap] for _ in range(per_row)), [])[:-1], hAlign="LEFT",
                 rowHeights=[None if i % 2 == 0 else gap for i in range(len(rows) - 1)])
    grid.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                              ("TOPPADDING", (0, 0), (-1, -1), 0), ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
                              ("VALIGN", (0, 0), (-1, -1), "TOP")]))
    return grid


def _axis_label(v) -> str:
    v = float(v)
    for size, unit in ((1e9, "B"), (1e6, "M"), (1e3, "K")):
        if abs(v) >= size:
            return f"{v / size:,.1f}".rstrip("0").rstrip(".") + unit
    return f"{v:,.0f}"


def _chart_card(title: str, drawing, legend: list, width: float) -> Table:
    """A chart in a card: title, legend (coloured names), the chart."""
    from reportlab.graphics.shapes import Drawing, Rect, String
    parts = [Paragraph(escape(title), S["t_title"])]
    if legend:
        lg = Drawing(width - 16, 12)
        x = 0
        for name, colour in legend:
            lg.add(Rect(x, 2, 7, 7, fillColor=colors.HexColor(colour), strokeColor=None))
            lg.add(String(x + 10, 2, name, fontName="Helvetica", fontSize=7, fillColor=MUTED))
            x += 18 + stringWidth(name, "Helvetica", 7)
        parts.append(lg)
    parts.append(drawing)
    card = Table([[parts]], colWidths=[width], hAlign="LEFT")
    card.setStyle(TableStyle([("BOX", (0, 0), (-1, -1), 0.5, RULE), ("VALIGN", (0, 0), (-1, -1), "TOP"),
                              ("LEFTPADDING", (0, 0), (-1, -1), 8), ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                              ("TOPPADDING", (0, 0), (-1, -1), 7), ("BOTTOMPADDING", (0, 0), (-1, -1), 7)]))
    return card


def _series_chart(kind: str, labels: list, series: list, width: float, height: float):
    """Line or grouped bar chart, one series per bank [(name, colour, values)]."""
    from reportlab.graphics.charts.barcharts import VerticalBarChart
    from reportlab.graphics.charts.linecharts import HorizontalLineChart
    from reportlab.graphics.shapes import Drawing
    from reportlab.graphics.widgets.markers import makeMarker
    d = Drawing(width, height)
    ch = VerticalBarChart() if kind == "bar" else HorizontalLineChart()
    ch.x, ch.y, ch.width, ch.height = 34, 18, width - 40, height - 26
    ch.data = [list(v) for _, _, v in series] or [[0]]
    top = max([max(v) for _, _, v in series if v] + [1])
    step = engine._nice_step(top)
    ch.valueAxis.valueMin, ch.valueAxis.valueStep = 0, step
    ch.valueAxis.valueMax = step * (int(top / step) + 1)
    ch.valueAxis.visibleGrid, ch.valueAxis.gridStrokeColor, ch.valueAxis.gridStrokeWidth = True, HAIRLINE, 0.5
    ch.valueAxis.strokeColor, ch.valueAxis.visibleTicks = HAIRLINE, False
    ch.valueAxis.labels.fontName = ch.categoryAxis.labels.fontName = "Helvetica"
    ch.valueAxis.labels.fontSize = ch.categoryAxis.labels.fontSize = 6.5
    ch.valueAxis.labels.fillColor = ch.categoryAxis.labels.fillColor = MUTED
    ch.valueAxis.labelTextFormat = _axis_label
    ch.categoryAxis.categoryNames = labels
    ch.categoryAxis.strokeColor, ch.categoryAxis.visibleTicks = HAIRLINE, False
    for i, (_, colour, _) in enumerate(series):
        c = colors.HexColor(colour)
        if kind == "bar":
            ch.bars[i].fillColor, ch.bars[i].strokeColor = c, None
        else:
            ch.lines[i].strokeColor, ch.lines[i].strokeWidth = c, 1.8
            ch.lines[i].symbol = makeMarker("FilledCircle", size=4, fillColor=c, strokeColor=colors.white)
    if kind == "bar":
        ch.groupSpacing, ch.barSpacing = 8, 1
    d.add(ch)
    return d


def _bank_bars(series: list, width: float, fmt) -> Table:
    """One month: a horizontal bar per bank with its value (a small table of bars: easy to read in print)."""
    from reportlab.graphics.shapes import Drawing, Rect
    top = max([v for _, _, v in series] + [1])
    bar_w = width - 150
    rows = []
    for name, colour, v in series:
        d = Drawing(bar_w, 12)
        d.add(Rect(0, 1, max(1.5, bar_w * v / top), 10, fillColor=colors.HexColor(colour), strokeColor=None))
        rows.append([Paragraph(escape(name), _style("bb", fontSize=7.5)), d, fmt(v)])
    t = Table(rows, colWidths=[80, bar_w, 54], hAlign="LEFT")
    t.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("FONTNAME", (2, 0), (2, -1), "Helvetica-Bold"),
                           ("FONTSIZE", (2, 0), (2, -1), 7.5), ("ALIGN", (2, 0), (2, -1), "RIGHT"),
                           ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 2),
                           ("TOPPADDING", (0, 0), (-1, -1), 5), ("BOTTOMPADDING", (0, 0), (-1, -1), 5)]))
    return t


def _note_list(title: str, items: list, width: float) -> Table:
    dots = {"good": colors.HexColor("#2E7D4F"), "warn": WARN_STROKE, "bad": colors.HexColor("#C2413D"),
            "info": colors.HexColor("#2B8A99")}
    rows = [[Paragraph(escape(title), S["note_h"]), ""]]
    for kind, text in items:
        rows.append(["●", Paragraph(escape(text), _style("hl", fontSize=7.8, leading=10))])
    t = Table(rows, colWidths=[12, width - 12], hAlign="LEFT")
    style = [("SPAN", (0, 0), (-1, 0)), ("BOX", (0, 0), (-1, -1), 0.5, HAIRLINE),
             ("BACKGROUND", (0, 0), (-1, -1), TOTAL_FILL), ("VALIGN", (0, 0), (-1, -1), "TOP"),
             ("LEFTPADDING", (0, 0), (-1, -1), 7), ("RIGHTPADDING", (0, 0), (-1, -1), 7),
             ("TOPPADDING", (0, 0), (-1, -1), 2.5), ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5),
             ("TOPPADDING", (0, 0), (-1, 0), 6), ("FONTSIZE", (0, 1), (0, -1), 7)]
    style += [("TEXTCOLOR", (0, i), (0, i), dots.get(kind, MUTED)) for i, (kind, _) in enumerate(items, 1)]
    t.setStyle(TableStyle(style))
    return t


def dashboard_pdf(data: dict) -> bytes:
    """The Reports Dashboard as an A4 PDF. data = what the dashboard shows: scope, period, compared, currencies (one,
    or every currency when "All" is chosen: each is shown on its own, never added), one_month,
    kpis [(label, value, sub, alert)], banks [(name, colour)], months [labels], processed {bank: [value per month]},
    amount {currency: {bank: [value per month]}}, status [(label, count, share)], sla [(bank, count, share)],
    report_columns + reports [rows], highlights [(kind, text)]."""
    curs = data["currencies"]
    half = (USABLE - GAP) / 2
    story = [band(data["scope"], f"{data['period']} · amounts in {' and '.join(curs)}"
                  + (" (each on its own)" if len(curs) > 1 else ""),
                  [("Period", data["period"]), ("Compared with", data["compared"] or "—"),
                   ("Generated on", f"{pd.Timestamp.now():%d %b %Y · %H:%M}")], title="Reports Dashboard"),
             NextPageTemplate("later"), Spacer(1, 5 * mm), _cards(data["kpis"], min(4, len(data["kpis"])), USABLE),
             Spacer(1, 6 * mm)]

    # ---- charts: over time (a range of months) or one bar per bank (one month); one amount chart per currency
    legend = data["banks"]
    if data["one_month"]:
        cards = [_chart_card(f"Payments processed by bank · {data['period']}",
                             _bank_bars([(n, c, data["processed"][n][0]) for n, c in legend], half - 16, fmt_count),
                             [], half)]
        cards += [_chart_card(f"Amount collected by bank ({cur})",
                              _bank_bars([(n, c, data["amount"][cur][n][0]) for n, c in legend
                                          if n in data["amount"][cur]], half - 16, _axis_label), [], half)
                  for cur in curs]
    else:
        cards = [_chart_card("Payments processed over time",
                             _series_chart("line", data["months"], [(n, c, data["processed"][n]) for n, c in legend],
                                           half - 16, 52 * mm), legend, half)]
        cards += [_chart_card(f"Amount collected over time ({cur})",
                              _series_chart("bar", data["months"], [(n, c, data["amount"][cur][n]) for n, c in legend
                                                                    if n in data["amount"][cur]], half - 16, 52 * mm),
                              [(n, c) for n, c in legend if n in data["amount"][cur]], half) for cur in curs]
    rows = [[cards[i], "", cards[i + 1] if i + 1 < len(cards) else ""] for i in range(0, len(cards), 2)]
    pair = Table(rows, colWidths=[half, GAP, half], hAlign="LEFT")
    pair.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0),
                              ("RIGHTPADDING", (0, 0), (-1, -1), 0), ("TOPPADDING", (0, 0), (-1, -1), 0),
                              ("BOTTOMPADDING", (0, 0), (-1, -2), 4 * mm), ("BOTTOMPADDING", (0, -1), (-1, -1), 0)]))
    story += [KeepTogether([section_heading("Activity", data["period"]), Spacer(1, 3 * mm), pair]), Spacer(1, 6 * mm)]

    # ---- reconciliation status and SLA violations
    status = pd.DataFrame([(n, v, s) for n, v, s in data["status"]],
                          columns=["Status", "Transactions", "% of processed"])
    sla = pd.DataFrame([(n, v, s) for n, v, s in data["sla"]], columns=["Bank", "SLA violations", "% of processed"])
    two = Table([[[Paragraph("Reconciliation status (notifications)", S["t_title"]), Spacer(1, 3),
                   data_table(status, half, 7.6, 4)], "",
                  [Paragraph("SLA violations by bank", S["t_title"]), Spacer(1, 3),
                   data_table(sla, half, 7.6, 4, warning=bool(sla["SLA violations"].sum()))]]],
                colWidths=[half, GAP, half], hAlign="LEFT")
    two.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0),
                             ("RIGHTPADDING", (0, 0), (-1, -1), 0), ("TOPPADDING", (0, 0), (-1, -1), 0),
                             ("BOTTOMPADDING", (0, 0), (-1, -1), 0)]))
    story += [KeepTogether([section_heading("Status", "Receipt of notifications and SLA"), Spacer(1, 3 * mm), two]),
              Spacer(1, 6 * mm)]

    # ---- the reports used, and the highlights
    reports = pd.DataFrame(data["reports"], columns=data["report_columns"])
    head = section_heading("Reports in this view", f"{len(reports)} monthly report{'s' if len(reports) != 1 else ''}"
                                                   " (first saved version)")
    head.keepWithNext = 1
    story += [head, Spacer(1, 3 * mm), data_table(reports, USABLE, 7.6, 4), Spacer(1, 6 * mm)]
    if data.get("merchants"):                         # the top merchants of the view (notifications received)
        merch = pd.DataFrame(data["merchants"], columns=["Merchant", "Currency", "Transactions", "Amount",
                                                         "% of received"])
        story += [data_table(merch, USABLE, 7.6, 4, title="Top merchants (notifications received)")]
        if data.get("merchants_missing"):
            story.append(Paragraph(escape(f"{data['merchants_missing']} report(s) made before the merchant table "
                                          "existed are not included."), _style("mm", fontSize=7, textColor=MUTED)))
        story.append(Spacer(1, 6 * mm))
    story += [KeepTogether([_note_list("Highlights", data["highlights"], USABLE)])]

    buf = io.BytesIO()
    doc = document(buf, A4, title=f"Reports Dashboard · {data['scope']} · {data['period']}", subject="Reports Dashboard")
    short = fit_text(data["scope"], "Helvetica-Bold", 8.5, USABLE * 0.5)
    doc.build(story, canvasmaker=numbered_canvas(f"{short} · Reports Dashboard", data["period"],
                                                 f"IremboPay · Reports Dashboard · {short} · {data['period']}"))
    return buf.getvalue()
