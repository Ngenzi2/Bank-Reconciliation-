"""
THE DASHBOARD'S MAIN VIEW (22/10/2026): four key cards (Amount, Rate, Transactions, SLA), then Payments over time,
Amount overview, Exceptions by type and Top merchants; each chart interactive on the screen (Altair) and drawn as an
SVG for its PDF, with its table.
Every number, curve, slice and line is computed from the facts of the saved reports (services/dashboard.py);
nothing is typed by hand. IremboPay colours: white / pale teal surfaces, dark teal text and figures, brand green for
results that are good, red / amber for problems only.

    kpi_cards / lines / dual_bars / merchant_bars / table_html / card_page / alt_lines / alt_dual_bars
"""
from __future__ import annotations

import base64
import math
from html import escape

import pandas as pd

TEAL, ACC, PALE, GREEN, AMBER, AMBER_L, BAD, MUT = ("#123239", "#2B8A99", "#EAF4F5", "#99E378", "#D4A72C", "#E9C46A",
                                                     "#C2413D", "#5E7378")


SVG_STYLE = (f"text{{font-family:Nunito,'Segoe UI',Helvetica,Arial,sans-serif}} .dv-ax{{font-size:10.5px;fill:{MUT}}} "
             f".dv-gl{{stroke:#E6EEF0;stroke-dasharray:3 4}} .dv-t1{{font-size:10.5px;fill:#CFE3E6}} "
             f".dv-t2{{font-size:15px;fill:#fff;font-weight:800}} .dv-dc{{font-size:28px;font-weight:800;fill:{TEAL}}} "
             f".dv-ds{{font-size:11.5px;fill:{MUT}}}")


def _img(svg: str, style: str = "width:100%;display:block") -> str:
    """An SVG drawing as an <img> (Streamlit removes inline <svg> from its HTML); its text styles go inside it."""
    svg = svg.replace("<svg ", '<svg xmlns="http://www.w3.org/2000/svg" ', 1)
    svg = svg.replace(">", f"><style>{SVG_STYLE}</style>", 1)
    return f'<img alt="" style="{style}" src="data:image/svg+xml;base64,{base64.b64encode(svg.encode()).decode()}"/>'


def short(v: float) -> str:
    for size, unit in ((1e9, "B"), (1e6, "M"), (1e3, "K")):
        if abs(v) >= size:
            return f"{v / size:,.1f}{unit}"
    return f"{v:,.0f}"


# ---- drawings ------------------------------------------------------------------------------------------------------
def _smooth(pts, floor=None):
    """A smooth curve through the points that never goes above or below its two neighbours (monotone cubic,
    Fritsch-Carlson): between two weeks it shows no value higher or lower than theirs, so no false peak."""
    n = len(pts)
    h = [pts[i + 1][0] - pts[i][0] for i in range(n - 1)]
    dl = [(pts[i + 1][1] - pts[i][1]) / h[i] for i in range(n - 1)]
    m = [dl[0]] + [0.0 if dl[i - 1] * dl[i] <= 0 else (dl[i - 1] + dl[i]) / 2 for i in range(1, n - 1)] + [dl[-1]]
    for i in range(n - 1):
        if dl[i] == 0:
            m[i] = m[i + 1] = 0.0
            continue
        a, b = m[i] / dl[i], m[i + 1] / dl[i]
        if a * a + b * b > 9:
            t = 3 / math.sqrt(a * a + b * b)
            m[i], m[i + 1] = t * a * dl[i], t * b * dl[i]
    path = f"M{pts[0][0]:.1f},{pts[0][1]:.1f}"
    for i in range(n - 1):
        (x1, y1), (x2, y2) = pts[i], pts[i + 1]
        path += (f" C{x1 + h[i] / 3:.1f},{y1 + m[i] * h[i] / 3:.1f} {x2 - h[i] / 3:.1f},{y2 - m[i + 1] * h[i] / 3:.1f} "
                 f"{x2:.1f},{y2:.1f}")
    return path


def _nice(vmax: float):
    """(axis top, step) with about 4 steps: 1, 2, 2.5, 5 x 10^n."""
    if vmax <= 0:
        return 1, 1
    raw = vmax / 4
    mag = 10 ** math.floor(math.log10(raw))
    step = next(m * mag for m in (1, 2, 2.5, 5, 10) if m * mag >= raw)
    return step * math.ceil(vmax / step), step


def _axis(L, R, T, B, top, step):
    out, v = "", 0.0
    while v <= top + 1e-9:
        y = B - v / top * (B - T)
        out += (f'<line x1="{L}" x2="{R}" y1="{y:.1f}" y2="{y:.1f}" class="dv-gl"/>'
                f'<text x="{L - 8}" y="{y + 4:.1f}" class="dv-ax" text-anchor="end">{short(v) if v else "0"}</text>')
        v += step
    return out


STYLE = f'''<style>
.dv{{font-family:"Nunito","Segoe UI",system-ui,sans-serif;color:{TEAL}}}
.dv-hero{{margin:0 0 16px}} .dv-hl{{font-size:14px;font-weight:800;color:{MUT}}}
.dv-hv{{font-size:46px;font-weight:800;letter-spacing:-1px;color:{TEAL};display:flex;align-items:center;gap:14px;line-height:1.15}}
.dv-hp{{font-size:14.5px;color:{MUT};margin-top:4px}} .dv-hp b{{color:{TEAL}}}
.dv-grid{{display:grid;grid-template-columns:repeat(4,minmax(0,1fr)) minmax(0,1.25fr);gap:16px}}
.dv-card{{background:#fff;border:1px solid #E1EEF0;border-radius:18px;padding:18px 18px;box-shadow:0 4px 16px rgba(18,50,57,.06);min-width:0}}
.dv-s2{{grid-column:span 2}} .dv-s3{{grid-column:span 3}}
.dv-k .dv-ic svg{{width:38px;height:38px}} .dv-kl{{font-size:13.5px;font-weight:800;color:{MUT};margin-top:10px}}
.dv-kv{{font-size:26px;font-weight:800;margin:4px 0 10px;color:{TEAL}}}
.dv-kf{{display:flex;justify-content:space-between;align-items:center;gap:6px}}
.dv-go{{width:28px;height:28px;border-radius:50%;background:{PALE};display:flex;align-items:center;justify-content:center;
  font-size:19px;color:{TEAL}!important;text-decoration:none!important;flex:none}}
.dv-ks{{font-size:12px;color:{MUT};margin-top:8px}}
.dv-pill{{font-size:12px;font-weight:800;padding:3px 9px;border-radius:999px;white-space:nowrap}}
.dv-pill.g{{background:{GREEN};color:{TEAL}}} .dv-pill.r{{background:#FBE5E4;color:{BAD}}}
.dv-pill.a{{background:#FBF0DC;color:#8A5B00}} .dv-pill.n{{background:#EEF2F3;color:{MUT}}}
.dv-hv .dv-pill{{font-size:17px;padding:5px 12px}}
.dv-ct{{display:flex;justify-content:space-between;align-items:center;margin-bottom:4px}}
.dv-ct h3{{margin:0;font-size:17px;color:{TEAL};font-weight:800;padding:0}} .dv-ct small{{color:{MUT};font-size:12px;font-weight:700}}
.dv-sub{{font-size:12px;color:{MUT};margin-bottom:6px}}
.dv-ch{{width:100%;display:block}} .dv-ax{{font-size:10.5px;fill:{MUT}}} .dv-gl{{stroke:#E6EEF0;stroke-dasharray:3 4}}
.dv-t1{{font-size:10.5px;fill:#CFE3E6}} .dv-t2{{font-size:15px;fill:#fff;font-weight:800}}
.dv-row{{display:flex;align-items:center;gap:10px;padding:10px 0;border-top:1px solid #EEF3F4}} .dv-row:first-of-type{{border-top:0}}
.dv-ri{{width:32px;height:32px;border-radius:9px;background:{PALE};display:flex;align-items:center;justify-content:center;
  color:{TEAL};font-weight:800;flex:none}}
.dv-rn{{font-weight:800;font-size:13.5px}} .dv-rs{{font-size:11.5px;color:{MUT}}} .dv-row .dv-pill{{margin-left:auto}}
.dv-dwrap{{display:flex;align-items:center;justify-content:space-between;gap:10px}}
.dv-dc{{font-size:28px;font-weight:800;fill:{TEAL}}} .dv-ds{{font-size:11.5px;fill:{MUT}}}
.dv-lg{{display:flex;gap:9px;align-items:flex-start;margin:9px 0;font-size:13.5px}}
.dv-dot{{width:11px;height:11px;border-radius:50%;margin-top:4px;flex:none}} .dv-mut{{color:{MUT};font-size:12px}}
.dv-wk{{max-width:280px;margin-top:8px}} .dv-empty{{font-size:13px;color:{MUT};padding:16px 0}}
@media (max-width:1100px){{.dv-grid{{grid-template-columns:repeat(2,minmax(0,1fr))}} .dv-s2,.dv-s3{{grid-column:span 2}}}}
</style>'''


# ---- the PDF of the view: the same blocks as on the screen (the same HTML), printed by Chrome --------------------
def _chrome() -> str | None:
    """Google Chrome (or Chromium / Edge) on this computer, used to print the view as a PDF; None if there is none."""
    import shutil
    from pathlib import Path
    for p in ("/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
              "/Applications/Chromium.app/Contents/MacOS/Chromium",
              "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
              r"C:\Program Files\Google\Chrome\Application\chrome.exe",
              r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"):
        if Path(p).exists():
            return p
    return next((w for w in map(shutil.which, ("google-chrome", "chromium", "chromium-browser", "chrome")) if w), None)


def print_page(grid: str, scope: str, period: str, compared: str, currencies: list, made: str) -> str:
    """One A4 landscape page: a band (bank, period, compared with, currencies, when it was made, all from the view
    on the screen), then the blocks of the view."""
    cur_txt = (f"amounts in {' and '.join(currencies)} (each on its own)" if len(currencies) > 1
               else f"amounts in {currencies[0]}" if currencies else "")
    band = (f'<div class="pb"><div class="pl"><div class="pk">{escape(scope.upper())}</div>'
            f'<div class="pt">Reports Dashboard</div><div class="ps">{escape(period)}'
            f'{" · " + escape(cur_txt) if cur_txt else ""}</div></div><div class="pr">'
            f'<div class="pm">PERIOD</div><div class="pv">{escape(period)}</div>'
            f'<div class="pm">COMPARED WITH</div><div class="pv">{escape(compared or "—")}</div>'
            f'<div class="pm">GENERATED ON</div><div class="pv">{escape(made)}</div></div></div>')
    return f'''<!doctype html><html><head><meta charset="utf-8">
<link href="https://fonts.googleapis.com/css2?family=Nunito:wght@400;600;700;800&display=swap" rel="stylesheet">
<style>
@page {{size: A4 landscape; margin: 8mm;}}
html, body {{margin: 0; background: #fff; -webkit-print-color-adjust: exact; print-color-adjust: exact;}}
body {{width: 1680px; zoom: .62; font-family: Nunito, "Segoe UI", system-ui, sans-serif;}}
.dv-grid {{grid-template-columns: repeat(4, minmax(0, 1fr)) minmax(0, 1.25fr) !important; gap: 14px !important;}}
.dv-s2 {{grid-column: span 2 !important;}}
.dv-card {{break-inside: avoid; box-shadow: none !important;}}
.dv-go {{display: none !important;}}
.dv-s2 > img {{height: 235px !important;}}
.dv-card > img {{max-height: 190px;}}
.pb {{background: {TEAL}; border-left: 9px solid {GREEN}; border-radius: 12px; display: flex; margin-bottom: 16px;
      font-family: Helvetica, Arial, sans-serif; color: #fff;}}
.pl {{flex: 1; padding: 26px 34px; display: flex; flex-direction: column; justify-content: center;}}
.pk {{color: {GREEN}; font-weight: 700; font-size: 19px; letter-spacing: .02em;}}
.pt {{font-size: 40px; font-weight: 700; margin: 2px 0 6px;}} .ps {{font-size: 19px; color: #D6E4E6;}}
.pr {{width: 400px; border-left: 1px solid rgba(255,255,255,.14); padding: 20px 32px;}}
.pm {{font-size: 16px; color: #D6E4E6; margin-top: 8px;}} .pm:first-child {{margin-top: 0;}}
.pv {{font-size: 19px; font-weight: 700;}}
</style></head><body>
{band}{grid}
</body></html>'''


def pdf(page_html: str, timeout: int = 60) -> bytes:
    """The PDF of a page printed by headless Chrome. RuntimeError when Chrome is missing or fails."""
    import subprocess
    import tempfile
    from pathlib import Path
    chrome = _chrome()
    if chrome is None:
        raise RuntimeError("Google Chrome was not found on this computer")
    with tempfile.TemporaryDirectory() as tmp:
        src, out = Path(tmp) / "dashboard.html", Path(tmp) / "dashboard.pdf"
        src.write_text(page_html, encoding="utf-8")
        run = subprocess.run([chrome, "--headless=new", "--disable-gpu", "--no-pdf-header-footer",
                              "--run-all-compositor-stages-before-draw", "--virtual-time-budget=4000",
                              f"--user-data-dir={tmp}/profile", f"--print-to-pdf={out}", src.as_uri()],
                             capture_output=True, timeout=timeout)
        if not out.exists() or out.stat().st_size == 0:
            raise RuntimeError(f"Chrome could not print the PDF ({run.stderr.decode(errors='ignore')[-300:]})")
        return out.read_bytes()


# =====================================================================================================================
# THE DASHBOARD (22/10/2026): four key cards, then two rows of two charts, each chart with a Table view and a PDF
# =====================================================================================================================
USD_C, EXC = ("#2E7D4F", {"SLA violated": BAD, "Paid with other channel": AMBER, "Pending notification": "#E58A2C"})


def _arrow(now, before, good_up=True, points=False, fmt=lambda v: f"{v:,}"):
    """'↗ 2.1% from 4,290' (green when good, red when bad), or None when there is nothing to compare with."""
    if now is None or before is None:
        return None
    if points:
        diff = (now - before) * 100
        if abs(diff) < 0.005:
            return f'<span class="dk-n">= no change from {fmt(before)}</span>'
        txt, up = f"{abs(diff):.2f} pts", diff > 0
    else:
        if not before:
            return None
        x = (now - before) / before * 100
        if abs(x) < 0.05:
            return f'<span class="dk-n">= no change from {fmt(before)}</span>'
        txt, up = f"{abs(x):.1f}%", x > 0
    good = up == good_up
    return f'<span class="{"dk-g" if good else "dk-r"}">{"↗" if up else "↘"} {txt}</span> from {fmt(before)}'


def kpi_cards(now: dict, before: dict | None, fu: dict | None = None) -> str:
    """AMOUNT (received; RWF and USD each in full, never added) · RATE (received notification ÷ processed
    transaction) · TRANSACTIONS (processed) · SLA (% violated). With a comparison: the change; else what it is."""
    rec_amt = lambda t: {c: a.get("received", 0.0) for c, a in (t or {}).get("amounts", {}).items()}
    amt, amt_b = rec_amt(now), (rec_amt(before) if before else None)
    curs = [c for c in ("RWF", "USD") if c in amt] + [c for c in amt if c not in ("RWF", "USD")]
    rate = lambda t: t["received"] / t["processed"] if t and t["processed"] else None
    sla = lambda t: t["sla"] / t["processed"] if t and t["processed"] else None
    pend = lambda t: t.get("pending", 0) / t["processed"] if t and t["processed"] else None
    # a share above 0 that would round to 0 (e.g. 1 of 50,000) is shown as "<0.01%", never as "0.00%"
    pct = lambda v, n=2: ("N/A" if v is None else f"<{10 ** -n:.{n}f}%" if 0 < v * 100 < 0.5 * 10 ** -n
                          else f"{v * 100:.{n}f}%")
    amount_big = "<br>".join(f"{c} {short(amt[c])}" for c in curs) or "–"
    amount_sub = (" · ".join(f"{c} " + (_arrow(amt[c], amt_b.get(c), fmt=lambda v: short(v)) or "no earlier figure")
                             for c in curs) if amt_b is not None else "Total received (notifications received)")
    cards = [
        ("AMOUNT", "money", amount_big, amount_sub),
        # 27/10/2026: TRANSACTIONS before RATE
        ("TRANSACTIONS", "tx", f"{now['processed']:,}",
         _arrow(now["processed"], before["processed"]) if before else "Processed transactions in this period"),
        ("RATE", "rate", pct(rate(now)),
         _arrow(rate(now), rate(before), points=True, fmt=pct) if before else "Received notification ÷ processed transaction"),
        ("SLA", "sla", pct(sla(now), 1),
         _arrow(sla(now), sla(before), good_up=False, points=True, fmt=lambda v: pct(v, 1)) if before
         else f"Violated: {now['sla']:,} of {now['processed']:,} processed transactions"),
        # 27/10/2026: pending notifications ÷ processed (debited at the bank, never notified to IremboPay)
        ("PENDING", "pending", pct(pend(now)),             # 27/10/2026: 2 decimals (4 of 10,506 = 0.04%, not 0.0%)
         _arrow(pend(now), pend(before), good_up=False, points=True, fmt=pct) if before
         else f"Not notified: {now.get('pending', 0):,} of {now['processed']:,} processed transactions"
              + (" · " + " · ".join(f"{c} {short(v)}" for c in curs
                                    if (v := (now.get('amounts', {}).get(c, {}) or {}).get('pending', 0.0)))
                 if any((now.get('amounts', {}).get(c, {}) or {}).get('pending') for c in curs) else "")),
    ]
    # the follow-up after the report (services/followup.py): the report's figures are never changed, the follow-up
    # is shown next to them
    if fu is not None:
        if fu.get("followed"):
            s = fu["summary"]
            # 30/10/2026: the money still to follow up is shown on the FOLLOW-UP card only, not on AMOUNT
            if s["received_later"] and now["processed"]:
                after = (now["received"] + s["received_later"]) / now["processed"]
                cards[2] = cards[2][:3] + (cards[2][3] + f'<div class="dk-f"><b>{after * 100:.2f}%</b> after follow-up '
                                           f'({s["received_later"]:,} received later)</div>',)
            red = " dk-red" if s["overdue"] else ""
            if s["total"]:
                # 30/10/2026: "4 / 75" then "Checked", the open items per type, and the money still open
                rows = "".join(f'<div><span>{t}</span><b>{s["by_type"][t][0]:,}</b></div>'
                               for t in ("Pending", "Paid twice"))
                sub = (f'Checked<div class="dk-fu">{rows}</div>'
                       + (f'{escape(fu["open_money"])} remaining' if s["open"] else "Nothing left to check"))
                cards.append(("FOLLOW-UP", "fu", f'<span class="{red.strip()}">{s["resolved"]:,} / {s["total"]:,}'
                                                 f'</span>', sub))
            else:
                cards.append(("FOLLOW-UP", "fu", '<span class="dk-msg">Nothing to check</span>',
                              "No pending or paid-twice payment after the report"))
        else:
            cards.append(("FOLLOW-UP", "fu", f'<span class="dk-msg">Not tracked for this {escape(fu.get("unit", "period"))}</span>',
                          "We track what happened to pending and paid-twice payments after the report"))
    return (f'{STYLE_D}<div class="dk" style="grid-template-columns:repeat({len(cards)},minmax(0,1fr))">' + "".join(
        f'<div class="dk-c"><div class="dk-h"><span>{label}</span>{KPI_ICON[ic]}</div>'
        f'<div class="dk-v{" dk-v2" if "<br>" in big else ""}">{big}</div><div class="dk-s">{sub}</div></div>'
        for label, ic, big, sub in cards) + "</div>")


def lines(labels: list, series: list, partial: list | None = None, fill: bool = False) -> str:
    """A line chart: series [(name, [values], colour)] over the labels (a '*' marks a partial period)."""
    W, H, L, R, T, B = 600, 270, 46, 588, 30, 232
    n = len(labels)
    if not n:
        return '<div class="dv-empty">No period to show.</div>'
    top, step = _nice(max([v for _, vals, _ in series for v in vals if v is not None and not pd.isna(v)] + [1]))
    xs = [L + 18 + (i * (R - L - 36) / (n - 1) if n > 1 else (R - L - 36) / 2) for i in range(n)]
    g = _axis(L, R, T, B, top, step)
    for name, vals, col in series:
        here = [(x, v, lab) for x, v, lab in zip(xs, vals, labels) if v is not None and not pd.isna(v)]
        xs_s, vals, labs = [h[0] for h in here], [h[1] for h in here], [h[2] for h in here]
        pts = [(x, B - v / top * (B - T)) for x, v in zip(xs_s, vals)]
        if len(pts) > 1:
            path = _smooth(pts, B)
            if fill:
                g += (f'<defs><linearGradient id="dlf" x1="0" x2="0" y1="0" y2="1"><stop offset="0" stop-color="{col}" '
                      f'stop-opacity=".18"/><stop offset="1" stop-color="{col}" stop-opacity="0"/></linearGradient></defs>'
                      f'<path d="{path} L{pts[-1][0]:.1f},{B} L{pts[0][0]:.1f},{B} Z" fill="url(#dlf)"/>')
            g += f'<path d="{path}" fill="none" stroke="{col}" stroke-width="2.4"/>'
        g += "".join(f'<circle cx="{x:.1f}" cy="{y:.1f}" r="3.2" fill="{col}"><title>{escape(name)} · '
                     f'{escape(lab)}: {v:,.0f}</title></circle>' for (x, y), v, lab in zip(pts, vals, labs))
    every = max(1, math.ceil(n / 12))
    for i, (x, lab) in enumerate(zip(xs, labels)):
        if i % every == 0 or i == n - 1:
            star = "*" if partial and partial[i] else ""
            g += f'<text x="{x:.1f}" y="{B + 18}" class="dv-ax" text-anchor="middle">{escape(lab)}{star}</text>'
    return _img(f'<svg viewBox="0 0 {W} {H}">{g}</svg>')


def dual_bars(labels: list, rwf: list, usd: list, partial: list | None = None) -> str:
    """RWF and USD side by side per period, each on its own scale (RWF on the left axis, USD on the right):
    the two currencies are never added or drawn on one scale."""
    W, H, L, R, T, B = 600, 270, 52, 548, 30, 232
    n = len(labels)
    if not n:
        return '<div class="dv-empty">No period to show.</div>'
    top_r, step_r = _nice(max(rwf + [1]))
    top_u, step_u = _nice(max(usd + [1]))
    g, v = "", 0.0
    while v <= top_r + 1e-9:
        y = B - v / top_r * (B - T)
        g += (f'<line x1="{L}" x2="{R}" y1="{y:.1f}" y2="{y:.1f}" class="dv-gl"/>'
              f'<text x="{L - 8}" y="{y + 4:.1f}" class="dv-ax" text-anchor="end">{short(v) if v else "0"}</text>')
        v += step_r
    v = 0.0
    while v <= top_u + 1e-9:
        y = B - v / top_u * (B - T)
        g += f'<text x="{R + 8}" y="{y + 4:.1f}" class="dv-ax" fill="{USD_C}">{short(v) if v else "0"}</text>'
        v += step_u
    g += (f'<text x="{L - 8}" y="12" class="dv-ax" text-anchor="end" font-weight="800" fill="{TEAL}">RWF</text>'
          f'<text x="{R + 8}" y="12" class="dv-ax" font-weight="800" fill="{USD_C}">USD</text>')
    slot = (R - L) / n
    bw = min(22, slot * .32)
    for i, lab in enumerate(labels):
        cx = L + slot * (i + .5)
        for val, top, col, dx, cur in ((rwf[i], top_r, TEAL, -bw - 1.5, "RWF"), (usd[i], top_u, USD_C, 1.5, "USD")):
            h = val / top * (B - T) if top else 0
            g += (f'<rect x="{cx + dx:.1f}" y="{B - h:.1f}" width="{bw:.1f}" height="{h:.1f}" rx="3" fill="{col}">'
                  f'<title>{escape(lab)} · {cur} {val:,.0f}</title></rect>')
        star = "*" if partial and partial[i] else ""
        g += f'<text x="{cx:.1f}" y="{B + 18}" class="dv-ax" text-anchor="middle">{escape(lab)}{star}</text>'
    return _img(f'<svg viewBox="0 0 {W} {H}">{g}</svg>')


def chart_legend(items: list) -> str:
    return '<div class="dl">' + "".join(f'<span><i style="background:{c}"></i>{escape(n)}</span>'
                                         for n, c in items) + "</div>"


def _mchange(m) -> str:
    """' · ↗ 8%' / ' · new' after a merchant, when a period is compared with (columns Before, Change %, Note)."""
    note = getattr(m, "Note", "") or ""
    if note == "new":
        return ' · <span class="dk-g">new</span>'
    ch = getattr(m, "_6", None)                       # "Change %"
    if ch is None or pd.isna(ch):
        return ""
    return f' · <span class="{"dk-g" if ch >= 0 else "dk-r"}">{"↗" if ch >= 0 else "↘"} {abs(ch):.0f}%</span>'


def merchant_bars(table: pd.DataFrame, cur: str, n: int = 5) -> str:
    """The top merchants: the bar = share of the received notifications (%), then the number of transactions; the
    amount on hover."""
    rows = table.head(n)
    if rows.empty:
        return '<div class="dv-empty">No merchant data in the reports of this period (made before 12/10/2026).</div>'
    palette = [TEAL, ACC, USD_C, AMBER, "#7A8F94"]
    top = max(rows["Share %"].max(), 1e-9)
    return STYLE_D + '<div class="dm">' + "".join(
        f'<div class="dm-r" data-tip="{escape(str(m.Merchant))} · {cur} {m.Amount:,.0f} received · '
        f'{m.Transactions:,} transactions · {m._4:.1f}%">'
        f'<span class="dm-n"><i style="background:{palette[i % 5]}"></i>{escape(str(m.Merchant))}</span>'
        f'<span class="dm-b"><span style="width:{m._4 / top * 100:.1f}%"></span></span>'
        f'<span class="dm-v"><b>{m._4:.1f}%</b><span>{m.Transactions:,}{_mchange(m)}</span></span></div>'
        for i, m in enumerate(rows.itertuples())) + (
        f'<div class="dm-note">Bar: share of the {cur} received notifications; the figure after it: transactions. '
        f'Hover a merchant for its amount.</div></div>')


def _fmt_share(v: float, unit: str) -> str:
    return f"{v:,.0f}" if unit == "processed" else f"{unit} {short(v)}"


def bank_list(table: pd.DataFrame, colors: dict, unit: str) -> str:
    """28/10/2026: next to the banks' ring: each bank (its colour), its figure and its share, the biggest first,
    then the most used bank. unit: "processed" (transactions) or a currency (amount processed)."""
    if table.empty:
        return ""
    head = "Processed" if unit == "processed" else f"Amount ({unit})"
    cell = "padding:7px 0;border-bottom:1px solid #E6EEF0"
    rows = "".join(
        f'<tr><td style="{cell};padding-left:19px;text-indent:-19px"><i style="display:inline-block;width:11px;'
        f'height:11px;border-radius:3px;'
        f'background:{colors.get(r.Code, MUT)};margin-right:8px;vertical-align:-1px"></i>{escape(r.Bank)}</td>'
        f'<td style="{cell};text-align:right;white-space:nowrap;padding-left:14px">{_fmt_share(r.Value, unit)}'
        f'{_bchange(r)}</td>'
        f'<td style="{cell};text-align:right;font-weight:800;white-space:nowrap;padding-left:18px">{r._4:.1f}%</td></tr>'
        for r in table.itertuples())
    top = table.iloc[0]
    return (f'<div style="font-family:Nunito,\'Segoe UI\',sans-serif;font-size:13px;color:{TEAL}">'
            f'<table style="width:100%;border-collapse:collapse"><thead><tr style="color:{MUT};font-size:11.5px">'
            f'<th style="text-align:left;font-weight:600">Bank</th><th style="text-align:right;font-weight:600;'
            f'white-space:nowrap;padding-left:14px">'
            f'{head}</th><th style="text-align:right;font-weight:600;padding-left:18px">Share</th></tr></thead><tbody>{rows}</tbody>'
            f'</table><div style="margin-top:12px;font-weight:800">Most used: {escape(top["Bank"])} '
            f'({top["Share %"]:.1f}%)</div></div>')


def _bchange(r) -> str:
    """' ↗ 8%' after a bank's figure when a period is compared with."""
    ch = getattr(r, "_6", None)                         # "Change %"
    if ch is None or pd.isna(ch):
        return ""
    if abs(ch) < 0.5:
        return f' <span style="color:{MUT};font-size:11.5px">= 0%</span>'
    return (f' <span style="color:{"#2E7D4F" if ch >= 0 else BAD};font-size:11.5px">'
            f'{"↗" if ch >= 0 else "↘"} {abs(ch):.0f}%</span>')


def bank_ring(table: pd.DataFrame, colors: dict, unit: str) -> str:
    """28/10/2026: the banks' share as a ring (pie with a hole, the total inside) + the list next to it - the drawing
    used in the PDFs (the screen uses alt_bank_ring)."""
    if table.empty:
        return '<div class="dv-empty">No processed transaction in this period.</div>'
    total, cx, cy, ro, ri = float(table["Value"].sum()), 110, 110, 100, 62
    parts, a0 = [], -math.pi / 2
    for r in table.itertuples():
        a1 = a0 + 2 * math.pi * r.Value / total
        col = colors.get(r.Code, MUT)
        if r.Value >= total:                            # one bank only: a full ring
            parts.append(f'<circle cx="{cx}" cy="{cy}" r="{(ro + ri) / 2}" fill="none" stroke="{col}" '
                         f'stroke-width="{ro - ri}"/>')
        else:
            p = lambda a, rad: f"{cx + rad * math.cos(a):.2f} {cy + rad * math.sin(a):.2f}"
            big = 1 if a1 - a0 > math.pi else 0
            parts.append(f'<path d="M {p(a0, ro)} A {ro} {ro} 0 {big} 1 {p(a1, ro)} L {p(a1, ri)} '
                         f'A {ri} {ri} 0 {big} 0 {p(a0, ri)} Z" fill="{col}" stroke="#fff" stroke-width="2"/>')
        a0 = a1
    centre = f"{total:,.0f}" if unit == "processed" else short(total)
    sub = "processed" if unit == "processed" else f"{unit} processed"
    svg = (f'<svg viewBox="0 0 220 220">{"".join(parts)}<text x="{cx}" y="{cy + 2}" text-anchor="middle" '
           f'class="dv-dc">{centre}</text><text x="{cx}" y="{cy + 22}" text-anchor="middle" class="dv-ds">{sub}'
           f'</text></svg>')
    return (f'<div style="display:flex;gap:22px;align-items:center;flex-wrap:wrap">'
            f'<div style="flex:0 0 200px">{_img(svg, "width:200px;display:block")}</div>'
            f'<div style="flex:1 1 220px">{bank_list(table, colors, unit)}</div></div>')


def alt_bank_ring(table: pd.DataFrame, colors: dict, unit: str):
    """The same ring on the screen: hovering a bank shows its figure, its share (and the change when compared)."""
    import altair as alt
    names = list(table["Bank"])
    t = table.assign(Shown=[_fmt_share(v, unit) for v in table["Value"]])
    tips = [alt.Tooltip("Bank:N"), alt.Tooltip("Shown:N", title="Processed" if unit == "processed" else f"Amount ({unit})"),
            alt.Tooltip("Share %:Q", format=".1f")]
    if "Change %" in t.columns:
        tips.append(alt.Tooltip("Change %:Q", format="+.1f", title="Change % vs compared"))
    ring = alt.Chart(t).mark_arc(innerRadius=62, outerRadius=100, stroke="#fff", strokeWidth=2).encode(
        theta=alt.Theta("Value:Q", stack=True),
        color=alt.Color("Bank:N", scale=alt.Scale(domain=names, range=[colors.get(c, MUT) for c in t["Code"]]),
                        legend=None),
        order=alt.Order("Value:Q", sort="descending"), tooltip=tips)
    total = float(table["Value"].sum())
    centre = pd.DataFrame({"t": [f"{total:,.0f}" if unit == "processed" else short(total)],
                           "s": ["processed" if unit == "processed" else f"{unit} processed"]})
    big = alt.Chart(centre).mark_text(fontSize=22, fontWeight=800, color=TEAL, dy=-6).encode(text="t:N")
    small = alt.Chart(centre).mark_text(fontSize=11.5, color=MUT, dy=14).encode(text="s:N")
    return _alt_theme(alt.layer(ring, big, small), height=230)


def table_html(df: pd.DataFrame) -> str:
    """The numbers behind a chart, as a plain table (numbers right-aligned)."""
    head = "".join(f"<th>{escape(str(c))}</th>" for c in df.columns)
    blank = lambda v: v is None or (isinstance(v, float) and pd.isna(v)) or str(v) in ("nan", "None", "<NA>")
    body = "".join("<tr>" + "".join(
        '<td class="n">–</td>' if blank(v) else
        f'<td class="{"n" if isinstance(v, (int, float)) else ""}">'
        f'{escape(f"{v:,.2f}" if isinstance(v, float) and not float(v).is_integer() else f"{v:,.0f}" if isinstance(v, (int, float)) else str(v))}</td>'
        for v in row) + "</tr>" for row in df.itertuples(index=False))
    return f'{STYLE_D}<table class="dt"><thead><tr>{head}</tr></thead><tbody>{body}</tbody></table>'


def card_page(title: str, sub: str, scope: str, period: str, body: str, made: str) -> str:
    """One chart (or its table) on an A4 landscape page, with the scope and period, for its PDF."""
    return f'''<!doctype html><html><head><meta charset="utf-8">
<link href="https://fonts.googleapis.com/css2?family=Nunito:wght@400;600;700;800&display=swap" rel="stylesheet">
<style>@page {{size: A4 landscape; margin: 12mm;}} html, body {{margin: 0; background: #fff;
-webkit-print-color-adjust: exact; print-color-adjust: exact; font-family: Nunito, "Segoe UI", system-ui, sans-serif;
color: {TEAL};}} .pb {{background: {TEAL}; border-left: 8px solid {GREEN}; border-radius: 10px; color: #fff;
padding: 14px 22px; margin-bottom: 18px;}} .pk {{color: {GREEN}; font-weight: 800; font-size: 13px;}}
.pt {{font-size: 24px; font-weight: 800;}} .ps {{font-size: 13px; color: #D6E4E6;}} .sub {{color: {MUT};
font-size: 13px; margin: -8px 0 12px;}} img {{max-height: 120mm;}}</style></head><body>
<div class="pb"><div class="pk">{escape(scope.upper())} · {escape(period)}</div><div class="pt">{escape(title)}</div>
<div class="ps">IremboPay Reports Dashboard · generated on {escape(made)}</div></div>
<div class="sub">{escape(sub)}</div>{body}</body></html>'''


_KPI = {
    "money": '<path d="M4 7h16v10H4z" fill="none" stroke-width="1.6"/><circle cx="12" cy="12" r="2.4" fill="none" stroke-width="1.6"/>',
    "rate": '<path d="M5 19L19 5" stroke-width="1.6"/><circle cx="7" cy="7" r="2.2" fill="none" stroke-width="1.6"/><circle cx="17" cy="17" r="2.2" fill="none" stroke-width="1.6"/>',
    "tx": '<path d="M4 8h13l-3-3M20 16H7l3 3" fill="none" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/>',
    "fu": '<path d="M5 4h14v16H5z" fill="none" stroke-width="1.6"/><path d="M8 12l3 3 5-6" fill="none" stroke-width="1.6" stroke-linecap="round"/>',
    "sla": '<circle cx="12" cy="13" r="7" fill="none" stroke-width="1.6"/><path d="M12 9v4l3 2M10 3h4" fill="none" stroke-width="1.6" stroke-linecap="round"/>',
    "pending": '<path d="M7 3h10M7 21h10M8 3c0 5 8 5 8 9s-8 4-8 9M16 3c0 5-8 5-8 9s8 4 8 9" fill="none" stroke-width="1.6" stroke-linecap="round"/>'}
KPI_ICON = {k: _img(f'<svg viewBox="0 0 24 24" stroke="{MUT}">{v}</svg>', "width:20px;height:20px;display:block")
            for k, v in _KPI.items()}

STYLE_D = f'''<style>
.dk{{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:13px;font-family:"Nunito","Segoe UI",system-ui,sans-serif;color:{TEAL}}}
.dk-c{{background:#fff;border:1px solid #E1EEF0;border-radius:12px;padding:15px 17px;box-shadow:0 1px 3px rgba(18,50,57,.06);min-width:0}}
.dk-h{{display:flex;justify-content:space-between;align-items:center;font-size:11px;font-weight:700;letter-spacing:.06em;color:{MUT}}}
.dk-v{{font-size:26px;font-weight:800;margin:7px 0 5px;line-height:1.15}} .dk-v2{{font-size:19px}}
.dk-s{{font-size:11.5px;color:{MUT}}} .dk-f{{font-size:11px;color:#8A5B00;margin-top:5px}} .dk-f b{{color:{TEAL}}}
.dk-fu{{border-top:1px solid #E1EEF0;margin:8px 0 6px;padding-top:6px}}
.dk-fu div{{display:flex;justify-content:space-between;gap:8px;line-height:1.7}} .dk-fu b{{color:{TEAL};font-weight:700}}
.dk-red{{color:{BAD}}}.dk-msg{{font-size:16px;line-height:1.3;display:block}} .dk-g{{color:#2E7D4F;font-weight:800}} .dk-r{{color:{BAD};font-weight:800}} .dk-n{{color:{MUT}}}
.dl{{display:flex;flex-wrap:wrap;justify-content:center;gap:14px;font-size:11.5px;color:{TEAL};margin-top:4px;font-family:"Nunito",system-ui,sans-serif}}
.dl i{{display:inline-block;width:9px;height:9px;border-radius:50%;margin-right:6px}}
.dm{{font-family:"Nunito","Segoe UI",system-ui,sans-serif;color:{TEAL}}}
.dm-r{{display:grid;grid-template-columns:minmax(0,1.3fr) minmax(0,2fr) 170px;align-items:center;gap:14px;padding:11px 8px;
  border-radius:10px;position:relative;cursor:default;transition:background .12s}}
.dm-r:hover{{background:{PALE}}} .dm-r:hover .dm-b span{{background:{ACC}}}
.dm-r:hover::after{{content:attr(data-tip);position:absolute;right:8px;top:-28px;background:{TEAL};color:#fff;font-size:11px;
  font-weight:700;padding:5px 10px;border-radius:8px;white-space:nowrap;z-index:5;box-shadow:0 4px 12px rgba(18,50,57,.18)}}
.dm-n{{font-size:12.5px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}} .dm-n i{{display:inline-block;width:10px;height:10px;border-radius:50%;margin-right:9px}}
.dm-b{{height:10px;border-radius:99px;background:#EEF3F4;overflow:hidden}} .dm-b span{{display:block;height:100%;border-radius:99px;background:{TEAL}}}
.dm-v{{font-size:11.5px;color:{MUT};white-space:nowrap;display:flex;justify-content:flex-start;gap:18px}}
.dm-v b{{color:{TEAL};font-size:12.5px;min-width:46px;text-align:right}}   /* 28/10/2026: % | transactions, apart */
.dm-note{{font-size:12px;color:{MUT};margin-top:10px}}
.dt{{width:100%;border-collapse:collapse;font-family:"Nunito","Segoe UI",system-ui,sans-serif;font-size:12px;color:{TEAL}}}
.dt th{{text-align:left;font-weight:800;color:{MUT};border-bottom:1px solid #E1EEF0;padding:7px 8px}}
.dt td{{border-bottom:1px solid #F0F4F5;padding:7px 8px}} .dt td.n{{text-align:right;font-variant-numeric:tabular-nums}}
.dt th:not(:first-child){{text-align:right}}
@media (max-width:1100px){{.dk{{grid-template-columns:repeat(2,minmax(0,1fr))}}}}
</style>'''


# ---- the same charts, interactive on the screen (Altair: the figures follow the cursor); the PDF keeps the drawings
def _alt_theme(chart, height=250):
    return (chart.properties(height=height)
            .configure_view(strokeWidth=0)
            .configure_axis(labelColor=MUT, titleColor=MUT, gridColor="#E6EEF0", gridDash=[3, 4], domain=False,
                            tickColor="#E6EEF0", labelFont="Nunito", titleFont="Nunito", labelFontSize=11)
            .configure_legend(orient="bottom", labelColor=TEAL, labelFont="Nunito", labelFontSize=12, title=None,
                              symbolType="circle")
            .configure(font="Nunito"))


def alt_lines(table: pd.DataFrame, series: list, fill: bool = False):
    """Line chart of table columns series [(column, colour)] over table["Period"]; hovering a period shows every
    figure of that period (a vertical guide follows the cursor)."""
    import altair as alt
    names = [c for c, _ in series]
    order = list(table["Period"])
    long = table.melt(id_vars="Period", value_vars=names, var_name="Figure", value_name="Value")
    x = alt.X("Period:N", sort=order, title=None, axis=alt.Axis(labelAngle=0))
    color = alt.Color("Figure:N", scale=alt.Scale(domain=names, range=[c for _, c in series]),
                      legend=None)                      # the legend is drawn under the chart (chart_legend)
    hover = alt.selection_point(fields=["Period"], nearest=True, on="pointerover", clear="pointerout", empty=False)
    base = alt.Chart(long).encode(x=x, y=alt.Y("Value:Q", title=None, axis=alt.Axis(format=",.0f", tickMinStep=1)),
                                  color=color)
    layers = []
    if fill and len(series) == 1:
        layers.append(base.mark_area(interpolate="monotone", opacity=.15))
    layers.append(base.mark_line(interpolate="monotone", strokeWidth=2.6))
    layers.append(base.mark_circle(size=40).encode(opacity=alt.value(1)))
    tips = [alt.Tooltip("Period:N")] + [alt.Tooltip(f"{c}:Q", format=",") for c in table.columns if c != "Period"]
    rule = (alt.Chart(table).mark_rule(color=TEAL, strokeDash=[4, 4], opacity=.5)
            .encode(x=x, opacity=alt.condition(hover, alt.value(.6), alt.value(0)), tooltip=tips).add_params(hover))
    points = base.mark_circle(size=110).encode(opacity=alt.condition(hover, alt.value(1), alt.value(0))).transform_filter(hover)
    return _alt_theme(alt.layer(*layers, rule, points))


def alt_dual_bars(table: pd.DataFrame):
    """RWF and USD received per period side by side, each on its own axis (RWF left, USD right); hovering a bar
    shows both amounts of that period."""
    import altair as alt
    order = list(table["Period"])
    x = alt.X("Period:N", sort=order, title=None, axis=alt.Axis(labelAngle=0))
    tips = [alt.Tooltip("Period:N"), alt.Tooltip("RWF received:Q", format=",.0f"),
            alt.Tooltip("USD received:Q", format=",.2f")]
    rwf = alt.Chart(table).mark_bar(xOffset=-9, width=16, cornerRadiusTopLeft=3, cornerRadiusTopRight=3, color=TEAL).encode(
        x=x, y=alt.Y("RWF received:Q", title="RWF", axis=alt.Axis(format="~s", titleColor=TEAL, orient="left")),
        tooltip=tips)
    usd = alt.Chart(table).mark_bar(xOffset=9, width=16, cornerRadiusTopLeft=3, cornerRadiusTopRight=3, color=USD_C).encode(
        x=x, y=alt.Y("USD received:Q", title="USD", axis=alt.Axis(format="~s", titleColor=USD_C, orient="right")),
        tooltip=tips)
    return _alt_theme(alt.layer(rwf, usd).resolve_scale(y="independent"))


FU_C = {"Resolved": "#2E7D4F", "Still open": AMBER, "Overdue": BAD}


def stacked_bars(labels: list, parts: list) -> str:
    """Stacked bars per period: parts [(name, [values], colour)] (drawn for the PDF)."""
    W, H, L, R, T, B = 600, 270, 46, 588, 30, 232
    n = len(labels)
    if not n:
        return '<div class="dv-empty">No period to show.</div>'
    totals = [sum(p[1][i] for p in parts) for i in range(n)]
    top, step = _nice(max(totals + [1]))
    g = _axis(L, R, T, B, top, step)
    slot = (R - L) / n
    bw = min(34, slot * .5)
    for i, lab in enumerate(labels):
        x, y = L + slot * (i + .5) - bw / 2, B
        for name, vals, col in parts:
            h = vals[i] / top * (B - T)
            if h:
                g += (f'<rect x="{x:.1f}" y="{y - h:.1f}" width="{bw:.1f}" height="{h:.1f}" fill="{col}">'
                      f'<title>{escape(lab)} · {escape(name)}: {vals[i]:,}</title></rect>')
                y -= h
        g += f'<text x="{x + bw / 2:.1f}" y="{B + 18}" class="dv-ax" text-anchor="middle">{escape(lab)}</text>'
    return _img(f'<svg viewBox="0 0 {W} {H}">{g}</svg>')


def alt_stacked(table: pd.DataFrame, parts: list):
    """The same stacked bars, interactive: hovering a bar shows every figure of its period."""
    import altair as alt
    names = [n for n, _ in parts]
    long = table.melt(id_vars="Period", value_vars=names, var_name="Status", value_name="Items")
    tips = [alt.Tooltip("Period:N")] + [alt.Tooltip(f"{c}:Q", format=",") for c in table.columns if c != "Period"]
    chart = (alt.Chart(long).mark_bar(size=34).encode(
        x=alt.X("Period:N", sort=list(table["Period"]), title=None, axis=alt.Axis(labelAngle=0)),
        y=alt.Y("Items:Q", stack=True, title=None, axis=alt.Axis(format=",d", tickMinStep=1)),
        color=alt.Color("Status:N", scale=alt.Scale(domain=names, range=[c for _, c in parts]), legend=None),
        order=alt.Order("order:Q")).transform_calculate(order=f"indexof({names!r}, datum.Status)"))
    hover = (alt.Chart(table).mark_bar(size=34, opacity=0).encode(
        x=alt.X("Period:N", sort=list(table["Period"])), y=alt.Y("Total:Q"), tooltip=tips))
    return _alt_theme(alt.layer(chart, hover))


# ---- comparison mode (23/10/2026): this period against the one compared with, lined up by position -----------------
PREV = "#9FB3B7"                                         # the period compared with: a lighter colour (dashed line)


def position_labels(kind: str, n: int) -> list:
    """How periods are lined up when two are compared: by their place in the period (Week 1 with Week 1)."""
    if kind == "week":
        return ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"][:n] + [f"Day {i}" for i in range(8, n + 1)]
    unit = {"month": "Week", "quarter": "Month", "year": "Q"}[kind]
    return [f"{unit} {i}" if unit != "Q" else f"Q{i}" for i in range(1, n + 1)]


def group_bars(labels: list, series: list) -> str:
    """Bars side by side per label, one per series [(name, [values], colour)], one scale (drawn for the PDF)."""
    W, H, L, R, T, B = 600, 270, 52, 588, 30, 232
    n = len(labels)
    if not n:
        return '<div class="dv-empty">No period to show.</div>'
    top, step = _nice(max([v for _, vals, _ in series for v in vals if v is not None and not pd.isna(v)] + [1]))
    g = _axis(L, R, T, B, top, step)
    slot = (R - L) / n
    bw = min(20, slot * .7 / max(len(series), 1))
    for i, lab in enumerate(labels):
        x0 = L + slot * (i + .5) - bw * len(series) / 2
        for j, (name, vals, col) in enumerate(series):
            v = vals[i] if i < len(vals) and vals[i] is not None and not pd.isna(vals[i]) else 0
            h = v / top * (B - T)
            g += (f'<rect x="{x0 + j * bw:.1f}" y="{B - h:.1f}" width="{bw - 2:.1f}" height="{h:.1f}" rx="2" '
                  f'fill="{col}"><title>{escape(lab)} · {escape(name)}: {v:,.0f}</title></rect>')
        g += f'<text x="{L + slot * (i + .5):.1f}" y="{B + 18}" class="dv-ax" text-anchor="middle">{escape(lab)}</text>'
    return _img(f'<svg viewBox="0 0 {W} {H}">{g}</svg>')


def _tips(table: pd.DataFrame, fmt: str) -> list:
    """Every column of a table in the tooltip: numbers formatted, text as it is."""
    import altair as alt
    out = []
    for c in table.columns:
        num = pd.api.types.is_numeric_dtype(pd.to_numeric(table[c], errors="coerce")) and \
            pd.to_numeric(table[c], errors="coerce").notna().any() and c not in ("Position", "Exception")
        out.append(alt.Tooltip(f"{c}:Q", format=fmt) if num else alt.Tooltip(f"{c}:N"))
    return out


def alt_compare_lines(table: pd.DataFrame, now: str, before: str, color: str):
    """Two lines over the same positions: this period (solid) and the one compared with (dashed, lighter); hovering
    a point (or near it) shows both periods' figures, their dates and the change."""
    import altair as alt
    table = table.copy()
    for c in (now, before, "Change %"):                  # numbers (a missing week = NaN, not a mixed column)
        if c in table:
            table[c] = pd.to_numeric(table[c], errors="coerce")
    order = list(table["Position"])
    long = table.melt(id_vars="Position", value_vars=[now, before], var_name="Period", value_name="Value")
    long = long.dropna(subset=["Value"]).merge(table, on="Position")
    x = alt.X("Position:N", sort=order, title=None, axis=alt.Axis(labelAngle=0))
    col = alt.Color("Period:N", scale=alt.Scale(domain=[now, before], range=[color, PREV]), legend=None)
    dash = alt.StrokeDash("Period:N", scale=alt.Scale(domain=[now, before], range=[[1, 0], [6, 4]]), legend=None)
    base = alt.Chart(long).encode(x=x, y=alt.Y("Value:Q", title=None, axis=alt.Axis(format=",.0f")), color=col)
    tips = _tips(table, ",")
    hover = alt.selection_point(fields=["Position"], nearest=True, on="pointerover", clear="pointerout", empty=False)
    points = base.mark_circle(size=60).encode(
        size=alt.condition(hover, alt.value(140), alt.value(60)), tooltip=tips).add_params(hover)
    # the whole height above each position answers the cursor (an invisible wide band), not only the small points
    band = alt.Chart(table).mark_rule(strokeWidth=36, opacity=0.001).encode(x=x, tooltip=tips)
    return _alt_theme(alt.layer(base.mark_line(interpolate="monotone", strokeWidth=2.6).encode(strokeDash=dash),
                                points, band))


def alt_group_bars(table: pd.DataFrame, cat: str, series: list, horizontal: bool = False):
    """Bars side by side per category (series [(column, colour)]), one scale; hovering shows every figure."""
    import altair as alt
    names = [c for c, _ in series]
    long = table.melt(id_vars=cat, value_vars=names, var_name="Series", value_name="Value")
    tips = _tips(table, ",.0f")
    long = long.merge(table, on=cat)
    color = alt.Color("Series:N", scale=alt.Scale(domain=names, range=[c for _, c in series]), legend=None)
    if horizontal:
        chart = alt.Chart(long).mark_bar(cornerRadiusEnd=3).encode(
            y=alt.Y(f"{cat}:N", sort=list(table[cat]), title=None, axis=alt.Axis(labelLimit=220)),
            yOffset=alt.YOffset("Series:N", sort=names),
            x=alt.X("Value:Q", title=None, axis=alt.Axis(format=",.0f")), color=color, tooltip=tips)
    else:
        chart = alt.Chart(long).mark_bar(cornerRadiusTopLeft=3, cornerRadiusTopRight=3).encode(
            x=alt.X(f"{cat}:N", sort=list(table[cat]), title=None, axis=alt.Axis(labelAngle=0)),
            xOffset=alt.XOffset("Series:N", sort=names),
            y=alt.Y("Value:Q", title=None, axis=alt.Axis(format="~s")), color=color, tooltip=tips)
    return _alt_theme(chart)


# =====================================================================================================================
# PENDING & DELAYED NOTIFICATIONS (27/10/2026): per period, against the processed transactions
#     Pending = debited at the bank, never notified to IremboPay (at report time) · Delayed = SLA violated
# =====================================================================================================================
NOTIF = [("Delayed % (after SLA)", AMBER), ("Pending % (never notified)", BAD)]


def _money(money: dict, metric: str) -> str:
    parts = [f"{c} {short(a.get(metric, 0.0))}" for c, a in money.items() if a.get(metric)]
    return " · ".join(parts) or "–"


def notif_table(child: list, labels: list) -> pd.DataFrame:
    """One row per period + a Total row: processed (count, amount), pending and delayed (count, % of processed,
    amount), received rate - all from the reports' figures."""
    pct = lambda n, d: round(n / d * 100, 2) if d else None
    rows, tot, money = [], {"processed": 0, "pending": 0, "sla": 0, "received": 0}, {}
    for c, lab in zip(child, labels):
        m = c.get("money", {})
        for cur, a in m.items():
            for k, val in a.items():
                money.setdefault(cur, {}).setdefault(k, 0.0)
                money[cur][k] += val
        for k in tot:
            tot[k] += c[k]
        rows.append((lab, c, m))
    rows.append(("Total", tot, money))
    return pd.DataFrame([{"Period": lab, "Processed": int(t["processed"]), "Processed amount": _money(m, "processed"),
                          "Pending (never notified)": int(t["pending"]), "Pending %": pct(t["pending"], t["processed"]),
                          "Pending amount": _money(m, "pending"), "Delayed (after SLA)": int(t["sla"]),
                          "Delayed %": pct(t["sla"], t["processed"]),
                          "Received rate %": pct(t["received"], t["processed"])} for lab, t, m in rows])
