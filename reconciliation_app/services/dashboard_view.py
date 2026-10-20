"""
THE DASHBOARD'S MAIN VIEW (HTML): headline, four key cards, the payments per week, the payments and amount
overviews, top merchants, exceptions by type, late payments by bank, needs attention.
Every number, curve, slice and line is computed from the facts of the saved reports (services/dashboard.py);
nothing is typed by hand. IremboPay colours: white / pale teal surfaces, dark teal text and figures, brand green for
results that are good, red / amber for problems only.

    view(...)   (headline HTML, grid HTML) of the view, for st.html
"""
from __future__ import annotations

import base64
import math
from html import escape

import pandas as pd

import engine
from services import dashboard as d

TEAL, ACC, PALE, GREEN, AMBER, AMBER_L, BAD, MUT = ("#123239", "#2B8A99", "#EAF4F5", "#99E378", "#D4A72C", "#E9C46A",
                                                     "#C2413D", "#5E7378")


SVG_STYLE = (f"text{{font-family:Nunito,'Segoe UI',Helvetica,Arial,sans-serif}} .dv-ax{{font-size:10.5px;fill:{MUT}}} "
             f".dv-gl{{stroke:#E6EEF0;stroke-dasharray:3 4}} .dv-t1{{font-size:10.5px;fill:#CFE3E6}} "
             f".dv-t2{{font-size:15px;fill:#fff;font-weight:800}} .dv-dc{{font-size:28px;font-weight:800;fill:{TEAL}}} "
             f".dv-ds{{font-size:11.5px;fill:{MUT}}}")


def _img(svg: str, style: str = "width:100%;display:block") -> str:
    """An SVG drawing as an <img> (Streamlit removes inline <svg> from its HTML); its text styles go inside it."""
    svg = svg.replace("<svg ", f'<svg xmlns="http://www.w3.org/2000/svg" ', 1)
    svg = svg.replace(">", f"><style>{SVG_STYLE}</style>", 1)
    return f'<img alt="" style="{style}" src="data:image/svg+xml;base64,{base64.b64encode(svg.encode()).decode()}"/>'


def short(v: float) -> str:
    for size, unit in ((1e9, "B"), (1e6, "M"), (1e3, "K")):
        if abs(v) >= size:
            return f"{v / size:,.1f}{unit}"
    return f"{v:,.0f}"


def pill(text: str, kind: str) -> str:
    return f'<span class="dv-pill {kind}">{escape(text)}</span>'


def change_pill(now, before, good_up=True, points=False):
    """'21.4% ↓' (red when bad, green when good); None when there is nothing to compare with."""
    if now is None or before is None:
        return None
    if points:
        diff = (now - before) * 100
        if abs(diff) < 0.005:
            return pill("= 0.00 pts", "n")
        return pill(f"{diff:+.2f} pts {'↑' if diff > 0 else '↓'}", "g" if (diff > 0) == good_up else "r")
    if not before:
        return None
    x = (now - before) / before * 100
    if abs(x) < 0.05:
        return pill("= 0.0%", "n")
    return pill(f"{abs(x):.1f}% {'↑' if x > 0 else '↓'}", "g" if (x > 0) == good_up else "r")


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


def area(series, color, gid, note=None):
    """series [(label, value, partial)]; note = (index, title, value, sub) for the callout."""
    W, H, L, R, T, B = 560, 270, 46, 548, 18, 232
    if len(series) < 2:
        return '<div class="dv-empty">Not enough periods to draw a curve.</div>'
    top, step = _nice(max(v for _, v, _ in series))
    xs = [L + 8 + i * (R - L - 16) / (len(series) - 1) for i in range(len(series))]
    pts = [(x, B - v / top * (B - T)) for x, (_, v, _) in zip(xs, series)]
    line = _smooth(pts, B)
    g = (f'<defs><linearGradient id="{gid}" x1="0" x2="0" y1="0" y2="1"><stop offset="0" stop-color="{ACC}" '
         f'stop-opacity=".22"/><stop offset="1" stop-color="{ACC}" stop-opacity="0"/></linearGradient></defs>'
         + _axis(L, R, T, B, top, step)
         + f'<path d="{line} L{pts[-1][0]:.1f},{B} L{pts[0][0]:.1f},{B} Z" fill="url(#{gid})"/>'
         + f'<path d="{line}" fill="none" stroke="{color}" stroke-width="2.6"/>')
    g += "".join(f'<circle cx="{x:.1f}" cy="{y:.1f}" r="3" fill="{color}"/>' for x, y in pts)
    every = max(1, math.ceil(len(series) / 12))
    for i, (x, (lab, _, part)) in enumerate(zip(xs, series)):
        if i % every == 0 or i == len(series) - 1:
            g += f'<text x="{x:.1f}" y="{B + 18}" class="dv-ax" text-anchor="middle">{escape(lab)}{"*" if part else ""}</text>'
    if note is not None:
        i, title, value, sub = note
        x, y = pts[i]
        bx = x + 12 if x < R - 175 else x - 177
        by = max(min(y - 10, B - 70), T)
        g += (f'<line x1="{x:.1f}" x2="{x:.1f}" y1="{T}" y2="{B}" stroke="{TEAL}" stroke-dasharray="4 4" opacity=".45"/>'
              f'<circle cx="{x:.1f}" cy="{y:.1f}" r="6" fill="{color}" stroke="#fff" stroke-width="2"/>'
              f'<rect x="{bx:.1f}" y="{by:.1f}" width="165" height="62" rx="10" fill="{TEAL}"/>'
              f'<text x="{bx + 12:.1f}" y="{by + 19:.1f}" class="dv-t1">{escape(title)}</text>'
              f'<text x="{bx + 12:.1f}" y="{by + 39:.1f}" class="dv-t2">{escape(value)}</text>'
              f'<text x="{bx + 12:.1f}" y="{by + 54:.1f}" class="dv-t1">{escape(sub)}</text>')
    return _img(f'<svg viewBox="0 0 {W} {H}">{g}</svg>')


def bars(series):
    W, H, L, R, T, B = 330, 170, 34, 324, 10, 146
    if not series:
        return '<div class="dv-empty">No period to show.</div>'
    top, step = _nice(max(v for _, v, _ in series))
    bw = (R - L) / len(series)
    g = (f'<defs><linearGradient id="dvbar" x1="0" x2="0" y1="0" y2="1"><stop offset="0" stop-color="{ACC}"/>'
         f'<stop offset="1" stop-color="{TEAL}"/></linearGradient></defs>' + _axis(L, R, T, B, top, step))
    for i, (lab, v, part) in enumerate(series):
        x = L + i * bw + 6
        h = v / top * (B - T) if top else 0
        g += (f'<rect x="{x:.1f}" y="{B - h:.1f}" width="{bw - 12:.1f}" height="{h:.1f}" rx="5" '
              f'fill="{"#B9D3D7" if part else "url(#dvbar)"}"><title>{escape(lab)}: {v:,.0f}</title></rect>'
              f'<text x="{x + (bw - 12) / 2:.1f}" y="{B + 16}" class="dv-ax" text-anchor="middle">{escape(lab)}</text>')
    return _img(f'<svg viewBox="0 0 {W} {H}">{g}</svg>')


def donut(parts, center, sub, size=210):
    cx = cy = size / 2
    r = size / 2 - 20
    total = sum(v for _, v, _ in parts)
    if not total:
        return _img(f'<svg viewBox="0 0 {size} {size}"><circle cx="{cx}" cy="{cy}" r="{r}" '
                f'fill="none" stroke="{GREEN}" stroke-width="24"/><text x="{cx}" y="{cy + 4}" class="dv-dc" '
                f'text-anchor="middle">0</text><text x="{cx}" y="{cy + 26}" class="dv-ds" text-anchor="middle">'
                f'{escape(sub)}</text></svg>', f"width:{size}px;flex:none")
    a, g = -math.pi / 2, ""
    for name, v, col in parts:
        if not v:
            continue
        frac = v / total
        if frac > .999:
            g += f'<circle cx="{cx}" cy="{cy}" r="{r}" fill="none" stroke="{col}" stroke-width="24"/>'
            continue
        b = a + frac * 2 * math.pi
        gap = .03 if frac > .03 else 0
        x1, y1 = cx + r * math.cos(a), cy + r * math.sin(a)
        x2, y2 = cx + r * math.cos(b - gap), cy + r * math.sin(b - gap)
        g += (f'<path d="M{x1:.1f},{y1:.1f} A{r},{r} 0 {1 if frac > .5 else 0} 1 {x2:.1f},{y2:.1f}" fill="none" '
              f'stroke="{col}" stroke-width="24"><title>{escape(name)}: {v:,}</title></path>')
        a = b
    g += (f'<text x="{cx}" y="{cy + 4}" class="dv-dc" text-anchor="middle">{escape(center)}</text>'
          f'<text x="{cx}" y="{cy + 26}" class="dv-ds" text-anchor="middle">{escape(sub)}</text>')
    return _img(f'<svg viewBox="0 0 {size} {size}">{g}</svg>', f"width:{size}px;flex:none")


def legend(parts, unit):
    total = sum(v for _, v, _ in parts) or 1
    return "".join(f'<div class="dv-lg"><span class="dv-dot" style="background:{c}"></span><div><b>{escape(n)} '
                   f'({v / total * 100:.1f}%)</b><div class="dv-mut">{v:,} {unit}</div></div></div>'
                   for n, v, c in parts if v)


_ICON = {
    "rec": f'<svg viewBox="0 0 40 40"><circle cx="20" cy="20" r="18" fill="{GREEN}"/><path d="M12 21l5 5 11-12" '
           f'stroke="{TEAL}" stroke-width="3.2" fill="none" stroke-linecap="round" stroke-linejoin="round"/></svg>',
    "ont": f'<svg viewBox="0 0 40 40"><circle cx="20" cy="20" r="18" fill="{PALE}"/><circle cx="20" cy="20" r="10" '
           f'stroke="{TEAL}" stroke-width="3" fill="none"/><path d="M20 14v7l4 3" stroke="{TEAL}" stroke-width="3" '
           f'fill="none" stroke-linecap="round"/></svg>',
    "amt": f'<svg viewBox="0 0 40 40"><rect x="3" y="9" width="34" height="22" rx="5" fill="{PALE}"/>'
           f'<rect x="3" y="14" width="34" height="5" fill="{ACC}"/><rect x="8" y="23" width="10" height="3" rx="1.5" '
           f'fill="{TEAL}"/></svg>',
    "risk": f'<svg viewBox="0 0 40 40"><path d="M20 3l17 30H3z" fill="#FBF0DC"/><path d="M20 14v9" stroke="{TEAL}" '
            f'stroke-width="3.2" stroke-linecap="round"/><circle cx="20" cy="28" r="2" fill="{TEAL}"/></svg>'}
ICON = {k: _img(v, "width:38px;height:38px;display:block") for k, v in _ICON.items()}


# ---- the view --------------------------------------------------------------------------------------------------------
def view(facts, sel, bef, now, before, now_by, bef_by, kind, first, last, span, span_before, scope, cur,
         no_change: str = "first period") -> tuple:
    """(headline HTML, grid HTML) of the main view for the period and banks chosen (sel = their facts, bef = the period compared)."""
    base = "month" if kind == "custom" else kind
    unit = "week" if base in ("month", "week", "day") else "month"
    end = pd.Timestamp(last if last is not None else first)
    end = d.period_end(d.period_key(end, base), base)
    banks = sorted({f.bank for f in sel}, key=lambda c: list(engine.BANKS).index(c) if c in engine.BANKS else 99)
    r, rb = d.rates(now), (d.rates(before) if before else None)
    amt_cur = "RWF" if "RWF" in now["currencies"] else (now["currencies"] or ["RWF"])[0]
    amt = lambda t, c, m="processed": (t or {}).get("amounts", {}).get(c, {}).get(m, 0.0) if t else None

    # ---- headline
    ch = change_pill(now["processed"], before["processed"] if before else None)
    head = (f'<div class="dv-hl">Payments processed · {escape(scope.lower() if scope == "All banks" else scope)}</div>'
            f'<div class="dv-hv">{now["processed"]:,} {ch or ""}</div>'
            f'<div class="dv-hp">' + (f'Prev. <b>{before["processed"]:,}</b> &nbsp;{escape(span_before)} &nbsp;·&nbsp; '
                                     if before else f'{escape(span)} &nbsp;·&nbsp; ')
            + f'{now["received"]:,} received by Irembo · {now["sla"]:,} late · {now["pending"]:,} pending</div>')

    # ---- the four cards
    risk = {c: a.get("pending", 0.0) + a.get("duplicate", 0.0) for c, a in now["amounts"].items()}
    risk_txt = " · ".join(f"{c} {v:,.0f}" for c, v in risk.items() if v) or "None"
    usd = amt(now, "USD") if "USD" in now["currencies"] and amt_cur != "USD" else None
    DEFINE = {"rec": "Notified = found and confirmed paid in the IremboPay export.",
             "ont": "On time = paid_at within the bank's SLA of its payment date.",
             "amt": "The amount of the payments counted as Processed.",
             "risk": "Pending + paid with another channel: not yet confirmed received."}
    cards = [
        ("rec", "Notified", "N/A" if r["receipt"] is None else f"{r['receipt'] * 100:.2f}%",
         change_pill(r["receipt"], rb["receipt"] if rb else None, points=True) or pill(no_change, "n"),
         f"target ≥ {d.pct_target(d.TARGET_RECEIPT)} · {now['processed'] - now['received']:,} not yet notified"),
        ("ont", "On time (SLA)", "N/A" if r["ontime"] is None else f"{r['ontime'] * 100:.1f}%",
         change_pill(r["ontime"], rb["ontime"] if rb else None, points=True) or pill(no_change, "n"),
         f"target ≥ {d.pct_target(d.TARGET_ONTIME)} · {now['sla']:,} late"),
        ("amt", f"Amount collected · {amt_cur}", short(amt(now, amt_cur)),
         change_pill(amt(now, amt_cur), amt(before, amt_cur) if before else None) or pill(no_change, "n"),
         (f'USD {short(usd)} ' + (change_pill(usd, amt(before, "USD") if before else None) or "")) if usd is not None
         else f"{amt_cur} only · never added to another currency"),
        ("risk", "Money at risk", risk_txt if len(risk_txt) < 18 else f"{amt_cur} {short(risk.get(amt_cur, 0.0))}",
         pill("to follow up", "a") if any(risk.values()) else pill("nothing at risk", "g"),
         f"{now['pending']:,} pending · {now['duplicate']:,} paid twice")]
    card_html = "".join(f'<div class="dv-card dv-k"><div class="dv-ic">{ICON[i]}</div><div class="dv-kl">{escape(l)} '
                        f'<span class="dv-info" title="{escape(DEFINE[i])}">ⓘ</span></div>'
                        f'<div class="dv-kv">{escape(v)}</div><div class="dv-kf">{p}<span class="dv-go">›</span>'
                        f'</div><div class="dv-ks">{s if "<span" in s else escape(s)}</div></div>'
                        for i, l, v, p, s in cards)

    # ---- per week (or per month for a quarter / a year), up to the end of the period
    tr = d.trend(facts, unit, end, 11, cur, banks or None)
    if len(tr):
        per = tr.groupby(["period", "label"], as_index=False).agg(p=("processed", "sum"), part=("partial", "any"))
        tra = d.trend(facts, unit, end, 11, amt_cur, banks or None)
        amt_per = tra.groupby(["period", "label"], as_index=False)["amount"].sum()
        series = [(row.label, int(row.p), bool(row.part)) for row in per.itertuples()]
        aseries = [(row.label, float(row.amount), False) for row in amt_per.itertuples()]
    else:
        series, aseries = [], []
    peak = max(range(len(series)), key=lambda i: series[i][1]) if series else None
    note = ((peak, f"{'Week of ' if unit == 'week' else ''}{series[peak][0]}", f"{series[peak][1]:,} payments",
             f"peak of the {len(series)} {unit}s") if series else None)
    week_bars = bars(series[-7:])
    unit_t = unit.capitalize()

    # ---- merchants of the period (the reports that have merchant data)
    mt, missing = d.merchants(sel, bef, amt_cur)
    with_data = sorted({engine.BANKS[f.bank].NAME for f in sel if f.merchants is not None and f.bank in engine.BANKS})
    merch_sub = escape(span) + ((" · " + ", ".join(escape(x) for x in with_data)
                                 + (f" (the only report{'s' if len(with_data) != 1 else ''} with merchant data)"
                                    if missing else "")) if with_data else "")
    merch_html = "".join(f'<div class="dv-row"><span class="dv-ri">{escape(str(m.Merchant)[:1])}</span><div>'
                         f'<div class="dv-rn">{escape(str(m.Merchant))}</div><div class="dv-rs">{m._4:.1f}% of '
                         f'{amt_cur} notifications · {m.Transactions:,}</div></div>{pill(f"{amt_cur} {short(m.Amount)}", "g")}</div>'
                         for m in mt.head(3).itertuples()) or \
        '<div class="dv-empty">No merchant data in the reports of this period (made before 12/10/2026).</div>'

    # ---- exceptions of the period
    exc = [("Late (SLA)", now["sla"], BAD), ("Paid with another channel", now["duplicate"], AMBER),
           ("Pending notification", now["pending"], AMBER_L)]
    exc_total = sum(v for _, v, _ in exc)
    late = [(engine.BANKS[c].FULL_NAME, t["sla"]) for c, t in now_by.items() if t is not None and t["sla"]]
    late.sort(key=lambda x: -x[1])
    palette = [BAD, TEAL, ACC, AMBER]
    late_parts = [(n, v, palette[i % len(palette)]) for i, (n, v) in enumerate(late)]
    # one bank in the view: its share of the late payments would always be 100 %, so the card shows its payments
    # on time vs late instead
    with_report = [c for c, t in now_by.items() if t is not None]
    one_bank = engine.BANKS[with_report[0]].FULL_NAME if len(with_report) == 1 and with_report[0] in engine.BANKS else None
    if one_bank:
        late_title, late_unit = f"On time vs late · {one_bank}", "payments"
        late_parts = [("On time", now["processed"] - now["sla"], TEAL), ("Late", now["sla"], BAD)]
        late_empty = '<div class="dv-empty">No transactions in this period.</div>'
    else:
        late_title, late_unit = "Late payments by bank", "late"
        late_empty = '<div class="dv-empty">No late payment. ✓</div>'
    weekend = [(engine.BANKS[c].NAME, t["weekend"]) for c, t in now_by.items()
               if t is not None and t["weekend_known"] and t["weekend"]]
    weekend_note = "".join(f'<div class="dv-mut dv-wk">{v:,} of {escape(n)}\'s late payments were made Fri–Sun and '
                           f'booked Monday</div>' for n, v in weekend)

    # ---- needs attention
    tag = {"bad": ("Off target", "r"), "warn": ("Watch", "a"), "info": ("No report", "n")}
    att = d.attention(now_by, bef_by, span)
    att_html = "".join(f'<div class="dv-row"><span class="dv-ri">!</span><div><div class="dv-rn">'
                       f'{escape(a["title"].split(" · ")[0])}</div><div class="dv-rs">'
                       f'{escape(a["title"].split(" · ", 1)[1])}</div></div>{pill(*tag[a["kind"]])}</div>'
                       for a in att[:4]) or '<div class="dv-empty">Every bank is on target. ✓</div>'

    hero = f'{STYLE}<div class="dv"><div class="dv-hero">{head}</div></div>'
    return hero, f'''{STYLE}<div class="dv">
<div class="dv-grid">
 {card_html}
 <div class="dv-card"><div class="dv-ct"><h3>Payments per {unit}</h3><span class="dv-go">›</span></div>
  <div class="dv-sub">Faded: partial {unit} (some days not saved)</div>{week_bars}</div>
 <div class="dv-card dv-s2"><div class="dv-ct"><h3>Payments overview</h3><span class="dv-go">›</span></div>
  <div class="dv-sub">{unit_t}s up to {escape(span)} · {escape(scope.lower() if scope == "All banks" else scope)}</div>
  {area(series, TEAL, "dva", note) if series else '<div class="dv-empty">No period to show.</div>'}</div>
 <div class="dv-card dv-s2"><div class="dv-ct"><h3>Amount overview · {amt_cur}</h3><span class="dv-go">›</span></div>
  <div class="dv-sub">Amount collected per {unit} · {amt_cur} only (never added to another currency)</div>
  {area(aseries, TEAL, "dvb") if aseries else '<div class="dv-empty">No period to show.</div>'}</div>
 <div class="dv-card"><div class="dv-ct"><h3>Top merchants</h3><span class="dv-go">›</span></div>
  <div class="dv-sub">{merch_sub}</div>{merch_html}</div>
 <div class="dv-card dv-s2"><div class="dv-ct"><h3>Exceptions by type <small>· {escape(span)}</small></h3><span class="dv-go">›</span></div>
  <div class="dv-dwrap"><div>{legend(exc, "payments") or '<div class="dv-empty">No exception. ✓</div>'}</div>
  {donut(exc, f"{exc_total:,}", "exceptions")}</div></div>
 <div class="dv-card dv-s2"><div class="dv-ct"><h3>{escape(late_title)} <small>· {escape(span)}</small></h3><span class="dv-go">›</span></div>
  <div class="dv-dwrap"><div>{legend(late_parts, late_unit) or late_empty}{weekend_note}</div>
  {donut(late_parts, f"{now['sla']:,}", "late payments")}</div></div>
 <div class="dv-card"><div class="dv-ct"><h3>Needs attention</h3><span class="dv-go">›</span></div>{att_html}</div>
</div></div>'''


STYLE = f'''<style>
.dv{{font-family:"Nunito","Segoe UI",system-ui,sans-serif;color:{TEAL}}}
.dv-hero{{margin:0 0 16px}} .dv-hl{{font-size:14px;font-weight:800;color:{MUT}}}
.dv-hv{{font-size:46px;font-weight:800;letter-spacing:-1px;color:{TEAL};display:flex;align-items:center;gap:14px;line-height:1.15}}
.dv-hp{{font-size:14.5px;color:{MUT};margin-top:4px}} .dv-hp b{{color:{TEAL}}}
.dv-grid{{display:grid;grid-template-columns:repeat(4,minmax(0,1fr)) minmax(0,1.25fr);gap:16px}}
.dv-card{{background:#fff;border:1px solid #E1EEF0;border-radius:18px;padding:18px 18px;box-shadow:0 4px 16px rgba(18,50,57,.06);min-width:0}}
.dv-s2{{grid-column:span 2}}
.dv-k .dv-ic svg{{width:38px;height:38px}} .dv-kl{{font-size:13.5px;font-weight:800;color:{MUT};margin-top:10px}}
.dv-info{{cursor:help;font-weight:600;color:{ACC}}}
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
@media (max-width:1100px){{.dv-grid{{grid-template-columns:repeat(2,minmax(0,1fr))}} .dv-s2{{grid-column:span 2}}}}
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
