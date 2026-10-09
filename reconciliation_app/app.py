"""
BANK RECONCILIATION APP (IremboPay)
===================================
The screens of the app: a side panel + 6 steps. It contains no reconciliation rules: those are in engine.py
(common rules) and banks/<bank>.py (what is specific to one bank).

Start it with "Start Reconciliation.command" (double-click). The uploaded files are read in memory and never kept;
only the reports, the review decisions and the run log are saved (see engine.py).

How Streamlit works (to read this file): the whole file runs again from top to bottom after every click. What must
be remembered between clicks is kept in `state` (st.session_state, section 4). The page is drawn in the order of
the code: side panel, banner, then the screen of the current step.

The page
    ┌ side panel ──────────┐ ┌ banner: iremboPay | <Bank> Reconciliation ───────────────────────────┐
    │ Reports              │ │ 1. Period 2. Upload 3. Check 4. Reconcile 5. Exception Resolution 6. Report & Export │
    │ Exception decisions  │ │                                                                       │
    │ Recent runs          │ │ the screen of the current step (or a saved report, read only)        │
    │ Settings             │ │                                                                       │
    └──────────────────────┘ └───────────────────────────────────────────────────────────────────────┘

CONTENTS
    1.  Loading the engine and the banks
    2.  Settings
    3.  Page setup and style
    4.  Memory between clicks (session state)
    5.  Shared helpers (navigation, downloads, files, key figures, summary tables)
    6.  Side panel
    7.  Banner and step bar
    8.  A saved report (read only, opened from the side panel)
    9.  Step 1 - Period
    10. Step 2 - Upload
    11. Step 3 - Check
    12. Step 4 - Reconcile
    13. Step 5 - Exception Resolution
    14. Step 6 - Report & Export (quick export + custom report)
"""
import base64
import importlib
import os
import sys
from pathlib import Path

import pandas as pd
import streamlit as st

APP_DIR = Path(__file__).resolve().parent


# =====================================================================================================================
# 1. LOADING THE ENGINE AND THE BANKS
#    A running app keeps the code it started with; when engine.py or a bank file has changed, it is loaded again.
# =====================================================================================================================
def _code_mtime():
    """Last change time of engine.py and the bank files."""
    return max(p.stat().st_mtime for p in [APP_DIR / "engine.py", *APP_DIR.glob("banks/*.py")])


try:
    import engine
    # reload the engine and the bank files when they have changed (otherwise a running app keeps the old version)
    if getattr(engine, "_loaded_mtime", None) != _code_mtime():
        for _name in [n for n in sys.modules if n.startswith("banks.")] + ["banks"]:
            if _name in sys.modules:           # not there yet when the app was started before banks/ existed
                importlib.reload(sys.modules[_name])
        engine = importlib.reload(engine)
        engine._loaded_mtime = _code_mtime()
except ValueError as _err:                     # a bank file is incomplete
    st.error(str(_err), icon="⛔")
    st.stop()


# =====================================================================================================================
# 2. SETTINGS
# =====================================================================================================================
# the "Bank Reconciliation" folder (output/ is created inside it); IM_RECON_BASE is only used for testing
BASE_DIR = Path(os.environ.get("IM_RECON_BASE", Path(__file__).resolve().parent.parent))
BRAND = "#123239"          # IremboPay dark teal
ACCENT = "#99E378"         # IremboPay green
STEPS = ["Period", "Upload", "Check", "Reconcile", "Exception Resolution", "Report & Export"]
PREVIEW_TEAL = "#2B8A99"   # charts of the custom report preview: 4.0:1 on the light theme, 4.7:1 on the dark one
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

# icons of the side panel titles (simple line drawings, coloured by the CSS of section 3)
ICON_REPORTS = ('<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" '
                'stroke-linejoin="round"><path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z"/>'
                '<path d="M14 3v5h5"/><path d="M9 13h6M9 17h4"/></svg>')
ICON_REVIEW = ('<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" '
               'stroke-linejoin="round"><path d="M9 3h6v3H9z"/><path d="M15 4.5h2a2 2 0 0 1 2 2V19a2 2 0 0 1-2 2H7a2 2 0 0 1'
               '-2-2V6.5a2 2 0 0 1 2-2h2"/><path d="m9 14 2 2 4-4"/></svg>')


# =====================================================================================================================
# 3. PAGE SETUP AND STYLE
#    The bank chosen in step 1 is needed first: its name is the page title. Then the CSS (IremboPay colours).
#    The Nunito font is set for the whole app in .streamlit/config.toml.
# =====================================================================================================================
state = st.session_state
if state.get("bank_code") not in engine.BANKS:
    state.bank_code = engine.DEFAULT_BANK
bank = engine.get_bank(state.bank_code)
MANY_BANKS = len(engine.BANKS) > 1

st.set_page_config(page_title=f"{bank.FULL_NAME} Reconciliation", page_icon=str(APP_DIR / "assets" / "irembopay_icon.png"), layout="wide",
                   initial_sidebar_state="expanded")
st.markdown(f"""
<style>
  /* page width, and room at the top for the banner */
  .block-container {{max-width: 1180px; padding-top: 3.6rem; padding-bottom: 3rem;}}
  header[data-testid="stHeader"] {{background: transparent; height: 3rem;}}
  /* banner of the steps: logo | <Bank> Reconciliation */
  .app-banner {{background: {BRAND}; border-left: 6px solid {ACCENT}; color: #fff;
               padding: 1.1rem 1.4rem; border-radius: 14px; margin-bottom: .8rem;}}
  .app-banner .brand {{display: flex; align-items: center; gap: .9rem;}}
  .app-banner .brand img {{height: 1.5rem; padding-right: .9rem; border-right: 1px solid rgba(214,228,230,.35);}}
  .app-banner h1 {{color: {ACCENT}; font-size: 1.55rem; margin: 0; padding: 0; line-height: 1.3;}}
  .app-banner p {{color: #D6E4E6; margin: .25rem 0 0 0; font-size: .95rem;}}
  /* banner of a saved report (green, so it looks different from the steps) */
  .view-banner {{background: {ACCENT}; color: {BRAND}; padding: 1rem 1.4rem; border-radius: 14px;
                margin-bottom: .8rem; border: 2px solid {BRAND};}}
  .view-banner .tag {{display: inline-block; background: {BRAND}; color: {ACCENT}; font-size: .78rem;
                     font-weight: 700; letter-spacing: .06em; padding: .15rem .6rem; border-radius: 999px;}}
  .view-banner h2 {{color: {BRAND}; font-size: 1.45rem; margin: .35rem 0 0 0; padding: 0;}}
  .view-banner p {{color: {BRAND}; margin: .2rem 0 0 0; font-size: .93rem; opacity: .85;}}
  /* step bar: done = light green, current = teal */
  .steps {{display: flex; gap: .4rem; flex-wrap: wrap; margin: .2rem 0 1rem 0;}}
  .step {{padding: .3rem .8rem; border-radius: 999px; font-size: .88rem;
          border: 1px solid rgba(128,128,128,.35); opacity: .6;}}
  .step.done {{background: rgba(46,160,67,.15); border-color: rgba(46,160,67,.6); opacity: 1;}}
  .step.now {{background: {BRAND}; border-color: {BRAND}; color: {ACCENT}; opacity: 1; font-weight: 650;}}
  /* card titles and small grey text */
  .card-title {{font-weight: 650; font-size: 1.12rem; margin-bottom: .3rem;}}
  .muted {{opacity: .7; font-size: .9rem;}}
  div[data-testid="stMetricValue"] {{font-size: 1.5rem;}}
  /* side panel: IremboPay colours (the Nunito font is set for the whole app in .streamlit/config.toml) */
  .side-head {{display: flex; align-items: center; gap: .55rem; margin: .1rem 0 .55rem 0;
               font-family: "Nunito", sans-serif; font-weight: 800; font-size: 1.08rem; color: {BRAND};}}
  .side-icon {{display: inline-flex; align-items: center; justify-content: center; width: 1.75rem; height: 1.75rem;
               border-radius: .5rem; background: {BRAND}; color: {ACCENT}; flex: none;}}
  .side-icon svg {{width: 1rem; height: 1rem;}}
  .side-count {{margin-left: auto; background: {ACCENT}; color: {BRAND}; font-weight: 800; font-size: .74rem;
                padding: .12rem .6rem; border-radius: 999px; white-space: nowrap;}}
  .side-empty {{font-family: "Nunito", sans-serif; color: #5B6B70; font-size: .9rem; margin: 0 0 .4rem .2rem;}}
  section[data-testid="stSidebar"] [data-testid="stExpander"] summary p {{
      font-family: "Nunito", sans-serif; font-weight: 800; color: {BRAND};}}
  section[data-testid="stSidebar"] [data-testid="stExpander"] summary [data-testid="stIconMaterial"] {{color: {BRAND};}}
  /* dropdown lists: exactly 2 choices shown (40px each); more than 2 -> scroll to see the rest */
  [data-testid="stSelectboxVirtualDropdown"] > div {{max-height: 80px !important; overflow-y: auto !important;}}
  section[data-testid="stSidebar"] [data-testid="stSelectbox"] input {{cursor: pointer; caret-color: transparent;}}
  .side-note {{font-size: .82rem; color: #5B6B70; margin: -.35rem 0 .45rem .2rem;}}
</style>""", unsafe_allow_html=True)


# =====================================================================================================================
# 4. MEMORY BETWEEN CLICKS (session state)
#    step            current step (1 to 6)
#    bank_code       bank chosen in step 1           report_label / picked / compare   the other step-1 choices
#    files           uploaded files, read in memory  read_errors                        files that could not be read
#    uploader_round  changed to empty the upload boxes after each upload
#    result          the last engine.RunResult       run_error                          last error, shown in red
#    view_report     path of the saved report open from the side panel (None = the steps are shown)
#    confirm         which delete is waiting for "Yes, delete"
#    pending         values set by a button, applied here before the widgets are drawn
#    toast           a short message shown once after the next reload
# =====================================================================================================================
# keep the step-1 choices when other steps are shown (Streamlit forgets widget values that are not on screen)
for _k in ("report_label", "picked", "compare", "bank_code"):
    if _k in state:
        state[_k] = state[_k]
if "pending" in state:                     # changes asked by a button (applied before the widgets are drawn)
    for _k, _v in state.pop("pending").items():
        state[_k] = _v
state.setdefault("step", 1)
state.setdefault("files", {})            # digest -> engine.UploadedFile (in memory only)
state.setdefault("read_errors", [])      # files that could not be read
state.setdefault("uploader_round", 0)
state.setdefault("result", None)
state.setdefault("run_error", None)
state.setdefault("compare", False)
if "report_label" not in state:
    state.report_label = "Month"
if "picked" not in state:                # default: last month
    state.picked = (pd.Timestamp.today().normalize().replace(day=1) - pd.offsets.MonthBegin(1)).date()


if state.get("toast"):
    st.toast(state.pop("toast"), icon="📄")


# =====================================================================================================================
# 5. SHARED HELPERS (used by several steps and by the side panel)
# =====================================================================================================================
# ---- navigation
def go(step):
    """Show another step."""
    state.step = step
    st.rerun()


# ---- download as Excel or PDF
@st.cache_data(show_spinner=False, max_entries=20)
def pdf_bytes(path: str, mtime: float) -> bytes:          # mtime: a new PDF is made when the report changes
    """The PDF of a report (made once, then kept in memory)."""
    return engine.report_pdf(Path(path))


def download_control(path: Path, key: str, label="⬇ Download", primary=False, fmt_label="collapsed"):
    """Format choice (Excel / PDF) + one download button."""
    fmt = st.segmented_control("Format", ["Excel", "PDF"], default="Excel", key=f"fmt_{key}",
                               label_visibility=fmt_label) or "Excel"
    if fmt == "PDF":
        try:
            data, name, mime = pdf_bytes(str(path), path.stat().st_mtime), path.with_suffix(".pdf").name, "application/pdf"
        except Exception as err:
            st.error(f"The PDF could not be made: {err}", icon="⚠️")
            return
    else:
        data, name, mime = path.read_bytes(), path.name, XLSX_MIME
    st.download_button(f"{label} ({fmt})", data=data, file_name=name, mime=mime, width="stretch",
                       type="primary" if primary else "secondary", key=f"dl_{key}")


# ---- side panel titles
def side_head(icon, title, count=None):
    """Side panel section title in the IremboPay colours, with an optional count badge."""
    badge = f'<span class="side-count">{count}</span>' if count else ""
    st.markdown(f'<div class="side-head"><span class="side-icon">{icon}</span><span>{title}</span>{badge}</div>',
                unsafe_allow_html=True)


# ---- uploaded files
def accepted_files():
    """Files accepted for the current period (duplicates and files outside the period are left out)."""
    acc, rej = [], []
    for f in state.files.values():
        reason = engine.accept_file(period, f, acc, bank)
        if reason:
            rej.append(reason)
        else:
            acc.append(f)
    return acc, rej


# ---- key figures + summary tables (step 5 and a saved report)
def _count(tables, title, cur=None):
    """(transactions, amount, %) of one table: the total row, or the row of one currency."""
    t = tables[title]
    row = t[t["Currency"].isna()] if cur is None else t[t["Currency"].eq(cur)]
    return int(row["Number of Transactions"].iloc[0]), row["Total Amount"].iloc[0], (row["%"].iloc[0] if "%" in t else None)


def render_kpis(tables, not_found):
    """The 6 key figures at the top: processed, received %, SLA violated, pending, duplicate, not found."""
    processed, _, _ = _count(tables, "Overall processed payments")
    rec_n, _, rec_pct = _count(tables, "Overall notifications received by Irembo")
    sla_n, _, sla_pct = _count(tables, "SLA Violated")
    t = tables["Overall processed payments"]
    per_currency = " · ".join(f"{cur} {int(n):,} ({0 if pd.isna(amt) else amt:,.0f})"
                              for cur, n, amt in zip(t["Currency"], t["Number of Transactions"], t["Total Amount"])
                              if pd.notna(cur))
    k = st.columns(6, border=True)
    k[0].metric("Processed", f"{processed:,}", help=per_currency)
    k[1].metric("Received", f"{rec_pct:.2%}", help=f"{rec_n:,} notifications received by Irembo")
    k[2].metric("SLA violated", f"{sla_n:,}", help=f"{sla_pct:.2%} of processed")
    k[3].metric("Pending (NEW)", f"{_count(tables, 'Pending notifications')[0]:,}")
    k[4].metric("Duplicate", f"{_count(tables, 'Duplicate Payment')[0]:,}", help="Paid with another channel")
    k[5].metric("Not found", f"{not_found:,}", help="Bank rows whose invoice was not in the payment file")


def _show_tables(tbls):
    """The 5 summary tables of one block, two by two."""
    for pair in [["Overall processed payments", "Overall notifications received by Irembo"],
                 ["SLA Violated", "Pending notifications"], ["Duplicate Payment"]]:
        cols = st.columns(2)
        for title, col in zip(pair, cols):
            t = tbls[title].copy()
            t["Currency"] = t["Currency"].fillna("Total")
            t["Total Amount"] = ["" if pd.isna(v) else f"{v:,.0f}" for v in t["Total Amount"]]   # Total row: blank
            config = {"Number of Transactions": st.column_config.NumberColumn("Transactions", format="localized"),
                      "Total Amount": st.column_config.TextColumn("Total amount", alignment="right")}
            if "%" in t:
                t["%"] = t["%"] * 100
                config["%"] = st.column_config.ProgressColumn("% of processed", format="%.2f%%", min_value=0, max_value=100)
            with col.container(border=True):
                st.markdown(f"**{title}**")
                st.dataframe(t, hide_index=True, width="stretch", column_config=config)


def render_summary(blocks, report_type):
    """The summary: one tab for the month + one per week (a day or a week: a single summary)."""
    main = next((b for b in blocks if b[0] in ("Summary-Monthly", "Summary-Daily")), blocks[0])
    week_blocks = [b for b in blocks if b[0] == "Summary-Weekly"]
    st.markdown("#### Summary")
    if report_type == "month" and week_blocks:
        month_name = main[1].split(":")[1].split("(")[0].strip().split(" ")[0] if ":" in main[1] else "Month"
        tabs = st.tabs([f"📅 {month_name}"] + [h.split(":")[0] for _, h, _ in week_blocks])
        with tabs[0]:
            _show_tables(main[2])
        for tab, (_, heading, tbls) in zip(tabs[1:], week_blocks):
            with tab:
                st.caption(heading)
                _show_tables(tbls)
    else:
        st.caption(main[1])
        _show_tables(main[2])


# ---- short figures (step 4) and the result of the exception resolution (steps 5 and 6)
def short_amount(value):
    """68,166,357 -> '68.2M', 470,064 -> '470K', 950 -> '950'."""
    if abs(value) >= 1e9:
        return f"{value / 1e9:.1f}B"
    if abs(value) >= 1e6:
        return f"{value / 1e6:.1f}M"
    if abs(value) >= 1e3:
        return f"{value / 1e3:.0f}K"
    return f"{value:,.0f}"


def result_line(counts: dict) -> str:
    """{'Pending notification': 6, 'Paid with other channel': 9} -> '9 Paid with other channel · 6 Pending
    notification' (largest first)."""
    return " · ".join(f"{n:,} {what}" for what, n in sorted(counts.items(), key=lambda x: -x[1]) if n)


# ---- custom report (step 6): the choices are widgets with keys "cr_..."; a report type fills them in
def apply_report_type():
    """A report type was chosen: tick its sections and set how each one is shown."""
    preset = engine.REPORT_TYPES[state.get("cr_type") or "Full Report"][1]
    for key, (_, _, views) in engine.CUSTOM_SECTIONS.items():
        state[f"cr_sec_{key}"] = key in preset
        state[f"cr_as_{key}"] = preset.get(key, views[0])


def tick_all_sections(value: bool):
    """Select all / Clear all."""
    for key in engine.CUSTOM_SECTIONS:
        state[f"cr_sec_{key}"] = value


def custom_options(groups: list) -> dict:
    """The choices made on screen, in the form engine.custom_content() reads."""
    sections = [k for k in engine.CUSTOM_SECTIONS if state.get(f"cr_sec_{k}")]
    views = {k: {"as": state.get(f"cr_as_{k}") or engine.CUSTOM_SECTIONS[k][2][0],
                 "group": state.get(f"cr_by_{k}") or groups[0], "currency": state.get(f"cr_cur_{k}") or "All"}
             for k in sections}
    return {"type": state.get("cr_type") or "Full Report", "sections": sections, "views": views,
            "orientation": state.get("cr_orient") or "Landscape", "density": state.get("cr_density") or "Standard",
            "include": {k: state.get(f"cr_inc_{k}", True) for k in engine.CUSTOM_INCLUDE}}


def _chart(series: list, group: str, kind: str):
    """One-series bar or line chart of the preview (no legend: the section title names it; hover shows the value)."""
    import altair as alt
    df = pd.DataFrame(series, columns=[group, "Transactions"])
    base = alt.Chart(df).encode(
        x=alt.X(f"{group}:N", sort=None, title=None, axis=alt.Axis(labelAngle=0 if len(df) <= 12 else -90)),
        y=alt.Y("Transactions:Q", title=None, axis=alt.Axis(format=",d", tickMinStep=1)),
        tooltip=[alt.Tooltip(f"{group}:N"), alt.Tooltip("Transactions:Q", format=",")])
    if kind == engine.BAR:
        mark = base.mark_bar(color=PREVIEW_TEAL, cornerRadiusTopLeft=4, cornerRadiusTopRight=4)
    else:
        mark = base.mark_line(color=PREVIEW_TEAL, strokeWidth=2,
                              point=alt.OverlayMarkDef(color=PREVIEW_TEAL, size=64, filled=True))
    return mark.properties(height=230)


def _as_text(df: pd.DataFrame) -> pd.DataFrame:
    """A table with every value written as in the PDF (amounts, %, dates)."""
    return pd.DataFrame({c: [engine.cell_text(c, v) for v in df[c]] for c in df.columns})


def render_custom_preview(content: dict):
    """The preview of a custom report: the same content the PDF and Excel files are made from."""
    with st.container(border=True):
        st.caption("REPORT PREVIEW")
        st.markdown(f"#### {content['title']}")
        st.caption(" · ".join(x for x in (content["period"], content["type"], content["stamp"],
                                          f"{content['orientation']}, {content['density'].lower()}") if x))
        if content["kpis"]:
            for col, (label, value) in zip(st.columns(len(content["kpis"]), border=True), content["kpis"]):
                col.metric(label, value)
        lines = content["currency_lines"] + ([content["exceptions_line"]] if content["exceptions_line"] else [])
        if lines:
            st.markdown("  \n".join(lines))
        if not content["sections"] and not content["kpis"]:
            st.caption("No section selected.")
        def invoice_list(sec, named=False):
            if sec["list"] is not None:
                st.caption((f"{sec['title']} · " if named else "") + f"{sec['list_title']} ({len(sec['list']):,})")
                if len(sec["list"]):
                    st.dataframe(_as_text(sec["list"]), hide_index=True, width="stretch",
                                 height=min(38 + 35 * len(sec["list"]), 320))

        # charts side by side, two per row, each in its own card; tables and lists take the whole width
        for kind, item in engine.layout_rows(content["sections"]):
            if kind == "section":
                st.divider()
                st.markdown(f"**{item['title'].upper()}**")
                if item["table"] is not None:
                    st.dataframe(_as_text(item["table"]), hide_index=True)
                invoice_list(item)
                continue
            for col, sec in zip(st.columns(2) if len(item) == 2 else [st], item):
                with col.container(border=True):
                    icon = ":material/bar_chart:" if sec["as"] == engine.BAR else ":material/show_chart:"
                    st.markdown(f"{icon} **{sec['title']}**")
                    st.caption(engine.series_label(sec))
                    st.altair_chart(_chart(sec["series"], sec["group"], sec["as"]), width="stretch")
            for sec in item:
                invoice_list(sec, named=True)


# =====================================================================================================================
# 6. SIDE PANEL (only the bank chosen in step 1)
#    Reports: view, download (Excel or PDF), delete   ·   Exception decisions: delete   ·   Recent runs   ·   Settings
# =====================================================================================================================
# side panel dropdowns: choose only, no typing (the box is made read-only; clicking it still opens the list)
st.html("""<script>
(function () {
  const lock = () => document.querySelectorAll('section[data-testid="stSidebar"] [data-testid="stSelectbox"] input')
    .forEach(i => { if (!i.readOnly) i.readOnly = true; });
  lock();
  if (!window.__sideLock) {
    window.__sideLock = new MutationObserver(lock);
    window.__sideLock.observe(document.body, {childList: true, subtree: true});
  }
})();
</script>""", unsafe_allow_javascript=True)

with st.sidebar:
    # ---- 6a. reports: pick one, then view, download or delete ----
    reports = engine.report_list(BASE_DIR, bank)                 # only the bank chosen in step 1
    side_head(ICON_REPORTS, "Reports", f"{len(reports)} saved" if reports else None)
    if not reports:
        st.markdown('<div class="side-empty">No reports yet.</div>', unsafe_allow_html=True)
    else:
        titles = [t for t, _ in reports]
        chosen = st.selectbox("Report", titles, key="side_report", label_visibility="collapsed",
                              filter_mode=None)
        path = dict(reports)[chosen]
        if state.get("view_report") and state.view_report != str(path):
            state.view_report = str(path)          # a report is open: show the newly chosen one at once
            state.confirm = None
            st.toast(f"Opened {chosen}", icon="📄")
        if state.get("view_report") == str(path):
            if st.button("✖ Close view", width="stretch", key="side_close_view"):
                state.view_report = None
                state.toast = f"Closed {chosen}"
                st.rerun()
        elif st.button("👁 View", type="primary", width="stretch", key="side_view"):
            state.view_report = str(path)
            state.confirm = None
            state.toast = f"Opened {chosen}"
            st.rerun()
        download_control(path, "side")
        # delete: asks "Yes, delete / Cancel" inside the side panel first
        if st.button("🗑 Delete", width="stretch", key="side_ask_report"):
            state.confirm = ("report", chosen)
        if state.get("confirm") == ("report", chosen):
            with st.container(border=True):
                st.markdown(f"**Delete this report?**  \n{chosen}")
                st.caption("The Excel file is removed. Its exception decisions are kept.")
                y, n = st.columns(2)
                if y.button("Yes, delete", type="primary", width="stretch", key="side_del_report"):
                    try:
                        engine.delete_report(BASE_DIR, path)
                        if state.get("view_report") == str(path):
                            state.view_report = None
                        if state.result is not None and state.result.report_file == path:
                            state.result = None
                            state.step = min(state.step, 3)
                        st.toast(f"Report deleted: {chosen}", icon="🗑️")
                    except engine.ReconError as err:
                        st.toast(str(err), icon="⚠️")
                    state.confirm = None
                    st.rerun()
                if n.button("Cancel", width="stretch", key="side_cancel_report"):
                    state.confirm = None
                    st.rerun()

    st.divider()

    # ---- 6b. exception decisions (the review store): pick a period, then delete ----
    reviews = engine.review_list(BASE_DIR, bank)
    side_head(ICON_REVIEW, "Exception decisions", f"{len(reviews)} saved" if reviews else None)
    if not reviews:
        st.markdown('<div class="side-empty">No exception decisions yet.</div>', unsafe_allow_html=True)
    else:
        labels = {title: (title, code, key, n) for title, code, key, n, _ in reviews}
        chosen_r = st.selectbox("Review", list(labels), key="side_review", label_visibility="collapsed",
                                filter_mode=None)
        title_r, code_r, key_r, n_r = labels[chosen_r]
        st.markdown(f'<div class="side-note">{n_r} invoice(s) decided</div>', unsafe_allow_html=True)
        if st.button("🗑 Delete these decisions", width="stretch", key="side_ask_review"):
            state.confirm = ("review", code_r, key_r)
        if state.get("confirm") == ("review", code_r, key_r):
            with st.container(border=True):
                st.markdown(f"**Delete these decisions?**  \n{title_r}")
                st.caption("These invoices will need to be resolved again. Reports are not changed.")
                y, n = st.columns(2)
                if y.button("Yes, delete", type="primary", width="stretch", key="side_del_review"):
                    try:
                        engine.delete_review(BASE_DIR, key_r, code_r)
                        if (state.result is not None and state.result.bank == code_r
                                and state.result.period.store_key == key_r):
                            state.result = None
                            state.step = min(state.step, 3)
                        st.toast(f"Exception decisions deleted: {title_r}", icon="🗑️")
                    except engine.ReconError as err:
                        st.toast(str(err), icon="⚠️")
                    state.confirm = None
                    st.rerun()
                if n.button("Cancel", width="stretch", key="side_cancel_review"):
                    state.confirm = None
                    st.rerun()

    st.divider()

    # ---- 6c. recent runs (short lines) + 6d. settings ----
    with st.expander("Recent runs", icon=":material/history:"):
        runs = engine.recent_runs(BASE_DIR, n=6, bank=bank)
        if runs.empty:
            st.caption("No runs yet.")
        for _, r in runs.iterrows():
            when = pd.to_datetime(r["run_at"], errors="coerce")
            icon = "✅" if r["result"] == "report made" else "⏸️"
            label = r["period"].replace("/", "-").replace(" - ", "_to_")
            st.caption(f"{icon} {when:%d/%m %H:%M} · {engine.period_title(r['report_type'], label)}")
    with st.expander("Settings", icon=":material/tune:"):
        if engine.option(bank, "SLA_DAYS_SETTING", True):
            sla_days = st.number_input(f"SLA: days allowed between bank {bank.COLUMNS['payment_date']} and paid_at",
                                       0, 10, 0,
                                       help=f"0 = paid_at must be on the same day as the bank {bank.COLUMNS['payment_date']}")
        else:                                  # a bank with a fixed SLA rule (e.g. Ecobank: same minute)
            sla_days = 0
            st.caption(f"SLA ({bank.NAME}): the bank {bank.COLUMNS['payment_date']} and paid_at must be in the "
                       "same minute.")


# =====================================================================================================================
# 7. BANNER AND STEP BAR (hidden while a saved report is open)
# =====================================================================================================================
if not state.get("view_report"):
    _h = lambda s: s.replace("&", "&amp;")
    _logo = base64.b64encode((APP_DIR / "assets" / "irembopay_wordmark.png").read_bytes()).decode()
    st.markdown(f"""<div class="app-banner"><div class="brand"><img src="data:image/png;base64,{_logo}" alt="IremboPay">
<h1>{_h(bank.FULL_NAME)} Reconciliation</h1></div></div>""",
                unsafe_allow_html=True)
if not state.get("view_report"):
    st.markdown('<div class="steps">' + "".join(
        f'<span class="step {"done" if i + 1 < state.step else "now" if i + 1 == state.step else ""}">'
        f'{"✓ " if i + 1 < state.step else ""}{i + 1}. {name}</span>' for i, name in enumerate(STEPS)) + "</div>",
                unsafe_allow_html=True)

# the period chosen in step 1 (used by every step)
period = engine.make_period((state.report_label or "Month").lower(), state.picked)
compare = state.compare


# =====================================================================================================================
# 8. A SAVED REPORT (read only, opened with "View" in the side panel; the steps are hidden meanwhile)
# =====================================================================================================================
if state.get("view_report"):
    vpath = Path(state.view_report)
    try:
        if not vpath.exists():
            raise engine.ReconError("This report no longer exists.")
        rep = engine.read_report(vpath)
    except engine.ReconError as err:
        state.view_report = None
        st.error(str(err), icon="⚠️")
        st.stop()
    vmain = next((b for b in rep["blocks"] if b[0] in ("Summary-Monthly", "Summary-Daily")), rep["blocks"][0])
    created = pd.Timestamp(vpath.stat().st_mtime, unit="s")
    when = (f"Reconciled on {rep['reconciled_on']}" if rep["reconciled_on"]          # older reports: the file time
            else f"Made on {created:%d/%m/%Y at %H:%M}")
    st.markdown(f"""<div class="view-banner"><span class="tag">SAVED REPORT · READ ONLY</span>
<h2>📄 {rep["bank"].replace("&", "&amp;")} · {rep["title"]}</h2><p>{when}</p></div>""",
                unsafe_allow_html=True)
    if st.button("←  Back", type="primary",
                 key="view_back"):
        state.view_report = None
        state.toast = f"Closed {rep['title']}"
        st.rerun()
    with st.container(border=True):
        render_kpis(vmain[2], rep["not_found"])
        v1, v2 = st.columns([3, 1], vertical_alignment="bottom")
        with v1:
            download_control(vpath, "view", label="⬇ Download report", primary=True)
        if v2.button("✖  Close", width="stretch", key="view_close"):
            state.view_report = None
            state.toast = f"Closed {rep['title']}"
            st.rerun()
    render_summary(rep["blocks"], rep["report_type"])
    st.stop()                                   # nothing else is drawn while a saved report is open


# =====================================================================================================================
# 9. STEP 1 - PERIOD: the bank, day / week / month, the date, and the optional comparison
# =====================================================================================================================
if state.step == 1:
    with st.container(border=True):
        st.markdown('<div class="card-title">① Choose the period</div>', unsafe_allow_html=True)
        if MANY_BANKS:
            st.segmented_control("Bank", list(engine.BANKS), key="bank_code", selection_mode="single",
                                 format_func=engine.bank_name)
        else:
            st.markdown(f"**Bank:** {bank.FULL_NAME}")
        if state.result is not None and state.result.bank != bank.CODE:     # bank changed: old result no longer valid
            state.result = None
        c1, c2, c3 = st.columns([1.3, 1, 1.7], vertical_alignment="bottom")
        c1.segmented_control("Report type", ["Day", "Week", "Month"], key="report_label", selection_mode="single")
        c2.date_input("Date", key="picked", format="DD/MM/YYYY", help="Any day inside the day / week / month you want")
        period = engine.make_period((state.report_label or "Month").lower(), state.picked)
        c3.markdown(f"**{period.name[0].upper() + period.name[1:]}**  \n<span class='muted'>"
                    f"{period.start:%d/%m/%Y} to {period.last_day:%d/%m/%Y}</span>", unsafe_allow_html=True)

        # ---- only a period that has ended can be reconciled (a day at midnight, a week on Monday, a month on the 1st)
        not_ended = engine.period_not_ended(period)
        if not_ended:
            st.error(not_ended.replace("\n", "  \n"), icon="⛔")

        # ---- comparison: the previous report of this bank must exist (or be uploaded, or made first)
        prev_path, prev = engine.find_previous_report(BASE_DIR, period, bank)
        st.checkbox(f"Compare with previous period ({prev.name})", key="compare",
                    help="Adds a 'VS prev' column to each table. When not ticked, the report is unchanged.")
        ready = not not_ended
        if state.compare and not not_ended:             # a blocked period: nothing to compare yet
            if prev_path is not None:
                created = pd.Timestamp(prev_path.stat().st_mtime, unit="s")
                st.success(f"Previous report found: {prev.name} (created {created:%d/%m/%Y})", icon="✅")
            else:
                ready = False
                st.warning(f"No {bank.NAME} report found for {prev.name}", icon="⚠️")
                u1, u2 = st.columns([2, 1], vertical_alignment="bottom")
                up = u1.file_uploader(f"Upload the {prev.name} report (.xlsx)", type=["xlsx"],
                                      key=f"prev_report_{state.uploader_round}")
                if up is not None:
                    try:
                        engine.save_uploaded_report(BASE_DIR, prev, up.name, up.getvalue(), bank)
                        st.toast(f"Report for {prev.name} added", icon="✅")
                    except engine.ReconError as err:
                        state.run_error = str(err)
                    state.uploader_round += 1
                    st.rerun()
                if u2.button(f"Make the {prev.name} report first", width="stretch"):
                    state.pending = {"picked": prev.start.date(), "compare": False}
                    st.rerun()
                if state.run_error:
                    st.error(state.run_error, icon="⚠️")
    c1, _, c3 = st.columns([1, 2, 1.3])
    if c3.button("Next: upload files →", type="primary", width="stretch", disabled=not ready,
                 help=None if ready else "This period has not ended yet" if not_ended
                 else "Add the previous report first, or untick the comparison"):
        state.run_error = None
        go(2)

# =====================================================================================================================
# 10. STEP 2 - UPLOAD: the bank statements + the IremboPay payment file (read in memory, never saved)
# =====================================================================================================================
elif state.step == 2:
    acc, rej = accepted_files()
    with st.container(border=True):
        st.markdown(f'<div class="card-title">② Upload the files for {period.name}</div>', unsafe_allow_html=True)
        # ---- what to upload (depends on the bank)
        monday = engine.extra_monday(period, bank)
        need_to = monday or period.last_day
        day_before = " and ".join(bank.DAY_BEFORE_CURRENCIES)
        per_currency = engine.option(bank, "ONE_FILE_PER_CURRENCY", True)
        pay_types = engine.option(bank, "PAYMENT_FILE_TYPES", ["csv"])
        pay_label = " or ".join(f".{t}" for t in pay_types)
        st.markdown(
            f"- **{bank.NAME} bank statements**"
            + (f", {' and '.join(bank.CURRENCIES)}" if per_currency else "")
            + f", covering **{period.start:%d/%m} to {need_to:%d/%m/%Y}**"
            + (f" (including **Monday {monday:%d/%m}**: late weekend payments are posted on Monday)"
               if monday is not None else "")
            + (f"\n- Recommended: the {day_before} statement of **{period.day_before:%d/%m/%Y}** (the day before)"
               if day_before else "")
            + ("\n- Also the statements of the **days before**, if some payments of this period were made then "
               "(step ③ tells you which days)" if engine.option(bank, "ACCEPT_EARLIER_DAYS", 0) else "")
            + f"\n- The **IremboPay payment file** ({pay_label}) for the period")

        # ---- the two upload boxes; each new file is read at once and listed below
        u1, u2 = st.columns(2)
        types = ", ".join(f".{t}" for t in bank.FILE_TYPES)
        bank_up = u1.file_uploader(f"{bank.NAME} bank files ({types})", type=bank.FILE_TYPES,
                                   accept_multiple_files=True, key=f"bank_{state.uploader_round}")
        pay_up = u2.file_uploader(f"IremboPay payment file ({pay_label})", type=pay_types,
                                  accept_multiple_files=True, key=f"pay_{state.uploader_round}")
        new = [(u, "bank") for u in bank_up or []] + [(u, "payments") for u in pay_up or []]
        if new:
            with st.spinner("Reading the files…"):
                for up, kind in new:
                    content = up.getvalue()
                    digest = engine.file_digest(content)
                    if digest in state.files:
                        state.read_errors.append(f"{up.name}: the same file was already added.")
                        continue
                    try:
                        if kind == "bank":
                            state.files[digest] = engine.read_bank_file(up.name, content, bank)
                        else:
                            state.files[digest] = engine.read_payment_file(up.name, content, bank)
                    except engine.ReconError as err:
                        state.read_errors.append(str(err))
                    except Exception as err:
                        state.read_errors.append(f"{up.name}: could not be read ({err}).")
            state.uploader_round += 1          # empty the upload boxes; the files are listed below
            state.result = None
            st.rerun()

        # ---- one summary line ----
        if per_currency:
            n_cur = {f"Bank {cur}": sum(f.kind == "bank" and f.currency == cur for f in acc) for cur in bank.CURRENCIES}
        else:
            n_cur = {"Bank statements": sum(f.kind == "bank" for f in acc)}
        n_pay = sum(f.kind == "payments" for f in acc)
        if acc:
            line = (f"**{len(acc)} file(s) added** · "
                    + " · ".join(f"{label}: **{n}**" for label, n in n_cur.items())
                    + f" · Payment file: **{n_pay}**")
            if all(n_cur.values()) and n_pay:
                st.success(line, icon="✅")
            else:
                st.warning(line, icon="⚠️")

        # ---- added files: one compact table (select rows to remove them) ----
        if acc:
            ordered = sorted(acc, key=lambda x: (x.kind, x.currency, x.first))
            table = pd.DataFrame({
                "File": [f.name for f in ordered],
                "Type": ["Bank" if f.kind == "bank" else "Payments" for f in ordered],
                "Currency": [f.currency or "–" for f in ordered],
                "From": [f.first.date() for f in ordered],
                "To": [f.last.date() for f in ordered],
                "Rows": [f.rows for f in ordered]})
            event = st.dataframe(
                table, hide_index=True, width="stretch", height=min(38 + 35 * len(table), 300),
                on_select="rerun", selection_mode="multi-row", key=f"files_{state.uploader_round}",
                column_config={"File": st.column_config.TextColumn("File", width="large"),
                               "From": st.column_config.DateColumn("From", format="DD/MM/YYYY"),
                               "To": st.column_config.DateColumn("To", format="DD/MM/YYYY"),
                               "Rows": st.column_config.NumberColumn("Rows", format="localized")})
            picked_rows = event.selection.rows
            if picked_rows and st.button(f"🗑 Remove {len(picked_rows)} selected file(s)"):
                for i in picked_rows:
                    state.files.pop(ordered[i].digest, None)
                state.result = None
                state.uploader_round += 1
                st.rerun()
            elif not picked_rows:
                st.caption("Tick a row on the left to remove a file.")

        # ---- ignored files: folded away ----
        ignored = [m.split(": ", 1) if ": " in m else [m, ""] for m in state.read_errors + rej]
        if ignored:
            with st.expander(f"⛔ {len(ignored)} file(s) ignored", expanded=False):
                st.dataframe(pd.DataFrame(ignored, columns=["File", "Why"]), hide_index=True, width="stretch",
                             column_config={"Why": st.column_config.TextColumn("Why", width="large")})
                if st.button(f"🗑 Remove all {len(ignored)} ignored file(s)"):
                    accepted_ids = {f.digest for f in acc}
                    for digest in [d for d in state.files if d not in accepted_ids]:
                        state.files.pop(digest)         # files outside the period
                    state.read_errors = []              # wrong files and files added twice
                    st.rerun()
    has_bank = any(f.kind == "bank" for f in acc)
    has_pay = any(f.kind == "payments" for f in acc)
    c1, _, c3 = st.columns([1, 2, 1.3])
    if c1.button("← Back", width="stretch"):
        go(1)
    if c3.button("Next: check →", type="primary", width="stretch", disabled=not (has_bank and has_pay),
                 help=None if has_bank and has_pay else "Add at least one bank file and the payment file"):
        state.run_error = None
        go(3)

# =====================================================================================================================
# 11. STEP 3 - CHECK: the checklist (engine.check_period), then "Run reconciliation"
# =====================================================================================================================
elif state.step == 3:
    acc, _ = accepted_files()
    chk = engine.check_period(period, acc, bank)
    with st.container(border=True):
        st.markdown(f'<div class="card-title">③ Check: {period.name}</div>', unsafe_allow_html=True)
        for ok, text in chk.lines:
            st.markdown(f"{'✅' if ok else '❌'} {text}")
        for w in chk.warnings:
            st.warning(w, icon="⚠️")
        if chk.blocking:
            st.error("Run is blocked until the missing bank files are added. Go back to Upload.", icon="⛔")
        if compare:
            st.caption(f"The report will be compared with {period.previous().name}.")
        if state.run_error:
            st.error(state.run_error, icon="⚠️")
    c1, _, c3 = st.columns([1, 1.5, 1.8])
    if c1.button("← Back", width="stretch"):
        go(2)
    if c3.button("▶  Run reconciliation", type="primary", width="stretch", disabled=bool(chk.blocking)):
        try:
            with st.spinner(f"Reconciling {period.name}…"):
                state.result = engine.run(BASE_DIR, period, acc, sla_days=sla_days, compare=compare, bank=bank)
            state.run_error = None
            go(4)
        except engine.ReconError as err:
            state.run_error = str(err)
            st.rerun()
        except Exception as err:
            state.run_error = f"Something unexpected went wrong: {err}"
            st.rerun()

# =====================================================================================================================
# 12. STEP 4 - RECONCILE: the result of the matching (bank rows, matched, not found, amounts), before any decision
# =====================================================================================================================
elif state.step == 4:
    res = state.result
    if res is None:
        go(3)
    with st.container(border=True):
        st.markdown(f'<div class="card-title">✅ Reconciliation completed · {bank.NAME.replace("&", "&amp;")} · '
                    f'{period.name[0].upper() + period.name[1:]}</div>', unsafe_allow_html=True)
        st.caption(f"Reconciled on {res.ran_at:%A %d/%m/%Y} at {res.ran_at:%H:%M}")
        amounts = [f"{cur} {short_amount(d['amount'])}" for cur, d in res.by_currency.items()]
        k = st.columns(4, border=True)
        k[0].metric("Bank rows", f"{res.bank_rows:,}")
        k[0].caption(f"payments of {period.name}")
        k[1].metric("Matched", f"{res.matched:,}")
        k[1].caption(f"{res.matched / res.bank_rows:.2%}" if res.bank_rows else "–")
        k[2].metric("Not found", f"{res.not_found:,}")
        k[2].caption("exceptions to resolve" if res.not_found else "none")
        k[3].metric("Amount", amounts[0] if amounts else "–")
        k[3].caption(" · ".join(amounts[1:]) or "")
        st.markdown("  \n".join(f"**{cur}** · {d['rows']:,} rows · {d['matched']:,} matched · "
                                f"{d['not_found']:,} not found" for cur, d in res.by_currency.items()))
        if res.not_found == 0:
            st.success("All rows matched. No exceptions to resolve.", icon="✅")
    c1, _, c3 = st.columns([1, 1.5, 1.8])
    if c1.button("← Back", width="stretch"):
        go(3)
    if res.not_found:
        n = res.not_found
        if c3.button(f"Review & Resolve {n:,} Exception{'s' if n > 1 else ''} →", type="primary", width="stretch"):
            go(5)
    elif c3.button("Next: download →", type="primary", width="stretch"):
        go(6)                                   # nothing to resolve: Exception Resolution is skipped

# =====================================================================================================================
# 13. STEP 5 - EXCEPTION RESOLUTION: the rows not found in IremboPay; "Save decisions" saves and runs again
# =====================================================================================================================
elif state.step == 5:
    res = state.result
    if res is None:
        go(3)
    with st.container(border=True):
        reviewed = res.not_found - res.not_reviewed
        st.markdown(f'<div class="card-title">⑤ Exception Resolution: the invoices not found in IremboPay '
                    f'<span class="muted">({reviewed} of {res.not_found} resolved)</span></div>', unsafe_allow_html=True)
        if res.not_found == 0:
            st.success("Nothing to resolve: every bank payment was found in IremboPay.", icon="✅")
        else:
            st.progress(reviewed / res.not_found)
            with st.expander("How to choose the Report table", expanded=res.not_reviewed > 0):
                st.markdown(
                    "| Report table | Use it when |\n|---|---|\n"
                    f"| **Notification received** | paid on {bank.NAME} and recorded in IremboPay |\n"
                    "| **Pending notification** | the invoice is still NEW in IremboPay |\n"
                    "| **Paid with other channel** | IremboPay says it was paid another way (type the provider if known) |\n"
                    "| **Failed notification** | the invoice is not in IremboPay at all |\n"
                    "| **Exclude** | not an Irembo payment |\n\n"
                    "Review and confirm all exception decisions before saving.  \n"
                    "*Suggested from* indicates a previously recorded decision requiring confirmation.")

            # ---- the review table: the bank columns are read only, the decision columns can be edited
            paid_col, amount_col = bank.COLUMNS["payment_date"], bank.COLUMNS["amount"]
            shown = ["Reviewed (Yes/No)", "Report table", paid_col, "Currency", amount_col, "invoice_number",
                     "Bank channel", engine.SUGGEST_COL, "SLA violated (Yes/No)", "Correct invoice number",
                     "payment_status", "payment_provider", "paid_at (YYYY-MM-DD HH:MM:SS)", "Merchant name", "Comment"]
            DASH = "—"                          # shown for "not chosen yet" in the dropdown columns
            DROPDOWNS = ["Reviewed (Yes/No)", "Report table", "SLA violated (Yes/No)"]
            table = res.review.copy()
            for c in DROPDOWNS:
                table[c] = table[c].fillna("").astype(str).replace("", DASH)
            edited = st.data_editor(
                table[shown], hide_index=True, width="stretch", key=f"review_{bank.CODE}_{period.store_key}",
                disabled=[paid_col, "Currency", amount_col, "invoice_number", "Bank channel", engine.SUGGEST_COL],
                column_config={
                    "Reviewed (Yes/No)": st.column_config.SelectboxColumn("Reviewed", options=[DASH, "Yes", "No"], width=95),
                    "Bank channel": st.column_config.TextColumn("Channel", width=80),
                    "Report table": st.column_config.SelectboxColumn("Report table", options=[DASH] + engine.TABLE_CHOICES,
                                                                     width="medium"),
                    paid_col: (st.column_config.DatetimeColumn(paid_col.capitalize(), format="DD/MM/YYYY HH:mm")
                               if paid_col in engine.option(bank, "DATETIME_COLUMNS", [])
                               else st.column_config.DateColumn(paid_col.capitalize(), format="DD/MM/YYYY")),
                    amount_col: st.column_config.NumberColumn("Amount", format="localized"),
                    "invoice_number": "Invoice",
                    engine.SUGGEST_COL: st.column_config.TextColumn("Suggested from"),
                    "SLA violated (Yes/No)": st.column_config.SelectboxColumn("SLA violated", options=[DASH, "Yes", "No"]),
                    "Correct invoice number": st.column_config.TextColumn("Correct invoice"),
                    "payment_status": st.column_config.TextColumn("Status"),
                    "payment_provider": st.column_config.TextColumn("Provider"),
                    "paid_at (YYYY-MM-DD HH:MM:SS)": st.column_config.TextColumn("paid_at", help="YYYY-MM-DD HH:MM:SS"),
                    "Merchant name": st.column_config.TextColumn("Merchant"),
                })
            for c in DROPDOWNS:
                edited[c] = edited[c].fillna(DASH).replace(DASH, "")
            yes_now = int(edited["Reviewed (Yes/No)"].fillna("").eq("Yes").sum())
            no_table = int((edited["Reviewed (Yes/No)"].fillna("").eq("Yes") & edited["Report table"].fillna("").eq("")).sum())
            st.caption(f"{yes_now} of {len(edited)} marked Yes"
                       + (f" · {no_table} without a Report table (they count as Failed notification)" if no_table else ""))
            # ---- the result at a glance, as chosen now (before saving)
            chosen = edited["Report table"].fillna("")
            now = {t: int(chosen.eq(t).sum()) for t in engine.TABLE_CHOICES}
            now["without a Report table"] = int(chosen.eq("").sum())
            st.markdown(f"**Result:** {result_line(now)}")
    c1, c2, c3 = st.columns([1, 1.5, 1.8])
    if c1.button("← Back", width="stretch"):
        go(4)
    if res.not_found and c3.button("💾  Save decisions", type="primary", width="stretch"):
        full = res.review.copy()
        for c in engine.EDIT_COLS:
            full[c] = edited[c].fillna("").astype(str).values
        try:
            engine.save_decisions(BASE_DIR, period, full, bank)
            acc, _ = accepted_files()
            with st.spinner("Saving and making the report…"):
                state.result = engine.run(BASE_DIR, period, acc, sla_days=sla_days, compare=compare, bank=bank)
            st.toast("Decisions saved", icon="💾")
            go(5 if state.result.stopped_for_review else 6)
        except engine.ReconError as err:
            st.error(str(err), icon="⚠️")
    if not res.stopped_for_review and c2.button("Next: download →", width="stretch"):
        go(6)
    if res.not_reviewed:
        st.info("The report is made as soon as every exception is resolved (Reviewed = Yes) and saved.", icon="⏸️")

# =====================================================================================================================
# 14. STEP 6 - REPORT & EXPORT
#     ① summary of the run   ② Quick Export: the standard report as it is (Excel / PDF), no configuration
#     ③ Custom Report (closed until "Customize Report"): type, sections, views, layout, preview, export
# =====================================================================================================================
elif state.step == 6:
    res = state.result
    if res is None or res.stopped_for_review or res.report_file is None:
        go(3)
    main = next((b for b in res.blocks if b[0] in ("Summary-Monthly", "Summary-Daily")), res.blocks[0])

    # ---- ① the summary of the run
    with st.container(border=True):
        st.markdown('<div class="card-title">📊 Report &amp; Export</div>', unsafe_allow_html=True)
        st.markdown(f"**{bank.FULL_NAME} · {period.name[0].upper() + period.name[1:]}**")
        st.caption(f"Reconciled {res.ran_at:%d %b %Y · %H:%M}"
                   + (f" · compared with {res.compared_with}" if res.compared_with else ""))
        k = st.columns(4, border=True)
        k[0].metric("Processed", f"{res.bank_rows:,}")
        k[1].metric("Matched", f"{res.matched:,}")
        k[2].metric("Exceptions", f"{res.not_found:,}")
        k[3].metric("Rate", f"{res.matched / res.bank_rows:.2%}" if res.bank_rows else "–")
        st.markdown("  \n".join(f"**{cur}** · {d['rows']:,} rows · {d['matched']:,} matched · "
                                f"{d['not_found']:,} exceptions" for cur, d in res.by_currency.items()))
        if res.review_counts:
            st.markdown(f"**Exceptions resolved:** {result_line(res.review_counts)}")
        for level, text in res.messages:
            if level == "warning":                  # only problems are shown (no "all good" messages)
                st.warning(text, icon="⚠️")
        with st.expander("Summary tables"):
            render_summary(res.blocks, period.report_type)

    # ---- ② Quick Export: the standard report, unchanged
    with st.container(border=True):
        st.markdown('<div class="card-title">⚡ Quick Export</div>', unsafe_allow_html=True)
        st.caption("Download the standard reconciliation report. No configuration required.")
        q1, q2, _ = st.columns([1, 1, 1])
        q1.download_button("⬇ Download Excel", data=res.report_file.read_bytes(), file_name=res.report_file.name,
                           mime=XLSX_MIME, type="primary", width="stretch", key="quick_xlsx")
        try:
            q2.download_button("⬇ Download PDF", data=pdf_bytes(str(res.report_file), res.report_file.stat().st_mtime),
                               file_name=res.report_file.with_suffix(".pdf").name, mime="application/pdf",
                               type="primary", width="stretch", key="quick_pdf")
        except Exception as err:
            q2.error(f"The PDF could not be made: {err}", icon="⚠️")

    # ---- ③ Custom Report: closed until asked for
    with st.container(border=True):
        st.markdown('<div class="card-title">⚙ Custom Report</div>', unsafe_allow_html=True)
        if not state.get("customize"):
            st.caption("Need a different report? Select the sections, visualizations and layout you want to include.")
            if st.button("Customize Report →", key="cr_open"):
                state.customize = True
                st.rerun()
        elif not getattr(res, "sheets", None):      # a run made before this version of the app
            st.info("Run the reconciliation again to customize this report.", icon="ℹ️")
        else:
            groups = engine.chart_groups(period)
            if "cr_type" not in state:              # first time: the Full Report
                state.cr_type = "Full Report"
                apply_report_type()

            st.markdown("**REPORT TYPE**")
            st.radio("Report type", list(engine.REPORT_TYPES), key="cr_type", horizontal=True,
                     captions=[d for d, _ in engine.REPORT_TYPES.values()], on_change=apply_report_type,
                     label_visibility="collapsed")

            st.markdown("**REPORT SECTIONS**")
            st.caption("Select the information to include.")
            s_cols = st.columns(2)
            for i, (key, (title, _, _)) in enumerate(engine.CUSTOM_SECTIONS.items()):
                s_cols[i % 2].checkbox(title, key=f"cr_sec_{key}")
            a1, a2, _ = st.columns([1, 1, 4])
            a1.button("Select all", on_click=tick_all_sections, args=(True,), key="cr_all", width="stretch")
            a2.button("Clear all", on_click=tick_all_sections, args=(False,), key="cr_none", width="stretch")

            chosen = [k for k in engine.CUSTOM_SECTIONS if state.get(f"cr_sec_{k}")]
            if chosen:
                st.markdown("**VISUALIZATION**")
                for key in chosen:
                    title, _, views = engine.CUSTOM_SECTIONS[key]
                    v1, v2, v3, v4 = st.columns([1.4, 1.2, 1, 1], vertical_alignment="center")
                    v1.markdown(title)
                    if state.get(f"cr_as_{key}") not in views:
                        state[f"cr_as_{key}"] = views[0]
                    shown_as = v2.selectbox("Display as", views, key=f"cr_as_{key}", disabled=len(views) == 1,
                                            label_visibility="collapsed")
                    if shown_as in (engine.BAR, engine.LINE):          # chart options: grouping and currency
                        v3.selectbox("Group by", groups, key=f"cr_by_{key}", label_visibility="collapsed",
                                     format_func=lambda g: f"By {g.lower()}")
                        v4.selectbox("Currency", ["All"] + bank.CURRENCIES, key=f"cr_cur_{key}",
                                     label_visibility="collapsed",
                                     format_func=lambda c: "All currencies" if c == "All" else c)

            st.markdown("**REPORT LAYOUT**")
            l1, l2 = st.columns(2)
            l1.radio("Orientation", ["Landscape", "Portrait"], key="cr_orient", horizontal=True)
            l2.radio("Density", ["Standard", "Compact"], key="cr_density", horizontal=True)
            st.caption("Include in the overview")
            i_cols = st.columns(len(engine.CUSTOM_INCLUDE))
            for col, (key, label) in zip(i_cols, engine.CUSTOM_INCLUDE.items()):
                if f"cr_inc_{key}" not in state:
                    state[f"cr_inc_{key}"] = True
                col.checkbox(label, key=f"cr_inc_{key}")

            content = engine.custom_content(res, custom_options(groups), bank)
            p1, _, p3 = st.columns([1.2, 2, 1])
            if p1.button("👁 Preview Report", key="cr_preview_btn", width="stretch"):
                state.cr_preview = True
            if p3.button("✖ Close", key="cr_close", width="stretch"):
                state.customize = state.cr_preview = False
                st.rerun()
            if state.get("cr_preview"):
                render_custom_preview(content)

            st.markdown("**EXPORT**")
            e1, e2 = st.columns([1, 1.6], vertical_alignment="bottom")
            fmt = e1.radio("Format", ["Excel", "PDF"], key="cr_fmt", horizontal=True)
            if not chosen:
                e2.button("⬇ Generate & Download Report", disabled=True, width="stretch", key="cr_dl_off",
                          help="Select at least one section")
            else:
                try:
                    if fmt == "PDF":
                        data, ext, mime = engine.custom_pdf(content), "pdf", "application/pdf"
                    else:
                        data, ext, mime = engine.custom_excel(content), "xlsx", XLSX_MIME
                    e2.download_button("⬇ Generate & Download Report", data=data,
                                       file_name=f"{content['file_stem']}.{ext}", mime=mime, type="primary",
                                       width="stretch", key="cr_dl")
                except Exception as err:
                    e2.error(f"The custom report could not be made: {err}", icon="⚠️")

    c1, _, c3 = st.columns([1, 1.5, 1.8])
    if c1.button("← Back", width="stretch"):
        go(5 if res.not_found else 4)           # Exception Resolution was skipped when nothing was missing
    if c3.button("🔄  New report", width="stretch"):
        state.files, state.read_errors, state.result = {}, [], None
        state.customize = state.cr_preview = False
        go(1)
