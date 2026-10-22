"""
BANK RECONCILIATION APP (IremboPay)
===================================
The screens of the app: a side panel + 3 steps. It contains no reconciliation rules: those are in engine.py
(common rules) and banks/<bank>.py (what is specific to one bank).

Start it with "Start Reconciliation.command" (double-click). Files from a bank's folders (or uploaded by hand)
are read in memory and never kept; only the reports, the review decisions and the run log are saved (engine.py).

How Streamlit works (to read this file): the whole file runs again from top to bottom after every click. What must
be remembered between clicks is kept in `state` (st.session_state, section 4). The page is drawn in the order of
the code: side panel, banner, then the screen of the current step.

The page
    ┌ side panel ──────────┐ ┌ banner: iremboPay | <Bank> Reconciliation ───────────┐
    │ Reports              │ │ 1. Choose  2. Check  3. Results                      │
    │ Recent runs          │ │                                                       │
    │ Settings             │ │ the screen of the current step (or a saved report)   │
    └──────────────────────┘ └───────────────────────────────────────────────────────┘

Reconciliation runs automatically once Check passes (no separate "Reconcile" screen); the rules decide pending /
duplicate / no invoice number by themselves (services/auto_review.py) - only the rare row they truly cannot
decide is listed as "Needs a look" on Results, and it never blocks the report.

CONTENTS
    1.  Loading the engine and the banks
    2.  Settings
    3.  Page setup and style
    4.  Memory between clicks (session state)
    5.  Shared helpers (navigation, downloads, files, key figures, summary tables)
    6.  Side panel
    7.  Banner and step bar
    8.  A saved report (read only, opened from the side panel)
    9.  Step 1 - Choose
    10. Step 2 - Check (runs the reconciliation)
    11. Step 3 - Results (quick export + custom report)
"""
import base64
import hashlib
import importlib
import os
import io
import re
import shutil
import subprocess
import sys
from pathlib import Path

import pandas as pd
import streamlit as st

APP_DIR = Path(__file__).resolve().parent


# =====================================================================================================================
# 1. LOADING THE ENGINE AND THE BANKS
#    A running app keeps the code it started with; when engine.py or a bank file has changed, it is loaded again.
# =====================================================================================================================
# the app's own code besides app.py; reloaded in this order (each one only uses the ones before it).
# integrations/ and config.py are not reloaded: a connected Drive client and its error class must stay the same.
RELOADED = ["services.app_settings", "pdf_report", "integrations.irembopay_checkout", "integrations.slack", "services.session_files",
            "services.report_history", "services.dashboard", "services.dashboard_view", "services.pay_files",
            "services.slack_message", "services.folder_files", "services.auto_review"]


def _code_mtime():
    """Last change time of engine.py, the bank files, pdf_report.py, services/ and the IremboPay check."""
    return max(p.stat().st_mtime for p in [APP_DIR / "engine.py", APP_DIR / "pdf_report.py",
                                           APP_DIR / "integrations" / "irembopay_checkout.py",
                                           APP_DIR / "integrations" / "slack.py",
                                           *APP_DIR.glob("banks/*.py"), *APP_DIR.glob("services/*.py")])


try:
    import engine
    # reload the code when it has changed (otherwise a running app keeps the old version until it is restarted)
    if getattr(engine, "_loaded_mtime", None) != _code_mtime():
        for _name in [n for n in sys.modules if n.startswith("banks.")] + ["banks"]:
            if _name in sys.modules:           # not there yet when the app was started before banks/ existed
                importlib.reload(sys.modules[_name])
        engine = importlib.reload(engine)
        for _name in RELOADED:
            if _name in sys.modules:           # loaded below on the first run
                importlib.reload(sys.modules[_name])
        engine._loaded_mtime = _code_mtime()
except ValueError as _err:                     # a bank file is incomplete
    st.error(str(_err), icon="⛔")
    st.stop()

import pdf_report                              # the PDF layout of a saved report (presentation only)
import config                                  # Google Drive settings (no secrets: only the path to the credentials)
from integrations import google_drive
from integrations.google_drive import DriveClient, DriveError
from integrations import irembopay_checkout, slack
from services import (app_settings, auto_review, dashboard, dashboard_view, folder_files, folder_picker, pay_files,
                     report_history, slack_message)
from services.session_files import SessionFolder


# =====================================================================================================================
# 2. SETTINGS
# =====================================================================================================================
# the "Bank Reconciliation" folder (output/ is created inside it); IM_RECON_BASE is only used for testing
BASE_DIR = Path(os.environ.get("IM_RECON_BASE", Path(__file__).resolve().parent.parent))
BRAND = "#123239"          # IremboPay dark teal
ACCENT = "#99E378"         # IremboPay green
# the side menu background: the two curves of the Irembo mark, in a faint green on the dark teal (like the brand
# slides, where they are a shade of the background colour); and the green menu sign next to "Payment Operations"
SIDE_BG = "url(data:image/svg+xml;base64," + base64.b64encode(b"""<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 380 1100">
<defs><radialGradient id="g" cx="50%" cy="0%" r="70%"><stop offset="0" stop-color="#2F6B5A" stop-opacity=".55"/>
<stop offset="1" stop-color="#133A3B" stop-opacity="0"/></radialGradient></defs>
<rect width="380" height="1100" fill="url(#g)"/>
<g fill="none" stroke="#99E378" stroke-opacity=".07" stroke-linecap="round" stroke-width="78" transform="translate(0 420)">
<path d="M150 210 C 40 250 -30 380 30 520"/>
<path d="M230 200 C 380 240 470 420 450 640 C 440 760 400 860 360 960"/>
</g></svg>""").decode() + ")"
SIDE_MENU = "url(data:image/svg+xml;base64," + base64.b64encode(b"""<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 40 30">
<g stroke="#99E378" stroke-width="3.4" stroke-linecap="round"><path d="M4 5h32"/><path d="M4 15h22"/><path d="M4 25h32"/></g>
</svg>""").decode() + ")"
STEPS = ["Choose", "Check", "Results"]
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

st.set_page_config(page_title=f"Bank Reconciliation", page_icon=str(APP_DIR / "assets" / "irembopay_icon.png"), layout="wide",
                   initial_sidebar_state="expanded")
st.markdown(f"""
<style>
  /* colours of text that change with the theme (light / dark, chosen in the menu > Settings) */
  .stApp {{--rc-green: light-dark(#2E7D4F, #7FD18F); --rc-muted: light-dark(#5B6B70, #A9BABE);
          --rc-red: light-dark(#C2413D, #F08A85); --rc-teal: light-dark(#2B8A99, #6CC3CF);
          --rc-ink: light-dark({BRAND}, #E6EEF0);}}
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
  .step.now {{background: {BRAND}; border-color: light-dark({BRAND}, {ACCENT}); color: {ACCENT}; opacity: 1; font-weight: 650;}}
  /* card titles and small grey text */
  .card-title {{font-weight: 650; font-size: 1.12rem; margin-bottom: .3rem;}}
  .muted {{opacity: .7; font-size: .9rem;}}
  /* No-Transaction Report preview: the Excel sheet as it will be (white paper, the report's own colours) */
  .ntx-paper {{background: #fff; color: #1D2B2F; border: 1px solid #9AA8AD; border-radius: 10px;
              padding: 1.3rem 1.4rem; overflow-x: auto; box-shadow: 0 2px 10px rgba(0,0,0,.12);}}
  .ntx-paper h3 {{color: #1D2B2F; font-size: 1.2rem; margin: 0; padding: 0;}}
  .ntx-paper .meta {{display: flex; justify-content: space-between; flex-wrap: wrap; gap: .4rem 1.5rem;
                    font-style: italic; color: #5B6B70; font-size: .88rem; margin: .2rem 0 1rem 0;}}
  .ntx-paper table {{border-collapse: collapse; width: 100%; min-width: 560px; font-size: .9rem; color: #1D2B2F;}}
  .ntx-paper th, .ntx-paper td {{border: 1px solid #000; padding: .3rem .55rem; text-align: left;}}
  .ntx-paper .band {{background: #D0DEE3; text-align: center; font-size: 1.15rem; font-weight: 400;}}
  .ntx-paper .num {{text-align: right;}}
  .ntx-paper .gap td {{border: none; padding: .45rem;}}
  .ntx-paper .status-head td {{background: #D0DEE3; font-weight: 700;}}
  .ntx-paper .status td {{background: #FFF3D6; color: #7A4E00;}}
  .ntx-paper .status .title {{font-weight: 700; font-size: 1rem;}}
  div[data-testid="stMetricValue"] {{font-size: 1.5rem;}}
  /* side panel: IremboPay colours (the Nunito font is set for the whole app in .streamlit/config.toml) */
  .side-head {{display: flex; align-items: center; gap: .55rem; margin: .1rem 0 .55rem 0;
               font-family: "Nunito", sans-serif; font-weight: 800; font-size: 1.08rem; color: var(--rc-ink);}}
  .side-icon {{display: inline-flex; align-items: center; justify-content: center; width: 1.75rem; height: 1.75rem;
               border-radius: .5rem; background: {BRAND}; color: {ACCENT}; flex: none;}}
  .side-icon svg {{width: 1rem; height: 1rem;}}
  .side-count {{margin-left: auto; background: {ACCENT}; color: {BRAND}; font-weight: 800; font-size: .74rem;
                padding: .12rem .6rem; border-radius: 999px; white-space: nowrap;}}
  .side-empty {{font-family: "Nunito", sans-serif; color: var(--rc-muted); font-size: .9rem; margin: 0 0 .4rem .2rem;}}
  section[data-testid="stSidebar"] [data-testid="stExpander"] summary p {{
      font-family: "Nunito", sans-serif; font-weight: 800; color: var(--rc-ink);}}
  section[data-testid="stSidebar"] [data-testid="stExpander"] summary [data-testid="stIconMaterial"] {{color: var(--rc-ink);}}
  /* dropdown lists: exactly 2 choices shown (40px each); more than 2 -> scroll to see the rest */
  [data-testid="stSelectboxVirtualDropdown"] > div {{max-height: 80px !important; overflow-y: auto !important;}}
  section[data-testid="stSidebar"] [data-testid="stSelectbox"] input {{cursor: pointer; caret-color: transparent;}}
  /* a wider side panel with a larger page menu (Dashboard, Reconciliation, Reports, Bank Management, Settings) */
  section[data-testid="stSidebar"][aria-expanded="true"] {{width: 380px !important; min-width: 380px !important;}}
  /* collapsed: width must go to 0 too (not just min-width), or a 380px-wide panel leaves part of itself showing -
     on a narrow screen that sliver sits on top of the page and covers real text */
  section[data-testid="stSidebar"][aria-expanded="false"] {{width: 0 !important; min-width: 0 !important;}}
  /* the side menu: dark teal with the Irembo curves in the background (green tint), "Payment Operations" on top,
     then Workspace / Configuration; the page open has a green bar and a green dot */
  section[data-testid="stSidebar"] > div {{background-color: #133A3B; background-image: {SIDE_BG};
      background-repeat: no-repeat; background-size: 100% auto; background-position: center top;}}
  [data-testid="stSidebarHeader"] {{background: transparent; margin: 0 0 .4rem 0; padding: 1.4rem .4rem .6rem .4rem;
      position: relative; height: auto; align-items: center;}}
  [data-testid="stSidebarHeader"]::before {{content: "Payment Operations"; display: flex; align-items: center; flex: 1;
      padding-left: 2.9rem; min-height: 2rem; color: #fff; font-size: 1.15rem; font-weight: 800;
      background: {SIDE_MENU} no-repeat left center / 1.9rem auto;}}
  [data-testid="stSidebarHeader"] [data-testid="stSidebarCollapseButton"] button,
  [data-testid="stSidebarHeader"] [data-testid="stSidebarCollapseButton"] [data-testid="stIconMaterial"] {{color: #D6E4E6;}}
  [data-testid="stSidebarHeader"] [data-testid="stSidebarCollapseButton"] {{display: block !important;}}
  [data-testid="stNavSectionHeader"] {{text-transform: uppercase; letter-spacing: .08em; font-size: .8rem !important;
      color: #9FBFC3 !important; font-weight: 700; padding: .3rem .8rem .3rem .8rem; margin-top: .3rem;}}
  [data-testid="stSidebarNavItems"] > div + div {{border-top: 1px solid rgba(255,255,255,.16); padding-top: .8rem;
      margin-top: 1rem;}}
  [data-testid="stSidebarNav"] a {{padding: .75rem .9rem; border-radius: 12px; margin: .2rem 0; position: relative;
      border-left: 4px solid transparent;}}
  [data-testid="stSidebarNav"] a:hover {{background: rgba(255,255,255,.06);}}
  [data-testid="stSidebarNav"] a span {{font-size: 1.1rem !important; font-weight: 600; color: #F2F7F8;}}
  [data-testid="stSidebarNav"] a [data-testid="stIconMaterial"] {{font-size: 1.45rem !important; color: #F2F7F8;
      margin-right: .6rem;}}
  [data-testid="stSidebarNav"] a[aria-current="page"] {{background: rgba(153,227,120,.13); border-left-color: {ACCENT};}}
  [data-testid="stSidebarNav"] a[aria-current="page"]::after {{content: ""; position: absolute; right: 1rem; top: 50%;
      width: 9px; height: 9px; margin-top: -4.5px; border-radius: 50%; background: {ACCENT};}}
  /* the side panel's own tools (Reconciliation page) on the dark background */
  section[data-testid="stSidebar"] .side-head {{color: #fff;}}
  section[data-testid="stSidebar"] .side-empty, section[data-testid="stSidebar"] .side-note,
  section[data-testid="stSidebar"] [data-testid="stCaptionContainer"],
  section[data-testid="stSidebar"] [data-testid="stCaptionContainer"] p {{color: #A9C4C8 !important;}}
  section[data-testid="stSidebar"] [data-testid="stExpander"] details {{border-color: rgba(255,255,255,.18);
      background: rgba(255,255,255,.04);}}
  section[data-testid="stSidebar"] [data-testid="stExpander"] summary p,
  section[data-testid="stSidebar"] [data-testid="stExpander"] summary [data-testid="stIconMaterial"] {{color: #fff !important;}}
  section[data-testid="stSidebar"] [data-testid="stBaseButton-secondary"] {{background: rgba(255,255,255,.07);
      border-color: rgba(255,255,255,.22); color: #fff;}}
  section[data-testid="stSidebar"] [data-testid="stBaseButton-secondary"] p {{color: #fff;}}
  section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] > p {{color: #E6F0F1;}}
  section[data-testid="stSidebar"] hr {{border-color: rgba(255,255,255,.16);}}
  .side-note {{font-size: .82rem; color: var(--rc-muted); margin: -.35rem 0 .45rem .2rem;}}
  /* Report History: header, folder map, folder groups, report cards */
  .rh-head {{background: {BRAND}; border-left: 6px solid {ACCENT}; border-radius: 14px; padding: .95rem 1.3rem;
             margin: .2rem 0 .9rem 0; color: #fff;}}
  .rh-crumb {{font-size: .8rem; color: #D6E4E6; letter-spacing: .02em;}}
  .rh-crumb b {{color: {ACCENT};}}
  .rh-title {{font-size: 1.3rem; font-weight: 800; color: {ACCENT}; margin: .15rem 0 .55rem 0;}}
  /* Step 2 Upload: two cards (bank statement / IremboPay file) */
  .up-head {{display: flex; gap: .9rem; align-items: flex-start; margin-bottom: .35rem;}}
  .up-tile {{flex: none; width: 3.1rem; height: 3.1rem; border-radius: 12px; display: flex; align-items: center;
             justify-content: center; background: rgba(46,125,79,.13); color: var(--rc-green);}}
  .up-tile svg {{width: 1.55rem; height: 1.55rem;}}
  .up-title {{font-weight: 800; font-size: 1.2rem; line-height: 1.3;}}
  .up-sub {{font-size: .92rem; opacity: .65; margin-top: .15rem;}}
  .up-drop {{text-align: center; padding: 1.1rem .8rem 1rem; border: 1px solid rgba(128,128,128,.28);
             border-radius: 12px; background: rgba(128,128,128,.05); margin: .2rem 0 .2rem 0;}}
  .up-drop svg {{width: 1.7rem; height: 1.7rem; color: var(--rc-green);}}
  .up-drop-title {{font-weight: 650; font-size: 1.12rem; margin-top: .45rem;}}
  .up-drop-sub {{font-size: .86rem; opacity: .65; margin-top: .3rem;}}
  .up-file {{padding: .8rem 1rem; border: 1px solid rgba(128,128,128,.28); border-radius: 12px; margin: .1rem 0 .2rem 0;}}
  .up-file-label {{font-size: .85rem; opacity: .65;}}
  .up-file-name {{font-weight: 650; font-size: 1.12rem; margin: .2rem 0 .15rem 0; word-break: break-all;}}
  .up-file-path {{font-size: .88rem; opacity: .65;}}
  .up-help {{font-size: .86rem; opacity: .7; line-height: 1.55; margin-top: .35rem;}}
  /* Google Drive: "Select IremboPay file", one card per file (a radio list drawn as cards) */
  .pay-pick-title {{font-weight: 800; font-size: 1.08rem; margin: .35rem 0 .1rem 0;}}
  .pay-pick-sub {{font-size: .9rem; opacity: .7; margin-bottom: .2rem;}}
  .st-key-drive_pay_pick, .st-key-drive_pay_pick [data-testid="stRadio"] {{width: 100% !important;}}
  .st-key-drive_pay_pick [role="radiogroup"] {{width: 100%; gap: .7rem; flex-wrap: nowrap;}}
  /* more than one file: exactly one card is shown; scrolling moves one card at a time (snaps) */
  .st-key-drive_pay_pick [role="radiogroup"] > div {{flex: none; height: 8.6rem; display: flex;
      flex-direction: column; justify-content: center; scroll-snap-align: start; scroll-snap-stop: always;}}
  .st-key-drive_pay_pick [role="radiogroup"]:has(> div:nth-child(2)) {{height: 8.6rem; overflow-y: auto;
      scroll-snap-type: y mandatory; overscroll-behavior: contain; padding-right: .45rem; scrollbar-width: thin;
      scrollbar-color: rgba(46,125,79,.45) transparent;}}
  .pay-missing {{background: rgba(212,167,44,.12); border: 1px solid rgba(212,167,44,.45); border-radius: 10px;
      padding: .65rem .9rem; font-size: .9rem; color: light-dark(#7A4E00, #F2C14E); margin: .1rem 0 .35rem 0;}}
  .pay-pick-more {{font-size: .82rem; color: var(--rc-green); font-weight: 650; margin: -.2rem 0 .1rem 0;}}
  .st-key-drive_pay_pick [role="radiogroup"] > div {{position: relative; width: 100%; box-sizing: border-box;
      padding: .85rem 3.2rem .85rem 3.3rem; border: 1px solid rgba(128,128,128,.38); border-radius: 12px;
      cursor: pointer; transition: border-color .15s, background .15s;}}
  .st-key-drive_pay_pick [role="radiogroup"] > div:hover {{border-color: rgba(46,125,79,.6);}}
  .st-key-drive_pay_pick [role="radiogroup"] > div[data-selected="true"] {{border-color: var(--rc-green);
      background: rgba(46,125,79,.07);}}
  .st-key-drive_pay_pick [role="radiogroup"] > div::before {{content: ""; position: absolute; left: 1rem; top: 50%;
      transform: translateY(-50%); width: 1.35rem; height: 1.35rem; background: no-repeat center / contain
      url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='%232E7D4F' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'%3E%3Crect x='3.5' y='3.5' width='17' height='17' rx='2'/%3E%3Cpath d='M3.5 9.5h17M9.5 9.5v11'/%3E%3C/svg%3E");}}
  .st-key-drive_pay_pick [role="radiogroup"] > div[data-selected="true"]::after {{content: ""; position: absolute;
      right: 1.05rem; top: 50%; transform: translateY(-50%); width: 1.35rem; height: 1.35rem; background: no-repeat
      center / contain url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='%232E7D4F' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'%3E%3Ccircle cx='12' cy='12' r='9.5'/%3E%3Cpath d='m8 12.3 2.7 2.7L16.2 9.5'/%3E%3C/svg%3E");}}
  .st-key-drive_pay_pick [data-testid="stRadioOption"]::before {{content: ""; position: absolute; inset: 0;}}
  .st-key-drive_pay_pick [data-testid="stRadioOption"] > div > div:first-child {{position: absolute; right: 1rem;
      top: 50%; transform: translateY(-50%); margin: 0;}}
  .st-key-drive_pay_pick [role="radiogroup"] > div[data-selected="true"] [data-testid="stRadioOption"] > div > div:first-child
      {{visibility: hidden;}}
  .st-key-drive_pay_pick [data-testid="stRadioOption"] p {{font-size: 1.05rem;}}
  .st-key-drive_pay_pick [data-testid="stRadioCaption"] {{padding-left: 0; margin-top: .2rem; pointer-events: none;}}
  .st-key-drive_pay_pick [data-testid="stRadioCaption"] p {{font-size: .88rem; line-height: 1.6; opacity: .85;}}
  .st-key-drive_pay_pick [role="radiogroup"] > div[data-selected="true"] [data-testid="stRadioCaption"]::after {{
      content: "Selected"; display: block; color: var(--rc-green); font-size: .85rem; font-weight: 650; margin-top: .3rem;}}
  /* the bank files added: cards in the upload area (scroll inside it) */
  .st-key-up_bank_files {{background: #F7FAFA; border: 1.5px solid #CFE3E6; border-radius: 18px; padding: .9rem 1rem;}}
  .st-key-up_bank_files [class*="st-key-upf_"] {{background: #fff; border-radius: 12px !important;
      border-color: #E1EEF0 !important; padding: .55rem .7rem !important;}}
  .st-key-up_bank_files [class*="st-key-upf_"] {{position: relative; gap: 0;}}
  .st-key-up_bank_files [class*="st-key-upf_x_"] {{position: absolute; right: .5rem; top: 50%; transform: translateY(-50%);
      width: auto !important;}}
  .upf {{padding-right: 1.6rem;}}
  .st-key-up_bank_card [data-testid="stFileUploaderDropzone"] > span > button {{min-height: 2.9rem !important; height: 2.9rem;}}
  /* the files area takes the free height of the card, so "Choose bank file" lines up with "Add selected file" */
  .st-key-up_bank_card [data-testid="stLayoutWrapper"]:has(> .st-key-up_bank_files) {{flex: 1 1 auto; display: flex;
      flex-direction: column;}}
  .st-key-up_bank_files {{flex: 1 1 auto; height: auto !important; min-height: 300px; max-height: 520px;}}
  .st-key-up_bank_card, .st-key-up_pay_card {{display: flex; flex-direction: column;}}
  .st-key-up_pay_card:not(:has(> .st-key-drive_pay_add)) > :last-child,
  .st-key-up_pay_card > .st-key-drive_pay_add {{margin-top: auto;}}     /* Add + Refresh at the bottom of the card */
  .upf {{display: flex; align-items: center; gap: .7rem; min-width: 0;}}
  .upf-ic {{width: 2.1rem; height: 2.4rem; border-radius: 7px; color: #fff; font-size: .68rem; font-weight: 800; flex: none;
      display: flex; align-items: flex-end; justify-content: center; padding-bottom: .3rem; box-sizing: border-box;}}
  .upf-ic.rwf {{background: {BRAND};}} .upf-ic.usd {{background: #2B8A99;}}
  .upf-t {{min-width: 0;}}
  .upf-d {{font-weight: 800; font-size: 1rem; color: {BRAND}; display: flex; align-items: center; gap: .45rem;
      white-space: nowrap;}}
  .upf-c {{font-size: .7rem; font-weight: 800; padding: .05rem .5rem; border-radius: 999px; background: #EAF4F5;}}
  .upf-c.usd {{background: #DCEFF2; color: #1F6E7A;}}
  .upf-n {{font-size: .8rem; color: var(--rc-muted); white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
      margin-top: .15rem;}}
  .st-key-up_bank_files [class*="st-key-upf_x_"] button {{color: #9AA8AD; min-height: 0; padding: 0;}}
  .st-key-up_bank_files [class*="st-key-upf_x_"] button:hover {{color: #C2413D;}}
  .upf-foot {{text-align: center; font-size: .85rem; color: var(--rc-green); font-weight: 700; margin: -.3rem 0 .2rem 0;}}
  /* both cards as tall as the taller one */
  [data-testid="stLayoutWrapper"]:has(> .st-key-up_bank_card),
  [data-testid="stLayoutWrapper"]:has(> .st-key-up_pay_card) {{flex: 1 1 auto;}}
  .st-key-up_bank_card, .st-key-up_pay_card {{height: 100%;}}
  /* the bank card: its upload area grows to fill the card (no empty space under the text when the IremboPay card
     is taller), its icon and text stay centred */
  .st-key-up_bank_card > [data-testid="stElementContainer"]:has(.up-drop-fill) {{flex: 1 1 auto; display: flex;}}
  .st-key-up_bank_card [data-testid="stElementContainer"]:has(.up-drop-fill) > div,
  .st-key-up_bank_card [data-testid="stElementContainer"]:has(.up-drop-fill) [data-testid="stMarkdown"],
  .st-key-up_bank_card [data-testid="stElementContainer"]:has(.up-drop-fill) [data-testid="stMarkdown"] > div,
  .st-key-up_bank_card [data-testid="stMarkdownContainer"]:has(> .up-drop-fill) {{height: 100%; width: 100%;
      margin-bottom: 0;}}
  .up-drop-fill {{height: 100%; box-sizing: border-box; margin: 0; display: flex; flex-direction: column;
      align-items: center; justify-content: center;}}
  /* under the two main buttons ("Choose bank file", "Add selected file" / "Choose IremboPay file") the same room:
     the bank card's note, the IremboPay card's "Refresh files" or note; so the two buttons stay on one line */
  .st-key-up_bank_card .up-help, .st-key-up_pay_card .up-help {{height: 4.4rem; box-sizing: border-box;
      overflow: hidden; margin-top: 0;}}
  .st-key-drive_pay_refresh {{height: 3.4rem;}}     /* + the 1rem gap between two elements = the note's 4.4rem */
  /* the upload widget becomes one big button ("Choose … file"); a file can still be dropped on it */
  .st-key-up_bank_card [data-testid="stFileUploaderDropzone"],
  .st-key-up_pay_card [data-testid="stFileUploaderDropzone"] {{padding: 0; border: none; background: transparent;
      min-height: 0 !important; height: auto;}}
  .st-key-drive_pay_refresh button, .st-key-drive_pay_add button {{min-height: 2.9rem; border-radius: 10px;}}
  .st-key-drive_pay_add button p, .st-key-drive_pay_refresh button p {{font-size: 1.02rem;}}
  .st-key-up_bank_card [data-testid="stFileUploaderDropzoneInstructions"],
  .st-key-up_pay_card [data-testid="stFileUploaderDropzoneInstructions"] {{display: none;}}
  .st-key-up_bank_card [data-testid="stFileUploader"],
  .st-key-up_pay_card [data-testid="stFileUploader"] {{width: 100%;}}
  .st-key-up_bank_card [data-testid="stFileUploaderDropzone"] > span,
  .st-key-up_pay_card [data-testid="stFileUploaderDropzone"] > span {{display: block; width: 100%;}}
  .st-key-up_bank_card [data-testid="stFileUploaderDropzone"] > span > button,
  .st-key-up_pay_card [data-testid="stFileUploaderDropzone"] > span > button {{
      width: 100%; min-height: 3.1rem; border-radius: 10px; background: {BRAND}; border-color: {BRAND}; color: #fff;}}
  .st-key-up_bank_card [data-testid="stFileUploaderDropzone"] > span > button > div,
  .st-key-up_pay_card [data-testid="stFileUploaderDropzone"] > span > button > div {{display: none;}}
  .st-key-up_bank_card [data-testid="stFileUploaderDropzone"] > span > button:hover,
  .st-key-up_pay_card [data-testid="stFileUploaderDropzone"] > span > button:hover {{background: #1b4750; color: #fff;}}
  .st-key-up_bank_card [data-testid="stFileUploaderDropzone"] > span > button::after {{
      content: "Choose bank file"; font-size: 1.05rem; font-weight: 600;}}
  .st-key-up_pay_card [data-testid="stFileUploaderDropzone"] > span > button::after {{
      content: "Choose IremboPay file"; font-size: 1.05rem; font-weight: 600;}}
  /* while files upload: no list of file chips popping in and out; a steady bar the size of the button instead */
  .st-key-up_bank_card [data-testid="stFileChips"], .st-key-up_pay_card [data-testid="stFileChips"],
  .st-key-up_bank_card [data-testid="stFileUploaderDropzone"]:has([data-testid="stFileChips"]) > span,
  .st-key-up_pay_card [data-testid="stFileUploaderDropzone"]:has([data-testid="stFileChips"]) > span {{display: none;}}
  .st-key-up_bank_card [data-testid="stFileUploaderDropzone"]:has([data-testid="stFileChips"])::after,
  .st-key-up_pay_card [data-testid="stFileUploaderDropzone"]:has([data-testid="stFileChips"])::after {{
      content: "Reading your files…"; display: flex; align-items: center; justify-content: center; width: 100%;
      min-height: 3.1rem; border-radius: 10px; background: rgba(46,125,79,.12); color: var(--rc-green); font-weight: 650;
      font-size: 1.02rem;}}
  /* the source switch fills the card's width: two equal halves, the chosen one in teal */
  .st-key-up_pay_card .st-key-pay_source {{width: 100% !important;}}
  .st-key-up_pay_card [data-testid="stButtonGroup"] {{width: 100%;}}
  .st-key-up_pay_card [data-testid="stButtonGroup"] [role="radiogroup"] {{width: 100%; max-width: 100% !important;
      display: flex; gap: .5rem;}}
  .st-key-up_pay_card [data-testid="stButtonGroup"] button {{flex: 1 1 0; min-height: 2.9rem; border-radius: 10px;
      background: rgba(128,128,128,.10); border: 1px solid transparent;}}
  .st-key-up_pay_card [data-testid="stButtonGroup"] button p {{font-size: 1.02rem;}}
  .st-key-up_pay_card [data-testid="stButtonGroup"] button[aria-checked="true"] {{background: {BRAND};
      border-color: {BRAND};}}
  .st-key-up_pay_card [data-testid="stButtonGroup"] button[aria-checked="true"] p {{color: #fff;}}
  .rh-empty {{text-align: center; padding: 1.6rem 1rem; opacity: .75; border: 1px dashed rgba(128,128,128,.45);
              border-radius: 12px; font-size: .92rem;}}
  .rh-sub {{font-size: .9rem; color: #D6E4E6; margin-top: -.25rem;}}
  /* stat cards, section titles, report rows (Report History) */
  .rh-k-label {{font-size: .88rem; opacity: .7;}}
  .rh-k-value {{font-size: 1.9rem; font-weight: 800; line-height: 1.25; margin: .35rem 0 .2rem 0;}}
  .rh-k-value.small {{font-size: 1.45rem; padding: .2rem 0 .15rem 0;}}
  .rh-k-sub {{font-size: .88rem; opacity: .65;}}
  .rh-sec {{display: flex; align-items: baseline; justify-content: space-between; gap: 1rem;}}
  .rh-sec-title {{font-weight: 800; font-size: 1.15rem;}}
  .rh-count {{font-size: .95rem; opacity: .65; white-space: nowrap;}}
  .rh-item {{display: flex; gap: 1rem; align-items: flex-start; padding: .2rem 0 .1rem 0;}}
  .rh-tile {{flex: none; width: 3.2rem; height: 3.2rem; border-radius: 12px; display: flex; align-items: center;
             justify-content: center; background: rgba(46,125,79,.13); color: var(--rc-green);}}
  .rh-tile.pdf {{background: rgba(194,65,61,.13); color: var(--rc-red);}}
  .rh-tile svg {{width: 1.6rem; height: 1.6rem;}}
  .rh-name {{font-weight: 800; font-size: 1.12rem; line-height: 1.3;}}
  .rh-period {{font-size: 1rem; opacity: .85; margin: .15rem 0 .35rem 0;}}
  .rh-badges {{display: flex; align-items: center; gap: .55rem; flex-wrap: wrap;}}
  .rh-badge {{font-size: .82rem; font-weight: 650; padding: .12rem .65rem; border-radius: 999px;
              background: rgba(46,125,79,.14); color: var(--rc-green);}}
  .rh-badge.pdf {{background: rgba(194,65,61,.13); color: var(--rc-red);}}
  .rh-badge.ver {{background: rgba(43,138,153,.15); color: var(--rc-teal);}}
  .rh-size {{font-size: .85rem; opacity: .65;}}
  .rh-saved {{font-size: .85rem; opacity: .65; margin-top: .3rem;}}
  /* step 1, "Compare": the previous report was found - a short note that fades out and folds away after 4 s */
  .prev-found {{background: rgba(46,125,79,.09); border: 1px solid rgba(46,125,79,.22); border-radius: 10px;
      padding: .7rem 1rem; color: var(--rc-green); font-size: .95rem;}}
  .prev-found span {{opacity: .85;}}
  [data-testid="stMarkdownContainer"]:has(> .prev-found) {{margin-bottom: 0;}}    /* else its bottom is cut off */
  [data-testid="stElementContainer"]:has(.prev-found) {{overflow: hidden; max-height: 6rem;
      animation: rc-note-out .5s ease 4s forwards;}}
  @keyframes rc-note-out {{to {{opacity: 0; max-height: 0; margin-top: -1rem;}}}}
  [data-testid="stElementContainer"]:has(.prev-found.save-note) {{animation-delay: 6s;}}   /* saved on Drive: 6 s */
  /* Reports Dashboard header: a light card (not the dark teal band of the IremboPay banner above it) */
  .dash-head {{display: flex; align-items: center; gap: 1rem; padding: 1rem 1.3rem; margin: .2rem 0 .9rem 0;
      border-radius: 14px; border: 1px solid rgba(46,125,79,.25);
      background: linear-gradient(120deg, rgba(153,227,120,.22), rgba(43,138,153,.08) 70%);}}
  .dash-head-icon {{flex: none; width: 3rem; height: 3rem; border-radius: 12px; background: {BRAND};
      color: {ACCENT}; display: flex; align-items: center; justify-content: center;}}
  .dash-head-icon svg {{width: 1.6rem; height: 1.6rem;}}
  .dash-head-title {{font-size: 1.25rem; font-weight: 800; color: var(--rc-ink); line-height: 1.3;}}
  .dash-head-sub {{font-size: .9rem; color: var(--rc-muted);}}
  .dash-head-src {{margin-left: auto; flex: none; font-size: .8rem; color: var(--rc-muted); white-space: nowrap;
      padding: .35rem .8rem; border-radius: 999px; border: 1px solid rgba(128,128,128,.25);
      background: light-dark(rgba(255,255,255,.75), rgba(255,255,255,.05));}}
  .dash-head-src b {{color: var(--rc-green);}}
  /* Reports Dashboard: key figure cards, card titles, status table, reports table, highlights */
  .dash-used {{font-size: 1.45rem; font-weight: 800; color: var(--rc-ink); line-height: 1.3;}}
  .dash-kpis {{display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: .8rem; margin: .2rem 0 1rem 0;}}
  .dash-kpi {{border: 1px solid rgba(128,128,128,.28); border-left: 4px solid {ACCENT}; border-radius: 12px;
      padding: .8rem .95rem;}}
  .dash-kpi.warn {{border-left-color: #D4A72C; background: rgba(212,167,44,.10);}}
  .dash-k-label {{font-size: .72rem; font-weight: 800; letter-spacing: .04em; text-transform: uppercase;
      color: var(--rc-muted);}}
  .dash-k-value {{font-size: 1.65rem; font-weight: 800; color: var(--rc-ink); line-height: 1.35;}}
  .dash-kpi.warn .dash-k-value {{color: light-dark(#7A4E00, #F2C14E);}}
  .dash-k-value.multi {{font-size: 1.05rem; line-height: 1.45; padding: .2rem 0;}}
  .dash-k-sub {{font-size: .8rem; color: var(--rc-muted);}}
  /* Reports Dashboard (manager view): key-figure tiles, coloured by their target, and the one-line findings */
  .dt-caption {{font-size: .82rem; color: var(--rc-muted); margin: .1rem 0 .45rem 0;}}
  .dt-groups {{display: grid; gap: .8rem; margin: 0 0 1rem 0; grid-template-columns: 1.45fr 1fr 1fr;}}
  .dt-group {{border: 1px solid rgba(128,128,128,.28); border-top: 4px solid {BRAND}; border-radius: 12px;
             padding: .55rem .5rem .6rem .5rem; min-width: 0;}}
  .dt-group.good {{border-top-color: #2E7D4F;}} .dt-group.warn {{border-top-color: #D4A72C;}}
  .dt-group.bad {{border-top-color: #C2413D;}}
  .dt-gt {{font-size: .82rem; font-weight: 800; color: var(--rc-ink); padding: 0 .35rem .35rem .35rem;
          border-bottom: 1px solid rgba(128,128,128,.18); margin-bottom: .4rem;}}
  .dt-gt span {{font-weight: 600; color: var(--rc-muted); margin-left: .4rem;}}
  .dt-gi {{display: grid;}}
  .dt-gi .dt-tile {{border: none; border-radius: 0; padding: .2rem .45rem; background: none;}}
  .dt-gi .dt-tile + .dt-tile {{border-left: 1px solid rgba(128,128,128,.18);}}
  @media (max-width: 1100px) {{ .dt-groups {{grid-template-columns: 1fr;}} }}
  .dt-tile {{border: 1px solid rgba(128,128,128,.28); border-top: 4px solid {BRAND}; border-radius: 11px;
            padding: .65rem .8rem; min-width: 0;}}
  .dt-tile.good {{border-top-color: #2E7D4F;}} .dt-tile.warn {{border-top-color: #D4A72C;}}
  .dt-tile.bad {{border-top-color: #C2413D; background: rgba(194,65,61,.05);}}
  .dt-l {{font-size: .66rem; font-weight: 800; letter-spacing: .06em; color: var(--rc-muted);}}
  .dt-v {{font-size: 1.2rem; font-weight: 800; color: var(--rc-ink); line-height: 1.25; overflow-wrap: anywhere;}}
  .dt-tile.good .dt-v {{color: var(--rc-green);}} .dt-tile.bad .dt-v {{color: var(--rc-red); font-size: 1.1rem;
          white-space: normal; padding-top: .15rem;}} .dt-tile.warn .dt-v {{color: light-dark(#8A5B00, #F2C14E);}}
  .dt-s {{font-size: .74rem; color: var(--rc-muted); line-height: 1.35;}}
  .dt-targets {{font-size: .86rem; color: var(--rc-ink); border: 1px solid rgba(128,128,128,.28); border-radius: 10px;
               padding: .5rem .75rem; margin: -.3rem 0 .9rem 0; line-height: 1.6;}}
  .st-key-dash_tools button[data-testid="stPopoverButton"] {{border-radius: 999px; border: 1px solid #DCE9EB;
      background: #fff; color: #123239; font-weight: 700; min-height: 2.6rem; width: 100%; padding: 0 .55rem;
      box-shadow: 0 2px 8px rgba(18,50,57,.06);}}
  .st-key-dash_tools [data-testid="stColumn"]:first-child button[data-testid="stPopoverButton"] {{background: #123239;
      color: #99E378; border-color: #123239;}}
  .st-key-dash_tools button[data-testid="stPopoverButton"] div[aria-hidden="true"] {{display: none;}}
  .st-key-dash_tools button[data-testid="stPopoverButton"] p {{white-space: nowrap; font-weight: 700;}}
  .st-key-dash_tools [data-testid="stTooltipHoverTarget"] {{justify-content: stretch !important;}}
  /* the toolbar wraps onto a second line rather than squeezing a label down to "…" on a narrower screen */
  .st-key-dash_tools [data-testid="stHorizontalBlock"] {{flex-wrap: wrap; row-gap: .5rem;}}
  .st-key-dash_tools [data-testid="stColumn"] {{min-width: 44px;}}
  .st-key-dash_tools [data-testid="stColumn"]:nth-child(1) {{min-width: 6.5rem;}}
  .st-key-dash_tools [data-testid="stColumn"]:nth-child(2) {{min-width: 11rem;}}
  .st-key-dash_tools [data-testid="stColumn"]:nth-child(3) {{min-width: 9rem;}}
  .dt-cover {{font-size: .84rem; color: var(--rc-muted); border: 1px solid rgba(128,128,128,.22);
              border-radius: 10px; padding: .45rem .75rem; margin: 0 0 .9rem 0;}}
  .dt-path {{font-size: .85rem; font-weight: 800; color: var(--rc-muted);}}
  .dt-find {{font-size: .86rem; color: var(--rc-ink); background: rgba(43,138,153,.08); border-left: 3px solid #2B8A99;
            border-radius: 6px; padding: .35rem .6rem; margin: .1rem 0 .4rem 0;}}
  .dash-cur {{font-size: .78rem; font-weight: 800; letter-spacing: .05em; color: var(--rc-muted); margin-top: .2rem;}}
  .dash-up {{color: var(--rc-green); font-weight: 700;}}  .dash-down {{color: var(--rc-red); font-weight: 700;}}
  .dash-flat {{font-weight: 700;}}
  .dash-card-title {{font-weight: 800; font-size: 1rem; color: var(--rc-ink); display: flex;
      justify-content: space-between; align-items: baseline; gap: .6rem; margin-bottom: .2rem;}}
  .dash-card-title span {{font-size: .78rem; font-weight: 600; color: var(--rc-muted);}}
  .dash-dot {{display: inline-block; width: .55rem; height: .55rem; border-radius: 50%; margin-right: .5rem;
      flex: none; background: var(--rc-teal);}}
  .dash-dot.good {{background: var(--rc-green);}} .dash-dot.warn {{background: #D4A72C;}}
  .dash-dot.bad {{background: var(--rc-red);}} .dash-dot.info {{background: var(--rc-teal);}}
  .dash-status, .dash-table {{width: 100%; border-collapse: collapse; font-size: .9rem;}}
  .dash-status td, .dash-table td {{padding: .45rem .4rem; border-bottom: 1px solid rgba(128,128,128,.2);}}
  .dash-status td:first-child {{white-space: nowrap;}}
  .dash-status td + td, .dash-table td + td, .dash-table th + th {{text-align: right;}}
  .dash-table th {{font-size: .75rem; font-weight: 800; color: var(--rc-muted); text-align: left;
      padding: .35rem .4rem; border-bottom: 1.5px solid var(--rc-ink);}}
  .dash-table tr.hi td {{background: rgba(212,167,44,.12); font-weight: 700;}}
  .dash-th {{font-size: .75rem; font-weight: 800; color: var(--rc-muted); border-bottom: 1.5px solid var(--rc-ink);
      padding-bottom: .3rem;}}
  .dash-th.r, .dash-td.r {{text-align: right;}}
  .dash-td {{font-size: .9rem; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;}}
  .dash-notes > div {{display: flex; align-items: baseline; padding: .45rem 0; font-size: .9rem;
      border-bottom: 1px solid rgba(128,128,128,.2);}}
  .dash-notes > div:last-child {{border-bottom: none;}}
  /* more than 2 saved reports: the list scrolls inside the card (room on the right for the scroll bar) */
  .st-key-rh_scroll {{padding-right: .7rem; overscroll-behavior: contain; scrollbar-width: thin;
      scrollbar-color: rgba(46,125,79,.45) transparent; scroll-snap-type: y proximity; scroll-padding-top: 1.4rem;
      padding-bottom: 4.5rem;}}             /* room so the last two reports also line up at the top */
  .st-key-rh_scroll [data-testid="stElementContainer"]:has(.rh-item) {{scroll-snap-align: start;}}
  /* the invisible frame below (light theme by default) takes no room on the page */
  [data-testid="stLayoutWrapper"]:has(> .st-key-theme_default), .st-key-theme_default {{position: absolute;
      width: 0; height: 0; overflow: hidden;}}
</style>""", unsafe_allow_html=True)

# Light theme by default: Streamlit starts on "System" (which follows the computer) and keeps the theme chosen in the
# menu (top right) in the browser. On the first visit of a browser, "Light" is stored there once (marked with
# "rc-theme-default") and the page is loaded again; a choice made later in the menu is never changed.
# The frame is drawn on every run (it does nothing after the first visit): if it were drawn only once, every element
# under it would move up one place on the next click and Streamlit would briefly show the old step card twice.
with st.container(key="theme_default"):
    st.iframe("""<script>
      try {
        const page = window.parent, store = page.localStorage;
        if (!store.getItem("rc-theme-default")) {
          store.setItem("rc-theme-default", "light");
          store.setItem("stActiveTheme-" + page.location.pathname + "-v2", JSON.stringify("Light"));
          page.location.reload();
        }
      } catch (e) {}
    </script>""", height=1)


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
state.setdefault("prev_reports", {})     # (bank, period key) -> the previous report for "Compare" (in memory)
if "session_folder" not in state:        # this session's own temporary folder (see services/session_files.py)
    state.session_folder = SessionFolder()
if "report_label" not in state:
    state.report_label = "Week"
if "picked" not in state:                # default: last full week
    _today = pd.Timestamp.today().normalize()
    state.picked = (_today - pd.Timedelta(days=_today.weekday() + 7)).date()


if state.get("toast"):
    st.toast(state.pop("toast"), icon="📄")


# =====================================================================================================================
# 5. SHARED HELPERS (used by several steps and by the side panel)
# =====================================================================================================================
# ---- navigation
def go(step):
    """Show another step."""
    state.step = step
    state.no_tx_preview = False                 # the No-Transaction preview page closes when the step changes
    st.rerun()


# ---- download as Excel or PDF
@st.cache_data(show_spinner=False, max_entries=20)
def pdf_bytes(path: str, mtime: float) -> bytes:          # mtime: a new PDF is made when the report changes
    """The PDF of a report (made once, then kept in memory)."""
    return pdf_report.report_pdf(Path(path))


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


# ---- Google Drive: one connection per browser session, made when first needed
def drive_client():
    """(DriveClient or None, problem in plain words). The settings are read again on each call (cheap), so a
    change in drive_settings.toml is seen without restarting the app."""
    settings = config.load()
    problems = settings.problems()
    if problems:
        return None, " ".join(problems)
    # OAuth: the saved sign-in (token); service account: its key. A new file = connect again.
    secret = settings.token_file if settings.is_oauth else settings.credentials_file
    if not secret.is_file():
        state.drive, state.drive_key = None, None
        if settings.is_oauth:
            return None, "Not signed in to Google Drive (side panel → Google Drive → Sign in with Google)."
        return None, f"No credentials file at {secret}."
    key = (settings.auth, str(secret), secret.stat().st_mtime)
    if state.get("drive_key") != key:
        try:
            state.drive = DriveClient.connect(settings)
            state.drive_key = (settings.auth, str(secret), secret.stat().st_mtime)   # a renewal rewrites the token
            state.drive_email = None
        except DriveError as err:
            state.drive, state.drive_key = None, None
            return None, str(err)
    return state.drive, ""


def local_time(value):
    """A Drive time ('2026-10-09T08:15:00.000Z', UTC) or a Timestamp, in this computer's time, without zone."""
    t = pd.to_datetime(value, errors="coerce", utc=True)
    if pd.isna(t):
        return pd.NaT
    from datetime import datetime
    return t.tz_convert(datetime.now().astimezone().tzinfo).tz_localize(None)


# ---- the run: the report is made in the session's temporary folder, kept in memory, and the file deleted at once
def drive_index_cached(client, root, fresh=False):
    """The list of Reconciliation Reports on Drive, made once and reused (Report History and "Compare" share it),
    so changing the bank or the month does not list Drive again. Saving a report clears it."""
    cached = state.get("history_index")
    if fresh or not cached or cached[0] != root:
        state.history_index = (root, report_history.drive_index(client, root))
        state.history_index_at = pd.Timestamp.now()
    return state.history_index[1]


def drive_index_age() -> float:
    """Seconds since Drive was listed (a large number when it never was)."""
    at = state.get("history_index_at")
    return (pd.Timestamp.now() - at).total_seconds() if at is not None else 1e9


def find_previous(prev):
    """The report of the previous period for "Compare": (info or None, problem). Looked for on Google Drive first,
    then among the older reports in output/; an uploaded one is kept in memory. Nothing is downloaded here."""
    key = (bank.CODE, prev.store_key)
    if key in state.prev_reports:
        return state.prev_reports[key], ""
    problem = ""
    client, _ = drive_client()
    root = config.load().reports_root_folder_id
    missing = state.setdefault("prev_missing", {})          # "not on Drive": asked again after a minute
    if client is not None and root and pd.Timestamp.now() - missing.get(key, pd.Timestamp(0)) > pd.Timedelta(minutes=1):
        try:
            f = report_history.find_previous(client, root, bank, prev, index=drive_index_cached(client, root))
            if f is None and drive_index_age() > 60:  # not in a list older than a minute: list Drive again
                f = report_history.find_previous(client, root, bank, prev,
                                                 index=drive_index_cached(client, root, fresh=True))
            if f is not None:
                state.prev_reports[key] = {"where": "on Google Drive", "name": f.name, "file_id": f.id}
                return state.prev_reports[key], ""
            missing[key] = pd.Timestamp.now()
        except DriveError as err:
            problem = f"Google Drive: {err}"
    path, _ = engine.find_previous_report(BASE_DIR, prev, bank)
    if path is not None:
        return {"where": "on this computer", "name": path.name, "path": path}, ""
    return None, problem


def previous_file(prev, prev_info, session) -> Path:
    """The previous report (find_previous) written in the session's temporary folder, for one run only."""
    if "bytes" in prev_info:
        data = prev_info["bytes"]
    elif "file_id" in prev_info:
        client, problem = drive_client()
        if client is None:
            raise engine.ReconError(f"Google Drive: {problem}")
        try:
            data = client.download(prev_info["file_id"])
        except DriveError as err:
            raise engine.ReconError(f"The previous report could not be downloaded: {err}") from None
    else:
        data = Path(prev_info["path"]).read_bytes()
    path = engine.report_path(BASE_DIR, prev, bank, root=session.path)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(data)
    return path


def previous_counts():
    """(counts, where, problem) of the previous period's report, for the "Previous week" column of the
    No-Transaction Report: counts = {table: {currency: n, "Total": n}}, None when no report was found."""
    session = state.session_folder
    prev = period.previous()
    prev_info, problem = find_previous(prev)
    if prev_info is None:
        return None, "", problem
    path = None
    try:
        path = previous_file(prev, prev_info, session)
        return engine.read_previous_counts(path, period, bank), f"{prev.name} (report {prev_info['where']})", ""
    except engine.ReconError as err:
        return None, "", str(err)
    finally:
        session.remove(path)


def build_no_tx_mem():
    """The memory bundle of a No-Transaction Report for the current bank and period (the Excel bytes, and the
    previous period's counts for its comparison column) - what the preview, Drive save and Slack message use."""
    made_at = pd.Timestamp.now()                # made once: the file downloaded is the one saved on Drive
    with st.spinner(f"Looking for the {bank.NAME} {period.previous().name} report…"):
        previous, compared_with, problem = previous_counts()
    return {"name": f"{bank.CODE}{engine.REPORT_SUFFIX}{period.label}.xlsx", "made_at": made_at,
            "xlsx": engine.no_transaction_excel(period, bank, made_at, previous),
            "previous": previous, "compared_with": compared_with, "problem": problem}


def run_reconciliation(acc):
    """engine.run for the current period, with the report written in the session's temporary folder. Rows not
    found in IremboPay are classified automatically before the report is made (no invoice number, or the
    IremboPay checkout check: services/auto_review.py); state.needs_a_look lists what still deserves a look.
    When a report is made, its Excel and PDF are read into memory (state.report_mem) and the file is deleted;
    the previous report (Compare) is written there only for the run. Errors are raised like engine.run."""
    session = state.session_folder
    res, prev_file, prev_info = None, None, None
    try:
        if compare:
            prev = period.previous()
            prev_info, problem = find_previous(prev)
            if prev_info is not None:                   # not found: engine.run() decides - N/A, or the same error
                prev_file = previous_file(prev, prev_info, session)
        res = engine.run(BASE_DIR, period, acc, sla_days=sla_days, compare=compare, bank=bank,
                         report_dir=session.path, previous_report=prev_file)
        state.needs_a_look = []
        if res.stopped_for_review:
            todo = res.review.loc[~res.review["Reviewed (Yes/No)"].map(engine.is_yes), "invoice_number"]
            with st.spinner("Checking the invoices not found in IremboPay…"):
                answers = irembopay_checkout.check_many(todo.dropna().astype(str))
            filled, state.needs_a_look = auto_review.classify(res.review, answers)
            engine.save_decisions(BASE_DIR, period, filled, bank)
            res = engine.run(BASE_DIR, period, acc, sla_days=sla_days, compare=compare, bank=bank,
                             report_dir=session.path, previous_report=prev_file)
        if prev_info is not None and res.previous_counts is not None:   # not when the export was too thin to trust
            res.compared_with = f"{period.previous().name} (report {prev_info['where']})"
        if res.report_file is not None:
            state.report_mem = {"name": res.report_file.name, "xlsx": res.report_file.read_bytes(),
                                "pdf": pdf_report.report_pdf(res.report_file)}
            state.drive_save = None                     # a new report: not saved on Drive yet
            local_copy = engine.report_path(BASE_DIR, period, bank)      # kept so the Dashboard can read it
            local_copy.parent.mkdir(parents=True, exist_ok=True)         # locally even when Drive is not connected
            shutil.copy2(res.report_file, local_copy)
        return res
    finally:
        if res is not None:
            session.remove(res.report_file)
        session.remove(prev_file)


# ---- Report History (section 8b): the folder map and the report cards
def _esc(text) -> str:
    import html
    return html.escape(str(text))



def _history_file_button(col, rec, client):
    """Download Excel / PDF: a Drive file is fetched only when clicked, then "Save … file" saves it from memory."""
    label = "Excel" if rec.ext == "xlsx" else "PDF"
    key = rec.file_id or str(rec.path)
    if rec.source != "Google Drive":
        col.download_button(f"Download {label}", data=rec.path.read_bytes(), file_name=rec.name,
                            mime=report_history.MIME[rec.ext], width="stretch", key=f"rh_dl_{key}",
                            icon=":material/download:")
        return
    got = state.get("history_file")
    if got and got[0] == rec.file_id:
        col.download_button(f"Save {label} file", data=got[2], file_name=got[1], mime=report_history.MIME[rec.ext],
                            type="primary", width="stretch", key=f"rh_save_{key}", on_click="ignore",
                            icon=":material/save:")
    elif col.button(f"Download {label}", width="stretch", key=f"rh_get_{key}", icon=":material/download:",
                    help="Fetch it from Google Drive, then click 'Save' to keep it"):
        try:
            with st.spinner(f"Downloading {rec.name}…"):
                state.history_file = (rec.file_id, rec.name, client.download(rec.file_id))
            st.rerun()
        except DriveError as err:
            st.error(str(err), icon="⚠️")


def _open_history_report(rec, client):
    """👁 View: a Drive report is read back into memory (no file kept); an older local one is opened from output/."""
    try:
        if rec.source == "Google Drive":
            with st.spinner(f"Opening {rec.name}…"):
                state.view_mem = {**report_history.read_drive_report(client, rec.file_id, rec.name,
                                                                      state.session_folder),
                                  "name": rec.name, "saved": local_time(rec.modified)}
        else:
            state.view_report = str(rec.path)
        st.rerun()
    except (DriveError, engine.ReconError) as err:
        st.error(str(err), icon="⚠️")


_SVG = ('<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" '
        'stroke-linejoin="round">{}</svg>')
UP_SVG = {     # step 2 cards: a bank, a cloud, a file with an upward arrow
    "bank": _SVG.format('<path d="M3 9.5 12 4l9 5.5zM5 10v7M9.7 10v7M14.3 10v7M19 10v7M3.5 20h17"/>'),
    "cloud": _SVG.format('<path d="M7 18.5h10.5a4 4 0 0 0 .4-7.98A6 6 0 0 0 6.3 9.6 4.5 4.5 0 0 0 7 18.5z"/>'),
    "file": _SVG.format('<path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z"/>'
                        '<path d="M14 3v5h5M12 17.5v-6M9.5 14 12 11.5l2.5 2.5"/>')}

TILE_SVG = {   # the icon tile of a report row: a sheet (Excel) or a page (PDF)
    "xlsx": '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" '
            'stroke-linejoin="round"><rect x="3.5" y="3.5" width="17" height="17" rx="2"/><path d="M3.5 9.5h17M9.5 '
            '9.5v11"/></svg>',
    "pdf": '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" '
           'stroke-linejoin="round"><path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z"/>'
           '<path d="M14 3v5h5M9 13h6M9 17h4"/></svg>'}


def _period_text(rec) -> str:
    """'September 2026', 'Week 28/09 – 04/10/2026', 'Tuesday 15 September 2026'."""
    p = engine.make_period(rec.kind, rec.start)
    return {"day": f"{p.start:%A %d %B %Y}", "week": f"Week {p.start:%d/%m} – {p.last_day:%d/%m/%Y}",
            "month": f"{p.start:%B %Y}"}[rec.kind]


def history_list(records, client, on_drive: bool):
    """The saved reports, one row per report (its Excel and PDF together): an icon tile, the bank and period,
    badges, when it was saved and where it is, then full-width buttons (View report / Download Excel / PDF)."""
    items = {}
    for r in records:                                   # newest period first (records are sorted)
        items.setdefault((r.folder, r.stem), {})[r.ext] = r
    for i, ((folder, stem), files) in enumerate(items.items()):
        if i:
            st.divider()
        first = files.get("xlsx") or next(iter(files.values()))
        when = max((local_time(f.modified) if f.source == "Google Drive" else f.modified) for f in files.values())
        bank_name = engine.BANKS[first.bank].FULL_NAME if first.bank in engine.BANKS else first.bank
        badges = "".join(f'<span class="rh-badge{"" if ext == "xlsx" else " pdf"}">{"Excel" if ext == "xlsx" else "PDF"}'
                         f'</span><span class="rh-size">{(files[ext].size or 0) / 1024:,.0f} KB</span>'
                         for ext in ("xlsx", "pdf") if ext in files)
        if first.version > 1:
            badges += f'<span class="rh-badge ver">Version {first.version}</span>'
        saved = f"Saved {when:%d %b %Y} at {when:%H:%M}" if pd.notna(when) else ""
        st.markdown(f'<div class="rh-item"><div class="rh-tile{"" if "xlsx" in files else " pdf"}">'
                    f'{TILE_SVG["xlsx" if "xlsx" in files else "pdf"]}</div><div>'
                    f'<div class="rh-name">{_esc(bank_name)} Reconciliation</div>'
                    f'<div class="rh-period">{_esc(_period_text(first))}</div>'
                    f'<div class="rh-badges">{badges}</div><div class="rh-saved">{saved}</div></div></div>',
                    unsafe_allow_html=True)
        cols = st.columns(1 + len(files))
        if "xlsx" in files and cols[0].button("View report", type="primary", width="stretch",
                                              icon=":material/visibility:",
                                              key=f"rh_view_{first.file_id or first.path}"):
            _open_history_report(files["xlsx"], client)
        for col, ext in zip(cols[1:], [e for e in ("xlsx", "pdf") if e in files]):
            _history_file_button(col, files[ext], client)
        if not on_drive:                                    # older local reports: Delete stays available
            target = str(files.get("xlsx", first).path)
            if st.button("Delete from this computer", type="tertiary", icon=":material/delete:",
                         key=f"rh_ask_{target}"):
                state.confirm = ("report", target)
            if state.get("confirm") == ("report", target):
                st.warning(f"Delete {Path(target).name} from output/? Its exception decisions are kept; "
                           "copies on Google Drive are not touched.", icon="🗑️")
                y, n_ = st.columns(2)
                if y.button("Yes, delete", type="primary", width="stretch", key=f"rh_del_{target}"):
                    try:
                        engine.delete_report(BASE_DIR, Path(target))
                        st.toast(f"Report deleted: {Path(target).name}", icon="🗑️")
                    except engine.ReconError as err:
                        st.toast(str(err), icon="⚠️")
                    state.confirm = None
                    st.rerun()
                if n_.button("Cancel", width="stretch", key=f"rh_cancel_{target}"):
                    state.confirm = None
                    st.rerun()


def drive_link(skey="drive_save"):
    """The Google Drive link of the Excel saved from this page (state[skey]), or None when it is not saved yet."""
    saved = [r for r in (state.get(skey) or {}).get("results", [])
             if r.ok and r.name.lower().endswith(".xlsx") and getattr(r.file, "id", None)]
    return f"https://drive.google.com/file/d/{saved[0].file.id}/view" if saved else None


def auto_delivery(res, mem):
    """Step 6, automatic: the report is saved on Google Drive (a report of the same period already there is replaced:
    Drive keeps its earlier content in the file's version history) and, once Drive has confirmed every file, the
    team message (services/slack_message.build_blocks, with the report's link) is sent to Slack. Each happens once
    per report; what happened is shown, with Retry when something failed and "Send again" for a manual resend."""
    key = (mem["name"], hashlib.md5(mem["xlsx"]).hexdigest())        # this very report (a new run = a new report)
    done = state.setdefault("auto_done", {}).setdefault(key, {})
    client, problem = drive_client()
    settings = config.load()
    root = settings.reports_root_folder_id
    layout = settings.report_layout
    location = report_history.drive_location(bank, period, layout)
    files = {ext: mem[ext] for ext in ("xlsx", "pdf") if mem.get(ext)}

    # ---- 1. Google Drive
    st.markdown("**☁ Google Drive**")
    if client is None or not root:
        st.warning("Not saved: Google Drive is not set up (" + (problem or "no reports folder, reports_root_folder_id")
                   + "). Set it up in Settings → Google Drive, then make the report again.", icon="⚠️")
        return

    def save(only=None):
        try:
            existing = report_history.existing_reports(client, root, bank, period, layout)
            results = report_history.save_report(client, root, bank, period, files,
                                                 "replace" if existing else "new", existing, only, layout)
        except DriveError as err:
            results = [report_history.SaveResult(f"{report_history.report_stem(bank, period)}.{ext}", False,
                                                 error=str(err)) for ext in files]
        old = {r.name: r for r in (state.get("drive_save") or {}).get("results", [])} if only else {}
        old.update({r.name: r for r in results})
        state.drive_save = {"results": list(old.values())}
        state.history_index = None                      # Reports and the Dashboard list Drive again

    if not done.get("save_tried"):
        with st.spinner("Saving the report to Google Drive…"):
            save()
        done["save_tried"] = True
    results = (state.get("drive_save") or {}).get("results", [])
    ok = [r for r in results if r.ok]
    failed = [r for r in results if not r.ok]
    if ok:
        links = " · ".join(f"[{'Excel' if r.name.lower().endswith('.xlsx') else 'PDF'}]"
                           f"(https://drive.google.com/file/d/{r.file.id}/view)" for r in ok if r.file is not None)
        st.markdown(f"✅ Saved in **{location}/** · {links}"
                    + (" · replaced the earlier report" if any(r.replaced for r in ok) else ""))
    for r in failed:
        st.error(f"Not saved: {r.name}. {r.error}", icon="⚠️")
    if failed:
        st.caption("The report is kept in this session: nothing is lost. The Slack message is sent once the report "
                   "is saved.")
        if st.button(f"↻ Retry saving ({len(failed)} file{'s' if len(failed) > 1 else ''})", type="primary",
                     key="auto_save_retry"):
            with st.spinner("Saving the report to Google Drive…"):
                save(only={r.name for r in failed})
            st.rerun()
        return

    # ---- 2. Slack: only once Drive has confirmed every file, so the message has the report's link
    st.markdown("**💬 Slack**")
    url = slack.load_webhook()
    if not url:
        st.info("Not sent: no Slack webhook is set. Add it in Settings → Slack.", icon=":material/settings:")
        return
    links = {("xlsx" if r.name.lower().endswith(".xlsx") else "pdf"): f"https://drive.google.com/file/d/{r.file.id}/view"
             for r in ok if r.file is not None}             # the Excel and the PDF saved on Drive

    def send():
        for native in (True, False):                    # Slack's table; else the same table as text
            payload = slack_message.build_blocks(bank, period, res.blocks, res.merchants, links, res.ran_at, native)
            try:
                slack.send_payload(payload, url)
                done.update(sent_at=pd.Timestamp.now(), payload=payload, error=None)
                return
            except slack.SlackError as err:
                done.update(error=str(err), payload=payload)
                if "invalid_blocks" not in str(err) and "400" not in str(err):
                    return                              # not a format problem: the text table would fail too

    if not done.get("send_tried"):
        with st.spinner("Sending the message to Slack…"):
            send()
        done["send_tried"] = True
    if done.get("sent_at") is not None:
        st.markdown(f"✅ Sent at **{done['sent_at']:%H:%M}** to {slack.masked(url)}")
    elif done.get("error"):
        st.error(f"Not sent: {done['error']}", icon="⚠️")
    if done.get("payload"):
        with st.expander("👁 See the message", expanded=True):
            st.code(slack_message.preview_text(done["payload"]), language=None, wrap_lines=True)
    label = "💬 Send again to Slack" if done.get("sent_at") is not None else "↻ Retry sending"
    if st.button(label, key="auto_send_again", help="Posts the message to the Slack channel again"):
        with st.spinner("Sending the message to Slack…"):
            send()
        st.rerun()


def slack_send_card(text: str, report_name: str, link, note: str = "", key: str = "slack", drive_hint: bool = True):
    """The message shown first, then sent to the channel of the saved webhook on click (the webhook is set in the
    side panel, Slack). `report_name` tells whether this very message was already sent ("Send again")."""
    url = slack.load_webhook()
    if not url:
        st.info("No Slack webhook is set yet. Add it in the side panel: **Slack** (under Settings).",
                icon=":material/settings:")
        return
    st.caption(" · ".join(x for x in (note, f"posts to {slack.masked(url)}") if x))
    if not link and drive_hint:
        st.caption("Save the report to Google Drive first to add its link (\"Here\") to the message.")
    with st.expander("👁 See message"):              # hidden until asked for; it is exactly what is sent
        st.code(text, language=None, wrap_lines=True)
    if state.pop("slack_just_sent", False):          # a short note, gone after 6 s (like "Report saved")
        st.markdown('<div class="prev-found save-note"><b>Message sent to Slack</b> <span>the channel of the saved '
                    'webhook</span></div>', unsafe_allow_html=True)
    already = state.get("slack_sent") == (report_name, text)
    if st.button("💬 Send again to Slack" if already else "💬 Send to Slack", type="primary", key=f"{key}_send",
                 help="Posts the message above to the Slack channel"):
        try:
            with st.spinner("Sending to Slack…"):
                slack.send(text, url)
            state.slack_sent, state.slack_just_sent = (report_name, text), True
            st.rerun()
        except slack.SlackError as err:
            st.error(str(err), icon="⚠️")


def save_to_drive_card(mem, skey="drive_save"):
    """Save the report on Google Drive: the standard Excel + PDF (Results), or the Excel alone of a No-Transaction
    Report (skey "no_tx_drive"). state[skey] remembers where it stands:
    {"ask": True, "existing": [...]}       the report is already on Drive: replace it or save a new version?
    {"mode", "existing", "results": [...]}  what Drive confirmed for each file (failed ones can be retried)
    A file is called saved only when Drive confirmed it (report_history.SaveResult.ok)."""
    client, problem = drive_client()
    settings = config.load()
    root = settings.reports_root_folder_id
    layout = settings.report_layout
    if client is None or not root:
        st.caption("Google Drive is not set up: " + (problem or "no reports folder (reports_root_folder_id)."))
        return
    location = report_history.drive_location(bank, period, layout)
    status = state.get(skey) or {}
    files = {ext: mem[ext] for ext in ("xlsx", "pdf") if mem.get(ext)}
    kinds = " and ".join("Excel" if ext == "xlsx" else "PDF" for ext in files)

    def upload(mode, existing, only=None):
        with st.spinner("Saving to Google Drive…"):
            new = report_history.save_report(client, root, bank, period, files, mode, existing, only, layout)
        results = {r.name: r for r in status.get("results", [])}
        results.update({r.name: r for r in new})
        state.history_index = None                      # Report History lists Drive again
        state[skey] = {"mode": mode, "existing": existing, "results": list(results.values())}
        state.drive_just_saved = True                   # the saved note is shown once, then fades out
        if len(results) == len(files) and all(r.ok for r in results.values()):   # every file confirmed
            state.drive_toast = f"Report saved to Google Drive · {location}/"
        st.rerun()

    if state.get("drive_toast"):                     # shown once, after the rerun; it closes by itself
        st.toast(state.pop("drive_toast"), icon="✅", duration="short")
    results = status.get("results")
    just_saved = state.pop("drive_just_saved", False)
    if results:                                      # after an upload: only what Drive confirmed is called saved
        saved = [r for r in results if r.ok]
        if saved and just_saved:                     # one short note, gone after 6 s (CSS of section 3)
            kinds = " and ".join("Excel" if r.name.lower().endswith(".xlsx") else "PDF" for r in saved)
            st.markdown(f'<div class="prev-found save-note"><b>Report saved to Google Drive</b> '
                        f'<span>{_esc(kinds)} · {_esc(location)}/'
                        + (" · replaced the earlier files" if any(r.replaced for r in saved) else "")
                        + '</span></div>', unsafe_allow_html=True)
        failed = [r for r in results if not r.ok]
        for r in failed:
            st.error(f"Not saved: {r.name}. {r.error}", icon="⚠️")
        if failed:
            st.caption("The report is kept in this session: nothing is lost. Try again when the problem is fixed.")
            if st.button(f"↻ Retry ({len(failed)} file{'s' if len(failed) > 1 else ''})", type="primary",
                         key="drive_retry"):
                upload(status["mode"], status.get("existing", []), only={r.name for r in failed})
        return
    if status.get("ask"):                            # the same report is already on Drive: never overwrite silently
        existing = status["existing"]
        when = max((local_time(f.modified) for f in existing), default=pd.NaT)
        st.warning(f"This report is already on Google Drive ({location}/): "
                   + ", ".join(sorted(f.name for f in existing))
                   + (f". Last saved {when:%d/%m/%Y %H:%M}." if pd.notna(when) else "."), icon="⚠️")
        if status.get("confirm_replace"):            # "Replace" was clicked: say what will change, then ask
            stem = report_history.report_stem(bank, period)
            replaced = sorted(f.name for f in existing if f.name in {f"{stem}.{ext}" for ext in files}) \
                or [f"{stem}.{ext}" for ext in files]
            st.markdown(f'<div class="pay-missing"><b>Replace the report on Google Drive?</b><br>'
                        f'{"<br>".join("• " + _esc(n) for n in replaced)}<br>'
                        f'will get the content of this new run (same names, same folder: {_esc(location)}/). '
                        'Google Drive keeps the earlier content in the file\'s version history.</div>',
                        unsafe_allow_html=True)
            y, n = st.columns(2)
            if y.button("Yes, replace it", type="primary", key="drive_replace_yes", width="stretch"):
                upload("replace", existing)
            if n.button("No, go back", key="drive_replace_no", width="stretch"):
                state[skey] = {**status, "confirm_replace": False}
                st.rerun()
            return
        b1, b2, b3 = st.columns(3)
        if b1.button("Replace the existing report", key="drive_replace", width="stretch"):
            state[skey] = {**status, "confirm_replace": True}    # nothing is replaced before "Yes"
            st.rerun()
        if b2.button(f"Save as a new version (v{report_history.next_version(existing)})", type="primary",
                     key="drive_version", width="stretch"):
            upload("version", existing)
        if b3.button("Cancel", key="drive_cancel", width="stretch"):
            state[skey] = None
            st.rerun()
        return
    st.caption(f"{kinds} {'are' if len(files) > 1 else 'is'} saved in {location}/")
    if st.button("☁ Save to Google Drive", type="primary", key="drive_save_btn"):
        try:
            with st.spinner("Checking Google Drive…"):
                existing = report_history.existing_reports(client, root, bank, period, layout)
        except DriveError as err:
            st.error(f"Google Drive could not be checked: {err}", icon="⚠️")
            return
        if existing:
            state[skey] = {"ask": True, "existing": existing}
            st.rerun()
        upload("new", [])


def pay_file_month(f):
    """(year, month) of an IremboPay file, read from its folder ('I&M / 09/2026') or its name ('…_2026-08.csv',
    'September 2026 …'); None when neither says it."""
    import calendar
    import re
    names = {m.lower(): i for i, m in enumerate(calendar.month_name) if m}
    names.update({m.lower(): i for i, m in enumerate(calendar.month_abbr) if m})
    for text in (f.folder, f.name):
        m = re.search(r"(?<!\d)(0?[1-9]|1[0-2])[/_.-](20\d\d)(?!\d)", text)
        if m:
            return int(m[2]), int(m[1])
        m = re.search(r"(?<!\d)(20\d\d)[/_.-](0?[1-9]|1[0-2])(?!\d)", text)
        if m:
            return int(m[1]), int(m[2])
        m = re.search(r"([A-Za-z]{3,9})[\s_-]*(20\d\d)", text)
        if m and m[1].lower() in names:
            return int(m[2]), names[m[1].lower()]
    return None


def pay_file_title(f) -> str:
    """'September 2026 payments' (or the file name when no month can be read)."""
    import calendar
    ym = pay_file_month(f)
    return f"{calendar.month_name[ym[1]]} {ym[0]} payments" if ym else Path(f.name).stem


def irembopay_from_drive(pay_types):
    """Step 2: pick an IremboPay file in the Drive folder (and its month subfolders). Returns (name, content,
    "payments") when "Add from Drive" is clicked and the download worked, else None. The file is then read and
    checked exactly like an uploaded one; it is kept in memory only."""
    client, problem = drive_client()
    folder = config.load().irembopay_folder_id

    def panel(label, name, path="", note=""):
        st.markdown(f'<div class="up-file"><div class="up-file-label">{_esc(label)}</div>'
                    f'<div class="up-file-name">{_esc(name)}</div>'
                    + (f'<div class="up-file-path">{_esc(path)}</div>' if path else "")
                    + (f'<div class="up-help">{_esc(note)}</div>' if note else "") + "</div>", unsafe_allow_html=True)

    if client is None or not folder:
        panel("Google Drive", "Not available",
              note=problem or "No IremboPay folder is set (irembopay_folder_id in drive_settings.toml).")
        return None
    if state.get("drive_pay_folder") != folder or state.get("drive_pay_list") is None:
        try:
            with st.spinner("Listing the IremboPay files on Google Drive…"):
                state.drive_pay_list = client.list_files(folder, pay_types)
            state.drive_pay_folder = folder
        except DriveError as err:
            state.drive_pay_list = None
            panel("Google Drive", "Could not list the files", note=str(err))
            return None
    # only the files of the bank being reconciled (its folder: "I&M", "Eco Bank", ...), the period's month first;
    # files in no bank folder come last; "Show all files" lists everything (services/pay_files.py)
    all_key = f"drive_pay_all_{bank.CODE}"
    mine, loose, has_folder = pay_files.files_for_bank(state.drive_pay_list, bank, engine.BANKS)
    loose_ids = {f.id for f in loose}
    ym = (period.start.year, period.start.month)

    def order(f):
        m = pay_file_month(f)
        return (m != ym, -(m[0] * 12 + m[1]) if m else 0, f.name.lower())

    if state.get(all_key):
        files = sorted(state.drive_pay_list, key=order)
    else:
        files = sorted(mine, key=order) + sorted(loose, key=order)
        if not mine:
            st.markdown(f'<div class="up-help">No IremboPay file in a <b>{_esc(bank.NAME)}</b> folder of IremboPay '
                        f'File on Google Drive.</div>', unsafe_allow_html=True)
    if len(files) > 6:                                       # many files: a search box first
        search = st.text_input("Search", key="drive_pay_search", placeholder="Search by file name or folder",
                               icon=":material/search:", label_visibility="collapsed").strip().lower()
        files = [f for f in files if search in f"{f.name} {f.folder}".lower()]

    picked = None
    if not files:
        panel("Available file", "No IremboPay file found",
              note=f"Put the export ({', '.join(pay_types)}) in the {bank.NAME} folder of IremboPay File "
                   "(or use \"Show all files\"), then refresh.")
    else:
        # one card per file (a radio list styled as cards); the file of the period being reconciled comes selected
        st.markdown('<div class="pay-pick-title">Select IremboPay file</div><div class="pay-pick-sub">Choose the '
                    'payment report that matches your bank statement period.</div>', unsafe_allow_html=True)
        ids = [f.id for f in files]
        match = [f.id for f in files if pay_file_month(f) == ym]
        if not match:                                        # no file of the period's month: say so, keep the others
            st.markdown(f'<div class="pay-missing"><b>No {_esc(bank.NAME)} IremboPay file for {_esc(period.name)} on '
                        f'Google Drive.</b> The other months are below: choose one only if it covers this period, or '
                        f'add the file to the {_esc(bank.NAME)} folder and refresh, or upload it from this computer.'
                        f'</div>', unsafe_allow_html=True)
        if state.get("drive_pay_pick") not in ids:          # the period's file comes selected; none if there is none
            state.drive_pay_pick = match[0] if match else None
        by_id = {f.id: f for f in files}
        chosen = st.radio("IremboPay file", ids, key="drive_pay_pick", label_visibility="collapsed",
                          format_func=lambda i: f"**{pay_file_title(by_id[i])}**",
                          captions=[f"Folder: {f.folder or 'IremboPay File'}"
                                    + (" (no bank folder)" if f.id in loose_ids else "")
                                    + f"  \nFilename: {f.name}" for f in files])
        picked = by_id.get(chosen)
        if len(files) > 1:                                   # one card is shown at a time
            st.markdown(f'<div class="pay-pick-more">{len(files)} files · scroll to see the next one</div>',
                        unsafe_allow_html=True)
    if state.get(all_key):
        if st.button(f"Only {bank.NAME} files", key="drive_pay_only", icon=":material/filter_list:", width="stretch"):
            state[all_key] = False
            st.rerun()
    elif not mine and st.button("Show all files", key="drive_pay_all", icon=":material/folder_open:", width="stretch",
                                help="Every file of IremboPay File, of every bank"):
        state[all_key] = True
        st.rerun()
    if st.button("Add selected file", type="primary", key="drive_pay_add", width="stretch", disabled=picked is None):
        try:
            with st.spinner(f"Downloading {picked.name}…"):
                return picked.name, client.download(picked.id), "payments"
        except DriveError as err:
            st.error(str(err), icon="⚠️")
    if st.button("Refresh files", key="drive_pay_refresh", icon=":material/refresh:", width="stretch",
                 help="List the files in IremboPay File again"):
        state.drive_pay_list = None
        st.rerun()
    return None


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


def bank_has_folders():
    """True when this bank has a statements and/or exports folder set in folders.toml."""
    spec = config.load_folders().get(bank.CODE) or {}
    return bool(spec.get("statements") or spec.get("exports"))


def found_files_for_period():
    """The statement/export files folders.toml finds for the current bank and period (None: no folder is set,
    so the app falls back to uploading files by hand)."""
    if not bank_has_folders():
        return None
    client, _ = drive_client()
    return folder_files.find_files(period, bank, config.load_folders(), drive=client)


def _folder_field(label: str, key: str, default: str, kind: str):
    """A text box with a Browse button next to it. Browse opens the computer's own folder window and saves the
    folder it picked for this bank at once (config.save_one_folder: the other folder, if already set, is never
    touched) - a short "Saved" note confirms it. The box can still be typed or pasted into directly; that still
    needs "Save folder settings" below, same as before."""
    st.text_input(label, value=default, key=key)
    if state.pop(f"{key}_just_saved", False):
        st.markdown('<div class="prev-found save-note"><b>Saved</b></div>', unsafe_allow_html=True)
    # its own row, not a narrow sub-column: a column just wide enough for "Browse..." can crop the button's own
    # label with an ellipsis (the same truncation the Dashboard toolbar had)
    if st.button("📂 Browse…", key=f"{key}_browse"):
        with st.spinner("Waiting for the folder window…"):
            path, error = folder_picker.browse_folder(f"Choose the {label}", state.get(key, default))
        if path:
            state[key] = path
            config.save_one_folder(bank.CODE, kind, path)
            state[f"{key}_just_saved"] = True
            st.rerun()
        elif error:
            st.error(error + " Type or paste the path above instead.", icon="⚠️")


def folder_settings_editor():
    """This bank's statements/exports folder (Choose): Browse to one - saved at once, no extra step - or type a
    local path, or "drive:<folder ID>" for a Google Drive folder, then "Save folder settings". Shown open while
    a folder is still missing; once both are set, a collapsed expander - so they can always be seen or changed,
    not only the first time."""
    spec = config.load_folders().get(bank.CODE) or {}
    cur_stm = spec.get("statements").value if spec.get("statements") else ""
    cur_exp = spec.get("exports").value if spec.get("exports") else ""
    stm_key, exp_key = f"folder_stm_input_{bank.CODE}", f"folder_exp_input_{bank.CODE}"
    configured = bool(cur_stm) and bool(cur_exp)

    def fields():
        f1, f2 = st.columns(2)
        with f1:
            _folder_field(f"{bank.FULL_NAME} statements folder", stm_key, cur_stm, "statements")
        with f2:
            _folder_field("IremboPay exports folder", exp_key, cur_exp, "exports")
        if st.button("💾 Save folder settings", key="folder_save"):
            config.save_folders(bank.CODE, state.get(stm_key, cur_stm), state.get(exp_key, cur_exp))
            st.toast(f"Folders saved for {bank.FULL_NAME}", icon="📁")
            st.rerun()

    if configured:
        with st.expander(f"📁 {bank.FULL_NAME} folders"):
            st.caption(f"Statements: {cur_stm}  \nIremboPay exports: {cur_exp}")
            fields()
    else:
        with st.container(border=True):
            st.markdown(f'<div class="card-title">📁 Set the {bank.FULL_NAME} folders</div>', unsafe_allow_html=True)
            st.caption("Not every folder is set yet for this bank, so the app cannot find all its files by itself. "
                       "Browse to each folder below (or type/paste the path, or \"drive:<folder ID>\" for a Google "
                       "Drive folder) - or use a file instead, below, without setting anything.")
            fields()


def render_upload_section():
    """"Use a file instead" (Choose): the bank statements + the IremboPay payment file, read in memory and never
    saved. Returns the files accepted for the current period (used by Check instead of a found folder)."""
    acc, rej = accepted_files()
    # ---- what to upload (depends on the bank)
    monday = engine.extra_monday(period, bank)
    need_to = monday or period.last_day
    day_before = " and ".join(bank.DAY_BEFORE_CURRENCIES)
    per_currency = engine.option(bank, "ONE_FILE_PER_CURRENCY", True)
    pay_types = engine.option(bank, "PAYMENT_FILE_TYPES", ["csv"])
    pay_label = ", ".join(f".{t}" for t in pay_types)
    types = ", ".join(f".{t}" for t in bank.FILE_TYPES)
    st.markdown(f'<div class="card-title">📎 Use a file instead: upload for {period.name}</div>', unsafe_allow_html=True)

    # ---- the two cards; each new file is read at once and listed below
    u1, u2 = st.columns(2)
    with u1.container(border=True, key="up_bank_card"):        # 1. bank statements: from this computer
        st.markdown(f'<div class="up-head"><div class="up-tile">{UP_SVG["bank"]}</div><div>'
                    f'<div class="up-title">1. Bank statement</div><div class="up-sub">{_esc(bank.FULL_NAME)}</div>'
                    f'</div></div>', unsafe_allow_html=True)
        bank_files = sorted((f for f in acc if f.kind == "bank"), key=lambda x: (x.currency != "RWF", x.currency, x.first))
        if bank_files:                 # the files added fill the upload area: one card each, ✕ removes it
            with st.container(height=380, border=False, key="up_bank_files"):
                for row in range(0, len(bank_files), 2):
                    for col, f in zip(st.columns(2, gap="small"), bank_files[row:row + 2]):
                        with col.container(border=True, key=f"upf_{f.digest}"):
                            days = (f"{f.first:%d/%m}" if f.first.date() == f.last.date()
                                    else f"{f.first:%d/%m} – {f.last:%d/%m}")
                            cur_cls = "usd" if f.currency == "USD" else "rwf"
                            st.markdown(f'<div class="upf" title="{_esc(f.name)}"><div class="upf-ic {cur_cls}">'
                                        f'{_esc(f.name.rsplit(".", 1)[-1].upper()[:4])}</div><div class="upf-t">'
                                        f'<div class="upf-d">{days}</div><div class="upf-n"><span class="upf-c {cur_cls}">'
                                        f'{_esc(f.currency or "")}</span> {f.rows:,} rows</div></div></div>',
                                        unsafe_allow_html=True)
                            if st.button("✕", key=f"upf_x_{f.digest}", type="tertiary"):
                                state.files.pop(f.digest, None)
                                state.result = None
                                state.uploader_round += 1
                                st.rerun()
            st.markdown(f'<div class="upf-foot">{len(bank_files)} file{"s" if len(bank_files) != 1 else ""}'
                        + (" · scroll to see them all" if len(bank_files) > 10 else "") + " · ✕ removes one</div>",
                        unsafe_allow_html=True)
        else:
            # the upload area on its own: it grows so this card has no empty space under its text (section 3 CSS)
            st.markdown(f'<div class="up-drop up-drop-fill">{UP_SVG["file"]}<div class="up-drop-title">Upload bank file'
                        f'</div><div class="up-drop-sub">{"Excel" if set(bank.FILE_TYPES) <= {"xls", "xlsx"} else "Bank"}'
                        f' format · {types}</div></div>', unsafe_allow_html=True)
        bank_up = st.file_uploader(f"{bank.NAME} bank files ({types})", type=bank.FILE_TYPES,
                                   accept_multiple_files=True, key=f"bank_{state.uploader_round}",
                                   label_visibility="collapsed")
        help_lines = [f"Upload the {' and '.join(bank.CURRENCIES) + ' ' if per_currency else ''}statements "
                      f"covering <b>{period.start:%d/%m} to {need_to:%d/%m/%Y}</b>."]
        if monday is not None:
            help_lines.append(f"Include <b>Monday {monday:%d/%m}</b>: late weekend payments are posted on Monday.")
        if day_before:
            help_lines.append(f"Recommended: the {day_before} statement of <b>{period.day_before:%d/%m/%Y}</b> "
                              "(the day before).")
        # once files are added, only their cards and the button stay (an empty block keeps the button's place, so
        # it stays level with "Add selected file" of the other card)
        st.markdown(f'<div class="up-help">{"<br>".join(help_lines) if not bank_files else ""}</div>',
                    unsafe_allow_html=True)

    pay_up, from_drive = None, None
    with u2.container(border=True, key="up_pay_card"):         # 2. IremboPay file: Google Drive or this computer
        st.markdown(f'<div class="up-head"><div class="up-tile">{UP_SVG["cloud"]}</div><div>'
                    f'<div class="up-title">2. IremboPay file</div><div class="up-sub">Choose a source</div>'
                    f'</div></div>', unsafe_allow_html=True)
        drive_ready = drive_client()[0] is not None and bool(config.load().irembopay_folder_id)
        if "pay_source" not in state:                          # Google Drive first when it is set up
            state.pay_source = "Google Drive" if drive_ready else "This computer"
        source = st.segmented_control("IremboPay file from", ["Google Drive", "This computer"], key="pay_source",
                                      label_visibility="collapsed") or "Google Drive"
        if source == "This computer":
            st.markdown(f'<div class="up-drop">{UP_SVG["file"]}<div class="up-drop-title">Upload IremboPay file</div>'
                        f'<div class="up-drop-sub">{pay_label.replace(".", "").upper()} format · {pay_label}</div>'
                        '</div>', unsafe_allow_html=True)
            pay_up = st.file_uploader(f"IremboPay payment file ({pay_label})", type=pay_types,
                                      accept_multiple_files=True, key=f"pay_{state.uploader_round}",
                                      label_visibility="collapsed")
            st.markdown(f'<div class="up-help">The IremboPay payment file for <b>{period.name}</b>.</div>',
                        unsafe_allow_html=True)
        else:
            from_drive = irembopay_from_drive(pay_types)

    with st.container(border=bool(acc or rej or state.read_errors)):   # what was added, and what was refused
        new = ([(u.name, u.getvalue(), "bank") for u in bank_up or []]
               + [(u.name, u.getvalue(), "payments") for u in pay_up or []] + ([from_drive] if from_drive else []))
        if new:
            with st.spinner("Reading the files…"):
                for name, content, kind in new:
                    digest = engine.file_digest(content)
                    if digest in state.files:
                        state.read_errors.append(f"{name}: the same file was already added.")
                        continue
                    try:
                        if kind == "bank":
                            state.files[digest] = engine.read_bank_file(name, content, bank)
                        else:
                            state.files[digest] = engine.read_payment_file(name, content, bank)
                            state.toast = f"IremboPay file added: {name}"      # a short pop-up after the reload
                    except engine.ReconError as err:
                        state.read_errors.append(str(err))
                    except Exception as err:
                        state.read_errors.append(f"{name}: could not be read ({err}).")
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
    return acc


def files_for_run():
    """The files to reconcile with: folder-found files when the bank has folders, else the uploaded ones."""
    if bank_has_folders():
        found = found_files_for_period()
        return folder_files.all_files(found) if found is not None else []
    acc, _ = accepted_files()
    return acc


# ---- key figures + summary tables (Results, and a saved report)
def _count(tables, title, cur=None):
    """(transactions, amount, %) of one table: the total row, or the row of one currency."""
    t = tables[title]
    row = t[t["Currency"].isna()] if cur is None else t[t["Currency"].eq(cur)]
    return int(row["Number of Transactions"].iloc[0]), row["Total Amount"].iloc[0], (row["%"].iloc[0] if "%" in t else None)


def render_kpis(tables):
    """The 5 key figures at the top: processed, received %, SLA violated, pending, duplicate (the rows the rules
    decided automatically are already in the table their decision chose: they are not counted again)."""
    processed, _, _ = _count(tables, "Overall processed payments")
    rec_n, _, rec_pct = _count(tables, "Overall notifications received by Irembo")
    sla_n, _, sla_pct = _count(tables, "SLA Violated")
    t = tables["Overall processed payments"]
    per_currency = " · ".join(f"{cur} {int(n):,} ({0 if pd.isna(amt) else amt:,.0f})"
                              for cur, n, amt in zip(t["Currency"], t["Number of Transactions"], t["Total Amount"])
                              if pd.notna(cur))
    k = st.columns(5, border=True)
    k[0].metric("Processed", f"{processed:,}", help=per_currency)
    k[1].metric("Received", f"{rec_pct:.2%}", help=f"{rec_n:,} notifications received by Irembo")
    k[2].metric("SLA violated", f"{sla_n:,}", help=f"{sla_pct:.2%} of processed")
    k[3].metric("Pending (NEW)", f"{_count(tables, 'Pending notifications')[0]:,}")
    k[4].metric("Duplicate", f"{_count(tables, 'Duplicate Payment')[0]:,}", help="Paid with another channel")


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


def _show_merchants(table):
    """Payment per merchant under a block's tables (reports made before 12/10/2026 have none: nothing is shown)."""
    if table is None or not len(table):
        return
    t = table.copy()
    t["%"] = t["%"] * 100
    st.markdown("**Payment per merchant**")
    config = {"Merchant name": st.column_config.TextColumn("Merchant"),
              "Number of Transactions": st.column_config.NumberColumn("Transactions", format="localized"),
              "Total Amount": st.column_config.NumberColumn("Total amount", format="localized"),
              "%": st.column_config.NumberColumn("% of received", format="%.2f%%")}
    st.dataframe(t, hide_index=True, width="stretch", column_config=config,
                 height=min(38 + 35 * len(t), 420))


def render_summary(blocks, report_type, merchants=None):
    """The summary: one tab for the month + one per week (a day or a week: a single summary); under each, the
    Payment per merchant table."""
    merchants = merchants or {}
    main = next((b for b in blocks if b[0] in ("Summary-Monthly", "Summary-Daily")), blocks[0])
    week_blocks = [b for b in blocks if b[0] == "Summary-Weekly"]
    st.markdown("#### Summary")
    if report_type == "month" and week_blocks:
        month_name = main[1].split(":")[1].split("(")[0].strip().split(" ")[0] if ":" in main[1] else "Month"
        tabs = st.tabs([f"📅 {month_name}"] + [h.split(":")[0] for _, h, _ in week_blocks])
        with tabs[0]:
            _show_tables(main[2])
            _show_merchants(merchants.get(main[1]))
        for tab, (_, heading, tbls) in zip(tabs[1:], week_blocks):
            with tab:
                st.caption(heading)
                _show_tables(tbls)
                _show_merchants(merchants.get(heading))
    else:
        st.caption(main[1])
        _show_tables(main[2])
        _show_merchants(merchants.get(main[1]))


# ---- short figures (Results)
def short_amount(value):
    """68,166,357 -> '68.2M', 470,064 -> '470K', 950 -> '950'."""
    if abs(value) >= 1e9:
        return f"{value / 1e9:.1f}B"
    if abs(value) >= 1e6:
        return f"{value / 1e6:.1f}M"
    if abs(value) >= 1e3:
        return f"{value / 1e3:.0f}K"
    return f"{value:,.0f}"


# ---- custom report (Results): the choices are widgets with keys "cr_..."; a report type fills them in
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
# 5b. PAGES (the menu at the top of the side panel): Dashboard first, then the reconciliation, the saved reports,
#     the banks' settings and the app's settings. Settings changed in the app are kept in app_settings.json.
# =====================================================================================================================
SETTINGS = app_settings.load()
# the targets set in Settings: the Dashboard's colours and findings, and the 🟢 / 🔴 of the Slack message
dashboard.TARGET_RECEIPT = slack_message.SUCCESS_OK = SETTINGS["targets"]["receipt"]
dashboard.TARGET_ONTIME = slack_message.SLA_OK = SETTINGS["targets"]["ontime"]
sla_days = app_settings.sla_days(SETTINGS, bank)       # the SLA days of the bank chosen (Bank Management)


def _page_dashboard(): pass
def _page_reconciliation(): pass
def _page_reports(): pass
def _page_banks(): pass
def _page_settings(): pass


PAGES = {"dashboard": st.Page(_page_dashboard, title="Dashboard", icon=":material/grid_view:", url_path="dashboard"),
         "reconciliation": st.Page(_page_reconciliation, title="Reconciliation", icon=":material/assignment_turned_in:",
                                   url_path="reconciliation", default=True),       # the page that opens first
         "reports": st.Page(_page_reports, title="Reports", icon=":material/filter_none:", url_path="reports"),
         "banks": st.Page(_page_banks, title="Bank Management", icon=":material/account_balance:",
                          url_path="bank-management"),
         "settings": st.Page(_page_settings, title="Settings", icon=":material/settings:", url_path="settings")}
# the side menu in two parts: the daily work, and the settings
_current = st.navigation({"Workspace": [PAGES["reconciliation"], PAGES["dashboard"], PAGES["reports"]],
                          "Configuration": [PAGES["banks"], PAGES["settings"]]})
PAGE = next(k for k, pg in PAGES.items() if pg.title == _current.title)
if state.get("goto") in PAGES and state.goto != PAGE:    # a page asked for (e.g. a test, or a link in the app)
    st.switch_page(PAGES[state.pop("goto")])
state.pop("goto", None)
if state.get("on_page") != PAGE:                        # another page: a saved report open on the last one closes
    state.on_page, state.view_mem, state.view_report = PAGE, None, None


def current_user() -> str:
    """The Google account signed in to Drive in the app ('' = not signed in): it decides who may change settings."""
    client, _ = drive_client()
    if client is not None and not state.get("drive_email"):
        try:
            state.drive_email = client.account_email()
        except DriveError:
            return ""
    return (state.get("drive_email") or "") if client is not None else ""


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
  if PAGE == "reconciliation":
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
    st.caption(f"SLA ({bank.NAME}): "
               + (f"{sla_days} day{'s' if sla_days != 1 else ''} allowed between the bank "
                  f"{bank.COLUMNS['payment_date']} and paid_at" if engine.option(bank, "SLA_DAYS_SETTING", True)
                  else f"the bank {bank.COLUMNS['payment_date']} and paid_at must be in the same minute")
               + " · set in Bank Management")

# =====================================================================================================================
# 7. BANNER AND STEP BAR (hidden while a saved report or Report History is open)
# =====================================================================================================================
viewing = bool(state.get("view_report") or state.get("view_mem"))
if not viewing:
    _h = lambda s: s.replace("&", "&amp;")
    _logo = base64.b64encode((APP_DIR / "assets" / "irembopay_wordmark.png").read_bytes()).decode()
    st.markdown(f"""<div class="app-banner"><div class="brand"><img src="data:image/png;base64,{_logo}" alt="IremboPay">
<h1>{ {'dashboard': 'Reports Dashboard', 'reports': 'Reports', 'banks': 'Bank Management',
        'settings': 'Settings'}.get(PAGE, _h(bank.FULL_NAME) + ' Reconciliation')}</h1></div></div>""",
                unsafe_allow_html=True)
if not viewing and PAGE == "reconciliation":
    st.markdown('<div class="steps">' + "".join(
        f'<span class="step {"done" if i + 1 < state.step else "now" if i + 1 == state.step else ""}">'
        f'{"✓ " if i + 1 < state.step else ""}{i + 1}. {name}</span>' for i, name in enumerate(STEPS)) + "</div>",
                unsafe_allow_html=True)

# the period chosen in step 1 (used by every step)
period = engine.make_period((state.report_label or "Week").lower(), state.picked)
compare = state.compare


# =====================================================================================================================
# 8. A SAVED REPORT (read only, opened from Report History; the steps are hidden meanwhile)
#    view_mem    a report read back from Google Drive: kept in memory only (no file on this computer)
#    view_report the path of an older report still in output/
# =====================================================================================================================
if viewing:
    if state.get("view_mem"):
        vm = state.view_mem
        rep, where, made = vm["report"], "Google Drive", vm.get("saved")
    else:
        vpath = Path(state.view_report)
        try:
            if not vpath.exists():
                raise engine.ReconError("This report no longer exists.")
            rep = engine.read_report(vpath)
        except engine.ReconError as err:
            state.view_report = None
            st.error(str(err), icon="⚠️")
            st.stop()
        where, made = "this computer", pd.Timestamp(vpath.stat().st_mtime, unit="s")
    vmain = next((b for b in rep["blocks"] if b[0] in ("Summary-Monthly", "Summary-Daily")), rep["blocks"][0])
    when = (f"Reconciled on {rep['reconciled_on']}" if rep["reconciled_on"]          # older reports: the file time
            else f"Made on {made:%d/%m/%Y at %H:%M}" if made is not None and pd.notna(made) else "")
    st.markdown(f"""<div class="view-banner"><span class="tag">SAVED REPORT · READ ONLY · {where.upper()}</span>
<h2>📄 {rep["bank"].replace("&", "&amp;")} · {rep["title"]}</h2><p>{when}</p></div>""",
                unsafe_allow_html=True)

    def close_view():
        state.view_report = state.view_mem = None              # the report read from Drive leaves memory
        state.toast = f"Closed {rep['title']}"
        st.rerun()

    if st.button("←  Back", type="primary", key="view_back"):
        close_view()
    with st.container(border=True):
        render_kpis(vmain[2])
        v1, v2 = st.columns([3, 1], vertical_alignment="bottom")
        with v1:
            if state.get("view_mem"):
                fmt = st.segmented_control("Format", ["Excel", "PDF"], default="Excel", key="fmt_view_mem",
                                           label_visibility="collapsed") or "Excel"
                ext = "xlsx" if fmt == "Excel" else "pdf"
                st.download_button(f"⬇ Download report ({fmt})", data=vm[ext], file_name=f"{Path(vm['name']).stem}.{ext}",
                                   mime=report_history.MIME[ext], type="primary", width="stretch", key="dl_view_mem")
            else:
                download_control(vpath, "view", label="⬇ Download report", primary=True)
        if v2.button("✖  Close", width="stretch", key="view_close"):
            close_view()
    render_summary(rep["blocks"], rep["report_type"], rep.get("merchants"))
    st.stop()                                   # nothing else is drawn while a saved report is open


# =====================================================================================================================
# 8c. REPORTS DASHBOARD (side panel "Dashboard"): the monthly reports saved on Google Drive, side by side
#     Only the numbers saved in the reports are read (services/dashboard.py); nothing is reconciled again. The
#     reports are downloaded once per session (then "Refresh"); the "Compare" option of step 1 is not changed.
# =====================================================================================================================

# I&M dark teal, Ecobank green, BOA light teal, BPR purple
BANK_COLORS = {"IM": BRAND, "ECO": "#5FAF3C", "BOA": "#2B8A99", "BPR": "#8E6FB5"}


def bank_color(code):
    """The colour of a bank in the dashboard charts (I&M is lightened on the dark theme, where dark teal vanishes)."""
    if code == "IM" and st.context.theme.type == "dark":
        return "#A9D6DC"
    return BANK_COLORS.get(code, "#8A9A9F")


def _short(v: float) -> str:
    """12,345,678 -> '12.3M' (the key figure cards); the full number is in the tooltip."""
    for size, unit in ((1e9, "B"), (1e6, "M"), (1e3, "K")):
        if abs(v) >= size:
            return f"{v / size:,.1f}{unit}"
    return f"{v:,.0f}"


def _delta(ch, *, pct=True, lower_is_better=False, points=False):
    """'▲ 4.1%' in green or red (red when the change is bad), or '' when there is nothing to compare with."""
    if ch is None:
        return ""
    diff, rel = ch
    if diff == 0:
        return '<span class="dash-flat">no change</span>'
    up = diff > 0
    good = up != lower_is_better
    if points:
        text = f"{abs(diff) * 100:.2f} pt"
    elif pct and rel is not None:
        text = f"{abs(rel):.1%}"
    else:
        text = f"{abs(diff):,.0f}"
    return f'<span class="{"dash-up" if good else "dash-down"}">{"▲" if up else "▼"} {text}</span>'


def dashboard_facts(client, root):
    """The facts of every monthly/weekly report, from Google Drive (when connected: client and root are set) and
    from the reports still kept locally in output/, merged without duplicates (services/dashboard.report_records:
    the same bank, kind and period counted once; Drive's copy wins a tie). Each report is read once per session.
    -> (facts to use, [problems])."""
    drive_records = []
    if client is not None and root:
        cached = state.get("history_index")
        if not cached or cached[0] != root:
            with st.spinner("Reading Reconciliation Reports on Google Drive…"):
                state.history_index = (root, report_history.drive_index(client, root))
                state.history_index_at = pd.Timestamp.now()
        drive_records = state.history_index[1].records
    local_records = report_history.local_records(BASE_DIR, None)
    records = [r for r in dashboard.report_records(drive_records + local_records)
               if r.kind in ("month", "week")]           # weekly reports are used too (dashboard.combine sorts it out)
    cache = state.setdefault("dash_facts", {})
    # a replaced report (same file, new content) is read again; local and Drive copies never share a key
    key = lambda r: (r.source, r.file_id or str(r.path), str(r.modified))
    todo = [r for r in records if key(r) not in cache]
    if todo:
        bar = st.progress(0.0, text=f"Reading {len(todo)} report{'s' if len(todo) > 1 else ''}…")
        for i, r in enumerate(todo, 1):
            try:
                cache[key(r)] = (report_history.read_drive_facts(client, r, state.session_folder)
                                 if r.source == "Google Drive" else report_history.read_local_facts(r))
            except (DriveError, engine.ReconError, KeyError, ValueError, OSError) as err:
                cache[key(r)] = f"{r.name}: {err}"
            bar.progress(i / len(todo), text=f"Reading reports… {i} of {len(todo)}")
        bar.empty()
    facts = [f for r in records if isinstance(cache.get(key(r)), list) for f in cache[key(r)]]
    return dashboard.combine(facts), [cache[key(r)] for r in records if isinstance(cache.get(key(r)), str)]


STATUS_COLOR = {"good": "#2E7D4F", "warn": "#D4A72C", "bad": "#C2413D", "none": "#9AA8AD", "": BRAND}
STATUS_ICON = {"Off target": "🔴", "Watch": "🟠", "On target": "🟢", "No report": "⚪"}


def _seg(label, options, key, default=None, fmt=None, help=None, wrap=False):
    """A segmented control that always has a value (the default when nothing is chosen); wrap=True: pills, which
    go on a second line when the card is narrow."""
    default = options[0] if default is None else default
    if state.get(key) not in options:
        state[key] = default
    if wrap:
        return st.pills(label, options, key=key, format_func=fmt, help=help) or default
    return st.segmented_control(label, options, key=key, format_func=fmt, help=help) or default


def tile_groups(tiles) -> str:
    """The key figures in three groups: what happened (payments, amounts), against the targets (notified,
    on time), to follow up (money at risk, needs attention). A group's top line takes its worst colour."""
    groups = [("📦", "Activity", "what happened", [t for t in tiles if t[0].startswith(("PAYMENTS", "AMOUNT"))]),
              ("🎯", "Against the targets", "how well", [t for t in tiles if t[0] in ("NOTIFIED", "ON TIME (SLA)")]),
              ("🔎", "To follow up", "what needs action", [t for t in tiles if t[0] in ("MONEY AT RISK",
                                                                                         "NEEDS ATTENTION")])]
    rank = {"bad": 3, "warn": 2, "good": 1, "": 0}
    html = '<div class="dt-groups">'
    for icon, title, sub, items in groups:
        worst = max((t[3] for t in items), key=lambda k: rank.get(k, 0), default="")
        html += (f'<div class="dt-group {worst}"><div class="dt-gt">{icon} {title}<span>{sub}</span></div>'
                 f'<div class="dt-gi" style="grid-template-columns: repeat({max(len(items), 1)}, minmax(0, 1fr))">'
                 + "".join(f'<div class="dt-tile {k}"><div class="dt-l">{_esc(a)}</div><div class="dt-v">{_esc(b)}</div>'
                           f'<div class="dt-s">{_esc(c)}</div></div>' for a, b, c, k in items) + "</div></div>")
    return html + "</div>"


def dashboard_page():
    st.markdown("<style>.block-container {max-width: 1600px;}</style>", unsafe_allow_html=True)   # this page is wide
    client, problem = drive_client()
    root = config.load().reports_root_folder_id
    drive_on = client is not None and bool(root)
    try:
        facts, errors = dashboard_facts(client if drive_on else None, root if drive_on else None)
    except DriveError as err:
        drive_on = False
        facts, errors = dashboard_facts(None, None)
        errors = [f"Google Drive: {err}"] + errors
    if not drive_on:
        st.caption(f"☁️ Google Drive not connected ({problem or 'no reports folder set'}) - showing the reports "
                   f"kept locally in {_esc(BASE_DIR / 'output')} only.")
    for e in errors:
        st.warning(f"Left out (could not be read): {e}", icon="⚠️")
    if not facts:
        st.markdown(f'<div class="rh-empty">📊 No monthly or weekly report found.<br>Save one with '
                    f'<b>☁ Save to Google Drive</b>, or check that reports exist in '
                    f'<code>{_esc(BASE_DIR / "output")}</code>.</div>', unsafe_allow_html=True)
        return

    # ---- the filters (pills at the top right): they apply to every figure, chart and table of the page.
    # Their labels show the current choice, so the choice is read from the session before the widgets are drawn.
    if "dash_pick" in state:                       # a bar clicked in "by bank": the bank filter follows it
        state.dash_bank = state.pop("dash_pick")
    if "dash_goto" in state:                       # a level chosen in the breakdown or the path: (level, period)
        level, key = state.pop("dash_goto")
        state.dash_kind, state[f"dash_period_{level}"] = level, key
    all_codes = list(engine.BANKS)
    # the hierarchy, all from the monthly reports: a year = its quarters = its months; a month holds its weeks and days
    kinds = {"year": "Year", "quarter": "Quarter", "month": "Month", "week": "Week", "day": "Day"}
    if state.get("dash_bank") not in ["All"] + all_codes:
        state.dash_bank = "All"
    if state.get("dash_kind") not in kinds:
        state.dash_kind = "month"
    compare_modes = {"previous": "The period before", "year": "Same period last year",
                     "pick": "Another period I choose", "none": "No comparison (this period only)"}
    if state.get("dash_compare") not in compare_modes:
        state.dash_compare = "none"                        # by default: this period only
    pick, kind, mode = state.dash_bank, state.dash_kind, state.dash_compare
    with_partial = bool(state.get("dash_partial", False))
    pick_banks = all_codes if pick == "All" else [pick]
    options = (dashboard.period_options([f for f in facts if f.bank in pick_banks], kind) if with_partial
               else dashboard.complete_options(facts, kind, pick_banks))
    if options and state.get(f"dash_period_{kind}") not in options:
        state[f"dash_period_{kind}"] = options[0]
    first = state.get(f"dash_period_{kind}") if options else None
    last = None
    # "Another period I choose": any other saved period of the same type (default: the one just before)
    vs_options = [o for o in options if o != first]
    if mode == "pick" and vs_options and state.get(f"dash_vs_{kind}") not in vs_options:
        before_default = dashboard.previous_period(kind, first, last, "previous")
        state[f"dash_vs_{kind}"] = (before_default[0] if before_default and before_default[0] in vs_options
                                    else vs_options[0])
    if not options:
        before_span = None
    elif mode == "pick":
        before_span = (state[f"dash_vs_{kind}"], None) if vs_options else None
    else:
        before_span = dashboard.previous_period(kind, first, last, mode)
    label = lambda k: dashboard.period_label(k, kind) + (
        "" if dashboard.is_complete(facts, kind, k, pick_banks) else " · partial")
    found = sorted({c for f in facts for m in f.counts.values() for c in m}, key=lambda c: (c != "RWF", c))

    hero_col, tools_col = st.columns([1, 1.3], vertical_alignment="top")
    with tools_col.container(key="dash_tools"):
        b1, b2, b3, b4, b5, b6 = st.columns([1.3, 3.3, 1.7, .6, .6, .6], gap="small")
        with b1.popover(kinds[kind], width="stretch", help="Year, quarter, month, week or day"):
            _seg("Period type", list(kinds), "dash_kind", default="month", fmt=kinds.get, wrap=True)
        period_text = (f"📅 {dashboard.describe(kind, first)}"
                       + (f" vs {dashboard.describe(kind, *before_span)}" if before_span else "")) if options else "📅 Period"
        with b2.popover(period_text, width="stretch", help="The period shown and the one it is compared with"):
            if options:
                st.selectbox("Period", options, key=f"dash_period_{kind}", format_func=label)
            st.selectbox("Compare with", list(compare_modes), key="dash_compare", format_func=compare_modes.get)
            if mode == "pick":
                if vs_options:
                    st.selectbox("Compared with", vs_options, key=f"dash_vs_{kind}", format_func=label)
                else:
                    st.caption("No other saved period to compare with.")
            st.toggle("Partial", key="dash_partial",
                      help="Also offer the periods whose reports are not all saved (e.g. a week that "
                           "starts in a month with no report). Off: only complete periods.")
        with b3.popover("🏦 " + ("All banks" if pick == "All" else engine.BANKS[pick].NAME), width="stretch",
                        help="The bank (or all banks)"):
            st.selectbox("Bank", ["All"] + all_codes, key="dash_bank",
                         format_func=lambda c: "All banks" if c == "All" else engine.BANKS[c].FULL_NAME)
        with b4.popover("⚙", width="stretch", help="Currency, refresh, targets, reset"):
            cur = _seg("Currency", ["All"] + found, "dash_cur",
                       help="RWF or USD: every count follows that currency. All: every currency (amounts are always "
                            "shown per currency, never added together)")
            if st.button("🔄 Refresh", key="dash_refresh", width="stretch", help="Read the reports on Google Drive again"):
                state.history_index, state.dash_facts = None, {}
                if state.get("drive") is not None:
                    state.drive.forget_folders()
                st.rerun()
            if st.button("🎯 Change targets", key="dash_targets", width="stretch", help="Settings → Targets"):
                st.switch_page(PAGES["settings"])
            if st.button("↺ Reset every filter", key="dash_reset", width="stretch"):
                for k in [k for k in state if str(k).startswith("dash_") and k not in ("dash_facts",)]:
                    del state[k]
                st.rerun()
            if st.button("→ Go to the reconciliation", key="dash_back", width="stretch"):
                st.switch_page(PAGES["reconciliation"])
        download_pop = b5.popover("⬇", width="stretch", help="Download this view: PDF or Excel")
        share_pop = b6.popover("⇪", width="stretch", help="Share this view on Slack")
    if not options:
        hero_col.info(f"No complete {kinds[kind].lower()} with saved reports yet"
                      + ("" if with_partial else " (switch on Partial, under the period, to see the incomplete ones)")
                      + ".")
        return

    # ---- the facts of the view (bank, status filter) and of the period compared with
    banks = all_codes if pick == "All" else [pick]
    sel_all = dashboard.select(facts, kind, first, last, banks)
    bef_all = dashboard.select(facts, kind, *before_span, banks) if before_span else []
    no_compare = ""
    if bef_all and not dashboard.is_complete(facts, kind, before_span[0], banks):
        no_compare = (f"{dashboard.describe(kind, *before_span)} is not complete (some of its reports are not saved): "
                      "no comparison, rather than a misleading change.")
        bef_all = []
    sel = [f for f in sel_all if f.bank in banks]
    # like for like: the period before counts only the banks that have a report now too (a bank whose report is
    # missing now would otherwise look like a fall); the coverage line names the report that is missing
    bef = [f for f in bef_all if f.bank in banks and f.bank in {x.bank for x in sel}]
    span = dashboard.describe(kind, first, last)
    span_before = dashboard.describe(kind, *before_span) if before_span and bef else ""
    if not sel:
        hero_col.markdown(f'<div class="rh-empty">📊 No report in this view ({_esc(span)}).</div>',
                          unsafe_allow_html=True)
        return
    now, before = dashboard.totals(sel, cur), (dashboard.totals(bef, cur) if bef else None)
    now_by = dashboard.by_bank(sel, cur, banks)
    bef_by = dashboard.by_bank(bef, cur, banks) if bef else {}
    scope_name = "All banks" if pick == "All" else engine.BANKS[pick].FULL_NAME

    # ---- the main view: headline, key cards, charts, merchants, exceptions, needs attention (all from the figures)
    tiles = dashboard.tiles(now, before, now_by, span_before)
    gaps = dashboard.partial(sel, kind, first, last)
    note = ""
    if gaps:                                    # only when "Partial" is on: what this period is missing
        note = " · this period is partial: " + "; ".join(
            f"{engine.BANKS[c].NAME} {n} of {full} {unit}" for c, k, n, full, unit in gaps)
    hero, grid = dashboard_view.view(facts, sel, bef, now, before, now_by, bef_by, kind, first, last, span,
                                     span_before, scope_name, cur,
                                     no_change="not compared" if mode == "none" else "no earlier report")
    with hero_col:
        st.html(hero)
        st.markdown(f'<div class="dt-cover">Coverage: {_esc(dashboard.coverage_line(facts, banks=banks))}{_esc(note)}'
                    + (f' · {_esc(no_compare)}' if no_compare else "") + "</div>", unsafe_allow_html=True)
    st.html(grid)

    # what the downloads (PDF, Excel) cover: the periods up to the one shown
    base = "month" if kind == "custom" else kind
    unit = base
    end_key = pd.Timestamp(last if last is not None else first)
    n = {"day": 14, "week": 12, "month": 6, "quarter": 8, "year": 3}[base]

    # ---- exports: PDF, Excel, Slack summary (all from the figures above)
    items = dashboard.attention(now_by, bef_by, span)
    tr_all = dashboard.trend(facts, kind, end_key, n, cur, banks)
    labels = list(dict.fromkeys(tr_all.sort_values("period")["label"])) if len(tr_all) else [span]
    names_here = [engine.BANKS[c].FULL_NAME for c in banks if now_by.get(c) is not None]
    sc = dashboard.scorecard(now_by, bef_by, found)
    export = {
        "scope": scope_name, "period": span, "compared": span_before, "currencies": found if cur == "All" else [cur],
        "one_month": len(labels) == 1,
        "kpis": [(a, b, c, k in ("bad", "warn")) for a, b, c, k in tiles],
        "banks": [(engine.BANKS[c].FULL_NAME, bank_color(c) if c != "IM" else BRAND) for c in banks
                  if now_by.get(c) is not None],
        "months": labels,
        "processed": {nm: [int(tr_all[(tr_all["bank"] == nm) & (tr_all["label"] == lb)]["processed"].sum())
                           for lb in labels] for nm in names_here},
        "amount": {c: {nm: [float(sum(f.amounts["processed"].get(c, 0.0) for f in facts
                                      if engine.BANKS[f.bank].FULL_NAME == nm and f.kind == ("week" if base == "week" else "month")
                                      and dashboard.period_label(dashboard.period_key(f.start, base), base, True) == lb))
                                for lb in labels] for nm in names_here}
                   for c in (found if cur == "All" else [cur])},
        "status": [("Notified", now["received"], now["received"] / now["processed"] if now["processed"] else 0.0),
                   ("Pending notification", now["pending"], now["pending"] / now["processed"] if now["processed"] else 0.0),
                   ("Paid via another channel", now["duplicate"],
                    now["duplicate"] / now["processed"] if now["processed"] else 0.0)],
        "sla": [(engine.BANKS[c].FULL_NAME, t["sla"], t["sla"] / t["processed"] if t["processed"] else 0.0)
                for c, t in now_by.items() if t is not None],
        "report_columns": ["Bank", "Status", "Payments", "Notified %", "On time %", "Pending", "Other channel"],
        "reports": [(r["Bank"], r["Status"], "–" if pd.isna(r.get("Payments")) else f"{int(r['Payments']):,}",
                     "–" if pd.isna(r.get("Notified %")) else f"{r['Notified %']:.2f}%",
                     "–" if pd.isna(r.get("On time %")) else f"{r['On time %']:.1f}%",
                     "–" if pd.isna(r.get("Pending")) else int(r["Pending"]),
                     "–" if pd.isna(r.get("Other channel")) else int(r["Other channel"])) for _, r in sc.iterrows()],
        "highlights": [(i["kind"] if i["kind"] in ("good", "warn", "bad") else "info", f"{i['title']} — {i['detail']}")
                       for i in items] or [("good", "Every bank with a report is on target.")],
        "merchants": [], "merchants_missing": 0}
    with download_pop:
        st.caption(f"{scope_name} · {span}")
        def dashboard_pdf_bytes():
            """Made when the button is clicked: the blocks on the screen (the same HTML, so the same live numbers)
            printed by Chrome; without Chrome, the earlier PDF layout of the same figures."""
            page = dashboard_view.print_page(grid, scope_name, span, span_before,
                                             now["currencies"] if cur == "All" else [cur],
                                             pd.Timestamp.now().strftime("%d %b %Y · %H:%M"))
            try:
                return dashboard_view.pdf(page)
            except (RuntimeError, OSError, subprocess.SubprocessError):
                return pdf_report.dashboard_pdf(export)
        try:
            st.download_button("⬇ PDF for the manager", data=dashboard_pdf_bytes, type="primary",
                               file_name=f"Reports_dashboard_{'all_banks' if pick == 'All' else pick}_"
                                         f"{re.sub(r'[^0-9A-Za-z]+', '_', span)}.pdf",
                               mime="application/pdf", width="stretch", key="dash_pdf")
        except Exception as err:                     # the page stays usable even if the PDF cannot be made
            st.error(f"The PDF could not be made: {err}", icon="⚠️")
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as xw:
            pd.DataFrame(tiles, columns=["Figure", "Value", "Detail", "Status"]).to_excel(xw, sheet_name="Key figures", index=False)
            sc.drop(columns="code").to_excel(xw, sheet_name="Bank scorecard", index=False)
            tr_all.drop(columns=["code"], errors="ignore").to_excel(xw, sheet_name=f"Trend by {unit}", index=False)
            pd.DataFrame(items).drop(columns=["score"], errors="ignore").to_excel(xw, sheet_name="Needs attention", index=False)
            for c in found:
                dashboard.merchants(sel, bef, c)[0].to_excel(xw, sheet_name=f"Merchants {c}", index=False)
            dashboard.coverage(facts, kind, first, end_key, banks).to_excel(xw, sheet_name="Data behind", index=False)
        st.download_button("⬇ Excel of this view", data=buf.getvalue(), mime=XLSX_MIME, width="stretch",
                           file_name=f"Reports_dashboard_{re.sub(r'[^0-9A-Za-z]+', '_', span)}.xlsx", key="dash_xlsx")
    with share_pop:
        text = "\n".join([f"Hello Team, here is the reconciliation dashboard for {span}"
                          + (f" (vs {span_before})" if span_before else "") + f" · {scope_name}", ""]
                         + [f"• {a.title()}: {b} ({c})" for a, b, c, _ in tiles] + ["", "• Needs attention:"]
                         + ([f"   • {i['title']} — {i['detail']}" for i in items[:5]] or ["   • Nothing: all on target"]))
        slack_send_card(text, f"dashboard {span} {scope_name}", None, key="dash_slack", drive_hint=False)


if PAGE == "dashboard":
    dashboard_page()
    st.stop()                                   # the steps are hidden while the dashboard is open


# =====================================================================================================================
# 8b. REPORT HISTORY (side panel "Report History"): reports on Google Drive first, older ones on this computer
#     Only names are listed; a report is downloaded only when "View" or "Get file" is clicked.
# =====================================================================================================================
if PAGE == "reports":
    t1, t2, t3 = st.columns([1.35, 2.6, .9], vertical_alignment="center")
    if t1.button("→  Go to the reconciliation", key="history_back", width="stretch"):
        state.history_file = None                           # a fetched file leaves memory
        st.switch_page(PAGES["reconciliation"])
    source = t2.segmented_control("Source", ["☁ Google Drive", "💻 This computer (older)"], default="☁ Google Drive",
                                  key="history_source", label_visibility="collapsed") or "☁ Google Drive"
    on_drive = source.startswith("☁")
    client, problem, index = None, "", None
    if on_drive:
        if t3.button("🔄 Refresh", key="history_refresh", width="stretch", help="List Google Drive again"):
            state.history_index = None
        client, problem = drive_client()
        root = config.load().reports_root_folder_id
        if client is None or not root:
            problem = problem or "No reports folder is set (reports_root_folder_id in drive_settings.toml)."
        else:
            cached = state.get("history_index")
            if not cached or cached[0] != root:
                try:
                    with st.spinner("Reading Reconciliation Reports on Google Drive…"):
                        state.history_index = (root, report_history.drive_index(client, root))
                        state.history_index_at = pd.Timestamp.now()
                except DriveError as err:
                    problem, state.history_index = str(err), None
            if state.get("history_index"):
                index = state.history_index[1]
    else:
        local = report_history.local_records(BASE_DIR, None)
        index = report_history.DriveIndex(local, [], sorted({r.folder for r in local}))

    # ---- header: where this is, and what is in it
    if on_drive:
        crumb, title = "My Drive › Bank Reconciliation › <b>Reconciliation Reports</b>", "📁 Reconciliation Reports"
    else:
        crumb, title = "This computer › Bank Reconciliation › <b>output</b>", "💻 Older reports on this computer"
    subtitle = ("Reconciliation reports saved on Google Drive" if on_drive
                else "Reports made before Google Drive was used. Listed only: they are not copied anywhere.")
    st.markdown(f'<div class="rh-head"><div class="rh-crumb">{crumb}</div><div class="rh-title">{title}</div>'
                f'<div class="rh-sub">{subtitle}</div></div>', unsafe_allow_html=True)

    if problem:
        st.markdown(f'<div class="rh-empty">☁️ {_esc(problem)}</div>', unsafe_allow_html=True)
    elif index is not None:
        # ---- three stat cards: how many reports, how many banks, the latest save
        n_reports = len({(r.folder, r.stem) for r in index.records})
        n_banks = len({r.bank for r in index.records})
        times = [t for t in ((local_time(r.modified) if on_drive else r.modified) for r in index.records)
                 if pd.notna(t)]
        latest = max(times) if times else None
        k = st.columns(3)
        for col, (label, value, sub, small) in zip(k, (
                ("Total reports", f"{n_reports:,}", "Available", False),
                ("Banks", f"{n_banks:,}", "With reports", False),
                ("Latest save", f"{latest:%d %b}" if latest is not None else "–",
                 f"{latest:%Y}" if latest is not None else "Nothing saved yet", True))):
            with col.container(border=True):
                st.markdown(f'<div class="rh-k-label">{label}</div>'
                            f'<div class="rh-k-value{" small" if small else ""}">{value}</div>'
                            f'<div class="rh-k-sub">{sub}</div>', unsafe_allow_html=True)

        # ---- Find a report: bank and month come from the report names, wherever the files are
        with st.container(border=True):
            st.markdown('<div class="rh-sec-title">Find a report</div>', unsafe_allow_html=True)
            query = st.text_input("Search", placeholder="Search by bank, month or filename", icon=":material/search:",
                                  label_visibility="collapsed", key=f"history_query_{on_drive}")
            banks_found = sorted({r.bank for r in index.records})
            f1, f2 = st.columns(2)
            h_bank = f1.selectbox("Bank", ["All"] + banks_found, key=f"history_bank_{on_drive}",
                                  label_visibility="collapsed",
                                  format_func=lambda c: "All banks" if c == "All" else
                                  (engine.BANKS[c].FULL_NAME if c in engine.BANKS else c))
            months = sorted({r.month for r in index.records if h_bank == "All" or r.bank == h_bank}, reverse=True)
            h_month = f2.selectbox("Month", ["All"] + months, key=f"history_month_{on_drive}",
                                   label_visibility="collapsed",
                                   format_func=lambda m: "All months" if m == "All"
                                   else f"{pd.Timestamp(m + '-01'):%B %Y}")
        q = query.strip().lower()
        bank_words = {r.bank: f"{engine.bank_name(r.bank)} {engine.BANKS[r.bank].FULL_NAME}" if r.bank in engine.BANKS
                      else r.bank for r in index.records}           # 'I&M' finds the I&M reports too
        records = [r for r in report_history.filter_records(index.records, None if h_bank == "All" else h_bank,
                                                            None if h_month == "All" else h_month)
                   if not q or q in f"{r.title} {r.name} {r.kind} {r.folder} {_period_text(r)} "
                                    f"{bank_words.get(r.bank, '')}".lower()]

        # ---- Saved reports: one row per report
        with st.container(border=True):
            n = len({(r.folder, r.stem) for r in records})
            st.markdown(f'<div class="rh-sec"><span class="rh-sec-title">Saved reports</span>'
                        f'<span class="rh-count">{n} result{"" if n == 1 else "s"}</span></div>',
                        unsafe_allow_html=True)
            st.divider()
            if not index.records:
                st.markdown('<div class="rh-empty">🗂 No report here yet.<br>Reports saved with <b>☁ Save to '
                            'Google Drive</b> appear here, by bank and month.</div>' if on_drive else
                            '<div class="rh-empty">🗂 No older report on this computer.</div>', unsafe_allow_html=True)
            elif not records:
                st.markdown('<div class="rh-empty">🔎 No report matches your search.</div>', unsafe_allow_html=True)
            elif n > 2:                         # more than 2 reports: two are shown, the others by scrolling
                with st.container(height=470, border=False, key="rh_scroll"):
                    history_list(records, client, on_drive)
            else:
                history_list(records, client, on_drive)
    st.stop()                                   # the steps are hidden while Report History is open


def drive_panel():
    """Settings page: Google Drive (who it is used as, sign in / out, test connection, upload test)."""
    drive_cfg = config.load()
    if drive_cfg.problems():
        st.caption("Not set up: " + " ".join(drive_cfg.problems()))
        st.caption("See *Google Drive setup* in README.md.")
    else:
        client, problem = drive_client()
        # ---- who Drive is used as: your Google account (OAuth, default) or the service account (fallback)
        if client is not None:
            if not state.get("drive_email"):
                try:
                    state.drive_email = client.account_email()
                except DriveError as err:
                    st.caption(f"❌ {err}")
            st.caption(("Signed in as " if drive_cfg.is_oauth else "Service account: ")
                       + f"**{state.get('drive_email') or '?'}**")
        elif drive_cfg.is_oauth:
            st.caption(problem)
        else:
            st.caption(f"❌ {problem}")
        st.caption("Reports saved to: " + ("Recon/<bank>/Weekly or Monthly" if drive_cfg.report_layout == "notebook"
                                           else "<bank> Bank/<month>") + f" (report_layout = \"{drive_cfg.report_layout}\")")
        if drive_cfg.is_oauth:
            if client is None and st.button("Sign in with Google", type="primary", width="stretch",
                                            key="drive_sign_in"):
                try:
                    with st.spinner("Waiting for the Google sign-in in your browser (up to 5 minutes)…"):
                        email = google_drive.sign_in(drive_cfg)
                    state.drive_key, state.drive_email, state.drive_checks = None, email, []
                    state.toast = f"Signed in to Google Drive as {email}"
                    st.rerun()
                except DriveError as err:
                    st.error(str(err), icon="⚠️")
            if client is not None and st.button("Sign out", width="stretch", key="drive_sign_out"):
                google_drive.sign_out(drive_cfg)
                state.drive, state.drive_key, state.drive_email, state.drive_checks = None, None, None, []
                state.toast = "Signed out of Google Drive"
                st.rerun()
        if client is not None:
            t1, t2 = st.columns(2)
            if t1.button("Test connection", width="stretch", key="drive_test"):
                checks = []
                for folder_id, label, write in ((drive_cfg.reports_root_folder_id, "Reports folder", True),
                                                (drive_cfg.irembopay_folder_id, "IremboPay folder", False)):
                    if folder_id:
                        checks += client.check_access(folder_id, label, write).lines
                    else:
                        checks.append((False, f"{label}: no folder ID set"))
                state.drive_checks = checks
            if t2.button("Upload test", width="stretch", key="drive_upload_test",
                         disabled=not drive_cfg.reports_root_folder_id,
                         help="Uploads a small test file to the reports folder, checks it, reads it back and "
                              "moves only that file to the bin. Reports are never touched."):
                with st.spinner("Testing an upload to Google Drive…"):
                    state.drive_checks = client.upload_test(drive_cfg.reports_root_folder_id).lines
        for ok, text in state.get("drive_checks", []):
            st.caption(f"{'✅' if ok else '❌'} {text}")
        if state.get("drive_checks") and st.button("✖ Clear", width="stretch", key="drive_clear_checks"):
            state.drive_checks = []                     # hide the results of the last test
            st.rerun()


def slack_panel():
    """Settings page: the Slack webhook (save, send a test, remove)."""
    # the webhook of the channel the report messages go to; kept on this computer only, never shown in full
    hook = slack.load_webhook()
    if hook:
        st.caption(f"✅ Messages go to **{slack.masked(hook)}**")
    else:
        st.caption("Not set up: paste the Incoming Webhook URL of the Slack channel. It is kept on this "
                   "computer only (outside the project folder) and is never shown again.")
    new_hook = st.text_input("Replace the webhook URL" if hook else "Slack webhook URL", type="password",
                             key="slack_settings_url", placeholder="https://hooks.slack.com/services/…")
    if st.button("Save webhook", type="primary", width="stretch", key="slack_settings_save",
                 disabled=not new_hook):
        try:
            slack.save_webhook(new_hook)
            state.slack_checks = []
            state.toast = "Slack webhook saved"
            st.rerun()
        except slack.SlackError as err:
            st.error(str(err), icon="⚠️")
    if hook:
        s1, s2 = st.columns(2)
        if s1.button("Send test", width="stretch", key="slack_settings_test",
                     help="Posts a short test message to the channel"):
            try:
                with st.spinner("Sending a test message…"):
                    slack.send("✅ Test message from the Bank Reconciliation app: the webhook works.", hook)
                state.slack_checks = [(True, "Test message sent: check the channel.")]
            except slack.SlackError as err:
                state.slack_checks = [(False, str(err))]
        if s2.button("Remove", width="stretch", key="slack_settings_remove",
                     help="Forget this webhook (messages cannot be sent until a new one is saved)"):
            state.slack_confirm_remove = True
        if state.get("slack_confirm_remove"):
            st.warning("Remove the Slack webhook from this computer?", icon="⚠️")
            y, n = st.columns(2)
            if y.button("Yes, remove", type="primary", width="stretch", key="slack_remove_yes"):
                slack.remove_webhook()
                state.slack_confirm_remove, state.slack_checks = False, []
                state.toast = "Slack webhook removed"
                st.rerun()
            if n.button("Cancel", width="stretch", key="slack_remove_no"):
                state.slack_confirm_remove = False
                st.rerun()
    for ok, text in state.get("slack_checks", []):
        st.caption(f"{'✅' if ok else '❌'} {text}")


# =====================================================================================================================
# 8d. BANK MANAGEMENT: the banks' business settings (editable by editors and administrators, kept in
#     app_settings.json) and their technical settings (read only, administrators only: they are in banks/<bank>.py)
# 8e. SETTINGS: the targets, who may change settings, Google Drive, Slack
# =====================================================================================================================
def access_line(user: str, role: str):
    """Who is using the app and what they may do here."""
    if not user:
        st.info("Sign in to Google Drive (Settings → Google Drive) to change settings. Until then they can only be "
                "viewed.", icon=":material/lock:")
    else:
        what = {"admin": "can change every setting and see the technical settings",
                "editor": "can change the targets and the SLA days",
                "viewer": "can view the settings (ask an administrator to make you an editor)"}[role]
        st.caption(f"Signed in as **{user}** · {app_settings.ROLES[role]}: {what}.")


def history_table(match=None):
    """The changes recorded in app_settings.json (newest first), optionally only those whose name contains `match`."""
    rows = [h for h in reversed(SETTINGS["history"]) if match is None or match in h["what"]]
    if rows:
        st.dataframe(pd.DataFrame(rows).rename(columns={"at": "When", "by": "Who", "what": "Setting",
                                                        "before": "Before", "after": "After"}),
                     hide_index=True, width="stretch", height=min(38 + 35 * len(rows), 280))
    else:
        st.caption("No change recorded yet.")


def bank_management_page():
    user = current_user()
    role = app_settings.role(user, SETTINGS)
    access_line(user, role)
    tabs = st.tabs([engine.BANKS[c].FULL_NAME for c in engine.BANKS])
    for tab, (code, b) in zip(tabs, engine.BANKS.items()):
        with tab:
            # ---- business settings: changed in the app, without changing the code
            with st.container(border=True):
                st.markdown('<div class="card-title">Business settings</div>', unsafe_allow_html=True)
                day_col = b.COLUMNS["payment_date"]
                if engine.option(b, "SLA_DAYS_SETTING", True):
                    now_days = app_settings.sla_days(SETTINGS, b)
                    new_days = st.number_input(f"SLA: days allowed between the bank {day_col} and IremboPay paid_at",
                                               0, app_settings.SLA_DAYS_MAX, now_days, key=f"bm_sla_{code}",
                                               disabled=not app_settings.can_edit(role),
                                               help=f"0 = paid_at must be on the same day as the bank {day_col}. "
                                                    "Used by every reconciliation of this bank from now on.")
                    if app_settings.can_edit(role) and st.button("Save", type="primary", key=f"bm_save_{code}",
                                                                  disabled=new_days == now_days):
                        try:
                            app_settings.set_sla_days(b, new_days, user)
                            state.toast = f"{b.NAME}: SLA days saved ({new_days})"
                            st.rerun()
                        except app_settings.SettingsError as err:
                            st.error(str(err), icon="⚠️")
                else:
                    st.caption(f"SLA rule of {b.NAME}: the bank {day_col} and IremboPay paid_at must be in the same "
                               "minute. It is part of the bank's rules (no days to set).")
                st.markdown("**Changes**")
                history_table(f"{b.NAME}:")
            # ---- technical settings: in the code, changed by a developer
            with st.container(border=True):
                st.markdown('<div class="card-title">Technical settings</div>', unsafe_allow_html=True)
                if role != "admin":
                    st.caption("Restricted to administrators: statement formats, columns and matching rules are part "
                               "of the code and are changed by a developer.")
                    continue
                st.caption(f"Read only: they are in reconciliation_app/banks/{b.__name__.rsplit('.', 1)[-1]}.py and "
                           "are changed by a developer (a change of rule is planned and approved first).")
                options = ["MONDAY_AFTER_WEEKEND", "DAY_BEFORE_CURRENCIES", "ONE_FILE_PER_CURRENCY",
                           "MISSING_DAYS_BLOCK", "ACCEPT_EARLIER_DAYS", "REVERSE_CHECK", "PAYMENT_FILE_TYPES",
                           "SLA_DAYS_SETTING", "SLA_DIFF_COLUMN", "DUPLICATE_KEY"]
                rows = [("Code", b.CODE), ("Name", b.FULL_NAME), ("IremboPay payment_provider", b.PROVIDER),
                        ("Currencies", ", ".join(b.CURRENCIES)), ("Statement file types", ", ".join(b.FILE_TYPES)),
                        ("Statement sheet in the report", b.STATEMENT_SHEET),
                        ("Paid statuses", ", ".join(sorted(b.COMPLETED_STATUSES))),
                        ("Pending statuses", ", ".join(sorted(b.PENDING_STATUSES)))]
                rows += [(f"Column: {k}", str(v)) for k, v in b.COLUMNS.items()]
                rows += [(f"Option: {k}", str(getattr(b, k))) for k in options if hasattr(b, k)]
                st.dataframe(pd.DataFrame(rows, columns=["Setting", "Value"]), hide_index=True, width="stretch",
                             height=min(38 + 35 * len(rows), 520))
                doc = (b.__doc__ or "").strip()
                if doc:
                    with st.expander("How this bank's statement is read and matched"):
                        st.code(doc, language=None, wrap_lines=True)


def settings_page():
    user = current_user()
    role = app_settings.role(user, SETTINGS)
    access_line(user, role)
    # ---- targets: the Dashboard and the Slack message compare the actual figures with them
    with st.container(border=True):
        st.markdown('<div class="card-title">🎯 Targets</div>', unsafe_allow_html=True)
        st.caption("The Dashboard colours each figure against these targets (green reached, red missed), and the "
                   "Slack message uses them for its 🟢 / 🔴.")
        t = SETTINGS["targets"]
        c1, c2 = st.columns(2)
        receipt = c1.number_input("Notifications received (%)", 50.0, 100.0, round(t["receipt"] * 100, 2), 0.1,
                                  key="set_t_receipt", disabled=not app_settings.can_edit(role),
                                  help="Received by Irembo / payments processed must be at least this")
        ontime = c2.number_input("Payments on time (SLA) (%)", 50.0, 100.0, round(t["ontime"] * 100, 2), 0.5,
                                 key="set_t_ontime", disabled=not app_settings.can_edit(role),
                                 help="(Payments processed - SLA violated) / payments processed must be at least this")
        changed = round(receipt / 100, 6) != round(t["receipt"], 6) or round(ontime / 100, 6) != round(t["ontime"], 6)
        if app_settings.can_edit(role) and st.button("Save targets", type="primary", key="set_t_save",
                                                     disabled=not changed):
            try:
                app_settings.set_targets(receipt / 100, ontime / 100, user)
                state.toast = "Targets saved"
                st.rerun()
            except app_settings.SettingsError as err:
                st.error(str(err), icon="⚠️")
        st.markdown("**Changes**")
        history_table("Target")
    # ---- who may change settings
    with st.container(border=True):
        st.markdown('<div class="card-title">👥 Access</div>', unsafe_allow_html=True)
        admins = sorted(app_settings.admins())
        st.markdown("**Administrators** (set up on this computer in drive_settings.toml, [access] admins): "
                    + (", ".join(admins) if admins else "none yet"))
        editors = SETTINGS["editors"]
        if role == "admin":
            text = st.text_area("Editors (one Google account per line)", "\n".join(editors), key="set_editors",
                                help="They can change the targets and the SLA days")
            new = [e for e in text.replace(",", "\n").splitlines() if e.strip()]
            if st.button("Save editors", type="primary", key="set_editors_save",
                         disabled=sorted(e.strip().lower() for e in new) == editors):
                try:
                    app_settings.set_editors(new, user)
                    state.toast = "Editors saved"
                    st.rerun()
                except app_settings.SettingsError as err:
                    st.error(str(err), icon="⚠️")
        else:
            st.markdown("**Editors**: " + (", ".join(editors) if editors else "none"))
        history_table("Editors")
    # ---- integrations (the same as in the side panel before)
    with st.container(border=True):
        st.markdown('<div class="card-title">☁ Google Drive</div>', unsafe_allow_html=True)
        drive_panel()
    with st.container(border=True):
        st.markdown('<div class="card-title">💬 Slack</div>', unsafe_allow_html=True)
        slack_panel()


if PAGE == "banks":
    bank_management_page()
    st.stop()
if PAGE == "settings":
    settings_page()
    st.stop()


# =====================================================================================================================
# 9. STEP 1 - CHOOSE: the bank, day / week / month, the date, and the optional comparison
#    (+ the No-Transaction Report: run_reconciliation() finds no rows in the period and opens this preview
#    automatically, instead of an error - nothing to choose up front)
# =====================================================================================================================
def render_no_tx_preview(period, bank, made_at, previous=None):
    """The No-Transaction Report as it will be in Excel: the same rows (engine.no_transaction_rows), texts and
    colours as engine.no_transaction_excel()."""
    rows = "".join(
        "<tr>" + "".join(f'<td class="num">{v:,}</td>' if isinstance(v, int) else
                         f'<td class="num">{_esc(v)}</td>' if v in ("N/A", "No change") or str(v).endswith("%")
                         else f"<td>{_esc(v or '')}</td>"
                         for v in r)
        + "</tr>" for r in engine.no_transaction_rows(bank, previous))
    st.markdown(f"""<div class="ntx-paper">
<h3>{_esc(bank.FULL_NAME)} – {_esc(period.name[0].upper() + period.name[1:])}</h3>
<div class="meta"><span>Reporting period: {period.start:%d/%m/%Y} to {period.last_day:%d/%m/%Y}</span>
<span>Generated on: {made_at:%A %d/%m/%Y}</span></div>
<table>
<tr><th class="band" colspan="6">Reconciliation Summary</th></tr>
<tr><th>Category</th><th>Currency</th><th>Number of Transactions</th><th>Total Amount</th>
<th>{_esc(engine.previous_column(period))}</th><th>VS prev {_esc(period.unit)} (%)</th></tr>
{rows}
<tr class="gap"><td colspan="6"></td></tr>
<tr class="status-head"><td>Reconciliation status</td><td colspan="5"></td></tr>
<tr class="status"><td colspan="6" class="title">{_esc(engine.NO_TX_STATUS)}</td></tr>
<tr class="status"><td colspan="6">{_esc(engine.NO_TX_NOTE)}</td></tr>
</table></div>""", unsafe_allow_html=True)


def _week_label(start):
    """A week by its own name (the month of its Thursday), e.g. 'First week of October 2026 (28 Sep - 4 Oct)'."""
    thursday = start + pd.Timedelta(days=3)
    name = ["First", "Second", "Third", "Fourth", "Fifth"][(thursday.day - 1) // 7]
    last = start + pd.Timedelta(days=6)
    return f"{name} week of {thursday:%B %Y} ({start:%d %b} – {last:%d %b})"


def period_options(kind, today):
    """The last 12 full periods of this kind (most recent first), as [(label, start date)]."""
    today = pd.Timestamp(today).normalize()
    if kind == "week":
        this_monday = today - pd.Timedelta(days=today.weekday())
        starts = [this_monday - pd.Timedelta(days=7 * i) for i in range(1, 13)]
        return [(_week_label(s), s) for s in starts]
    first_of_month = today.replace(day=1)
    starts = [pd.Timestamp(first_of_month - pd.DateOffset(months=i)) for i in range(1, 13)]
    return [(f"{s:%B %Y}", s) for s in starts]


CUSTOM_DATE = "Custom date…"


def period_picker(kind, column):
    """A visible way to pick any past day/week/month: a dropdown of the last 12, named the way the team names
    them, or 'Custom date' for anything older (never hidden behind a link)."""
    if kind == "day":
        column.date_input("Date", key="picked", format="DD/MM/YYYY", help="Any day you want")
        return
    options = period_options(kind, pd.Timestamp.today())
    starts = {label: start for label, start in options}
    labels = [label for label, _ in options] + [CUSTOM_DATE]
    current = pd.Timestamp(state.picked)
    current_start = current - pd.Timedelta(days=current.weekday()) if kind == "week" else current.replace(day=1)
    default = next((label for label, start in options if start.normalize() == current_start.normalize()), CUSTOM_DATE)
    choice = column.selectbox(kind.capitalize(), labels, index=labels.index(default))
    if choice == CUSTOM_DATE:
        column.date_input("Custom date", key="picked", format="DD/MM/YYYY",
                          help="Any day inside the week or month you want")
    else:
        state.picked = starts[choice].date()


if state.step == 1 and state.get("no_tx_preview"):
    mem = state.no_tx_mem
    made_at = mem["made_at"]
    with st.container(border=True):
        st.markdown('<div class="card-title">No-Transaction Report</div>', unsafe_allow_html=True)
        st.caption(f"No {bank.FULL_NAME} transactions were found for {period.name}, so this is the report "
                   "made instead. Check it, then download it, or go back to change the bank or the period.")
        if mem.get("previous") is not None:
            st.caption(f"{engine.previous_column(period)}: the transactions of {mem['compared_with']}.")
        else:
            st.warning(f"No {bank.NAME} report found for {period.previous().name}"
                       + (f" ({mem['problem']})" if mem.get("problem") else "")
                       + f": the {engine.previous_column(period)} column shows N/A. Make or upload that report "
                       "(Compare, in step 1), then generate again.", icon="⚠️")
        render_no_tx_preview(period, bank, made_at, mem.get("previous"))
        st.download_button("⬇ Download Excel report", data=mem["xlsx"], file_name=mem["name"], mime=XLSX_MIME,
                           type="primary", key="no_tx_download")

    # ---- ☁ Save to Google Drive and 💬 Send to Slack: the same as Results (Excel only: there is no PDF here)
    with st.container(border=True):
        st.markdown('<div class="card-title">☁ Save to Google Drive</div>', unsafe_allow_html=True)
        save_to_drive_card(mem, skey="no_tx_drive")
    with st.container(border=True):
        st.markdown('<div class="card-title">💬 Send to Slack</div>', unsafe_allow_html=True)
        link = drive_link("no_tx_drive")
        slack_send_card(slack_message.build_no_transactions(bank, period, link, mem.get("previous")), mem["name"],
                        link,
                        key="no_tx_slack")
    c1, _, _ = st.columns([1, 2, 1.3])
    if c1.button("← Back", width="stretch", key="no_tx_back"):
        state.no_tx_preview = False
        st.rerun()

elif state.step == 1:
    with st.container(border=True):
        st.markdown('<div class="card-title">① Choose the bank and period</div>', unsafe_allow_html=True)
        st.caption("How it works: pick the bank and period, the app finds the files itself (set a folder below "
                   "once), Check lists what it found, then Reconcile runs by itself - only the rare row the "
                   "rules truly cannot decide needs a look, on Results, and it never blocks the report.")
        if MANY_BANKS:
            st.segmented_control("Bank", list(engine.BANKS), key="bank_code", selection_mode="single",
                                 format_func=engine.bank_name)
        else:
            st.markdown(f"**Bank:** {bank.FULL_NAME}")
        if state.result is not None and state.result.bank != bank.CODE:     # bank changed: old result no longer valid
            state.result = None
        c1, c2, c3 = st.columns([1.3, 1.6, 1.4], vertical_alignment="bottom")
        c1.segmented_control("Report type", ["Day", "Week", "Month"], key="report_label", selection_mode="single")
        period_picker((state.report_label or "Week").lower(), c2)
        period = engine.make_period((state.report_label or "Week").lower(), state.picked)
        c3.markdown(f"**{period.name[0].upper() + period.name[1:]}**  \n<span class='muted'>"
                    f"{period.start:%d/%m/%Y} to {period.last_day:%d/%m/%Y}</span>", unsafe_allow_html=True)

        # ---- only a period that has ended can be reconciled (a day at midnight, a week on Monday, a month on the 1st)
        not_ended = engine.period_not_ended(period)
        if not_ended:
            st.error(not_ended.replace("\n", "  \n"), icon="⛔")

        # ---- files: a folder found for this bank, or "Use a file instead" (always available, so nobody is blocked)
        # the key carries a "round" number so the checkbox's displayed value can be forced (a checkbox that gets
        # unmounted - e.g. while Check is shown - and remounted with the same key can keep its old frontend value
        # otherwise, the same reason state.uploader_round exists for the file uploaders below).
        has_folders = bank_has_folders()
        if state.get("use_upload_bank") != bank.CODE:        # a new bank: start from its own folder, not the last
            state.use_upload = not has_folders
            state.use_upload_bank = bank.CODE
            state.use_upload_round = state.get("use_upload_round", 0) + 1
        if has_folders:
            checked = st.checkbox("Use a file instead of the folder for this run", value=state.use_upload,
                                  key=f"use_upload_cb_{state.get('use_upload_round', 0)}",
                                  help="Upload the files by hand this once, instead of the folder set for this bank.")
            state.use_upload = checked
        else:
            state.use_upload = True
    folder_settings_editor()
    with st.container(border=True):
        prev = period.previous()
        st.checkbox(f"Compare with previous period ({prev.name})", key="compare",
                    help="Adds a 'VS prev' column to each table. When not ticked, the report is unchanged.")
        ready = not not_ended
        if state.compare and not not_ended:             # a blocked period: nothing to compare yet
            with st.spinner(f"Looking for the {bank.NAME} {prev.name} report…"):   # shown only if it takes a moment
                found, problem = find_previous(prev)
            if found is not None:                       # a short note that fades out by itself (section 3)
                st.markdown(f'<div class="prev-found"><b>{_esc(prev.name)} report found</b> '
                            f'<span>{_esc(found["where"])} · used for the comparison</span></div>',
                            unsafe_allow_html=True)
            else:
                # not a block: if the IremboPay export doesn't go back that far either, the comparison shows
                # N/A on its own (engine.run()); uploading one here is only needed when it would otherwise work
                st.warning(f"No {bank.NAME} report found for {prev.name}" + (f" ({problem})" if problem else ""),
                           icon="⚠️")
                u1, u2 = st.columns([2, 1], vertical_alignment="bottom")
                up = u1.file_uploader(f"Upload the {prev.name} report (.xlsx)", type=["xlsx"],
                                      key=f"prev_report_{state.uploader_round}")
                if up is not None:
                    try:                        # checked by the engine, then kept in memory only
                        session = state.session_folder
                        path = engine.save_uploaded_report(BASE_DIR, prev, up.name, up.getvalue(), bank,
                                                           root=session.path)
                        state.prev_reports[(bank.CODE, prev.store_key)] = {
                            "where": "(uploaded)", "name": path.name, "bytes": path.read_bytes()}
                        session.remove(path)
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
    if c3.button("Next: check →", type="primary", width="stretch", disabled=not ready,
                help=None if ready else "This period has not ended yet" if not_ended
                else "Add the previous report first, or untick the comparison"):
        state.run_error = None
        go(2)

# =====================================================================================================================
# 10. STEP 2 - CHECK: files (a folder, or "Use a file instead"), the checklist, then "Run reconciliation"
# =====================================================================================================================
elif state.step == 2:
    if state.get("use_upload") or not bank_has_folders():
        acc = render_upload_section()
        found = None
    else:
        found = found_files_for_period()
        acc = folder_files.all_files(found) if found is not None else []
    chk = engine.check_period(period, acc, bank)
    with st.container(border=True):
        st.markdown(f'<div class="card-title">② Check: {period.name}</div>', unsafe_allow_html=True)
        if found is not None:
            for line in folder_files.describe(found):
                st.markdown(f"📁 {line}")
            for problem in found.problems:
                st.warning(problem, icon="⚠️")
        for ok, text in chk.lines:
            st.markdown(f"{'✅' if ok else '❌'} {text}")
        for w in chk.warnings:
            st.warning(w, icon="⚠️")
        if chk.blocking:
            st.error("Run is blocked until the missing bank files are added."
                     + (" Check the folder, or use a file instead (Choose)." if found is not None
                        else " Add the missing file above."), icon="⛔")
        if compare:
            st.caption(f"The report will be compared with {period.previous().name}.")
        if state.run_error:
            st.error(state.run_error, icon="⚠️")
    c1, _, c3 = st.columns([1, 1.5, 1.8])
    if c1.button("← Back", width="stretch"):
        state.use_upload = found is None     # back to Choose with the same choice this screen was shown with
        state.use_upload_round = state.get("use_upload_round", 0) + 1
        go(1)
    if c3.button("▶  Run reconciliation", type="primary", width="stretch", disabled=bool(chk.blocking)):
        try:
            with st.spinner(f"Reconciling {period.name}…"):
                state.result = run_reconciliation(acc)
            state.run_error = None
            if state.result.stopped_for_review:    # the rules could not decide a row automatically: this should
                state.run_error = ("A row could not be classified automatically. This should not happen - "
                                   "please tell a developer.")
            else:
                go(3)
        except engine.NoTransactionsError:
            # the files cover the period, but none of their rows fall in it: every category would be 0 anyway,
            # so show the No-Transaction Report directly instead of an error
            state.no_tx_mem = build_no_tx_mem()
            state.no_tx_drive = None
            state.run_error = None
            state.step = 1
            state.no_tx_preview = True
            st.rerun()
        except engine.ReconError as err:
            state.run_error = str(err)
            st.rerun()
        except Exception as err:
            state.run_error = f"Something unexpected went wrong: {err}"
            st.rerun()

# =====================================================================================================================
# 11. STEP 3 - RESULTS
#     ① summary of the run   ② Quick Export: the standard report as it is (Excel / PDF), no configuration
#     ③ Custom Report (closed until "Customize Report"): type, sections, views, layout, preview, export
# =====================================================================================================================
elif state.step == 3:
    res = state.result
    mem = state.get("report_mem")
    if (res is None or res.stopped_for_review or res.report_file is None
            or not mem or mem["name"] != res.report_file.name):           # the report is kept in memory only
        go(2)
    main = next((b for b in res.blocks if b[0] in ("Summary-Monthly", "Summary-Daily")), res.blocks[0])

    # ---- ① the summary of the run
    with st.container(border=True):
        st.markdown('<div class="card-title">📊 Report &amp; Export</div>', unsafe_allow_html=True)
        st.markdown(f"**{bank.FULL_NAME} · {period.name[0].upper() + period.name[1:]}**")
        st.caption(f"Reconciled {res.ran_at:%d %b %Y · %H:%M}"
                   + (f" · compared with {res.compared_with}" if res.compared_with else ""))
        # rows not found in IremboPay are classified automatically (no manual review for normal cases): what their
        # decision became is simply in the table it chose (Pending / Duplicate), not counted or listed again here
        k = st.columns(4, border=True)
        k[0].metric("Processed", f"{res.bank_rows:,}")
        k[1].metric("Matched", f"{res.matched:,}")
        k[2].metric("Rate", f"{res.matched / res.bank_rows:.2%}" if res.bank_rows else "–")
        k[3].metric("Not found", f"{res.not_found:,}")
        k[3].caption("classified automatically" if res.not_found else "none")
        st.markdown("  \n".join(f"**{cur}** · {d['rows']:,} rows · {d['matched']:,} matched"
                                for cur, d in res.by_currency.items()))
        for level, text in res.messages:
            if level == "warning":                  # only problems are shown (no "all good" messages)
                st.warning(text, icon="⚠️")
        if state.get("needs_a_look"):
            n = len(state.needs_a_look)
            with st.expander(f"👀 Needs a look later ({n}): could not check these invoices on IremboPay's own "
                             "payment page"):
                st.caption("They are already counted as Pending, and the report was made normally: nothing "
                          "waited for this. Check them again later, once IremboPay's page or the internet "
                          "connection is back.")
                st.write(", ".join(state.needs_a_look))
        with st.expander("Summary tables"):
            render_summary(res.blocks, period.report_type, res.merchants)
        with st.expander("💬 Team message (copy for Slack)"):
            st.caption("Same figures as the report above, ready to paste into Slack with nothing to edit.")
            pending = res.sheets.get("Pending notification") if getattr(res, "sheets", None) else None
            pending_statuses = pending["payment_status"].value_counts().to_dict() if pending is not None and len(pending) else None
            team_message = slack_message.build(bank, period, res.blocks, res.merchants, pending_statuses,
                                               res.previous_counts, drive_link())
            st.code(team_message, language=None, wrap_lines=True)

    # ---- ② Quick Export: the standard report, unchanged
    with st.container(border=True):
        st.markdown('<div class="card-title">⚡ Quick Export</div>', unsafe_allow_html=True)
        st.caption("Download the standard reconciliation report. No configuration required.")
        q1, q2, _ = st.columns([1, 1, 1])
        q1.download_button("⬇ Download Excel", data=mem["xlsx"], file_name=mem["name"],
                           mime=XLSX_MIME, type="primary", width="stretch", key="quick_xlsx")
        q2.download_button("⬇ Download PDF", data=mem["pdf"], file_name=Path(mem["name"]).with_suffix(".pdf").name,
                           mime="application/pdf", type="primary", width="stretch", key="quick_pdf")

    # ---- saved to Google Drive and sent to Slack automatically (once per report)
    with st.container(border=True):
        st.markdown('<div class="card-title">📤 Saved &amp; shared</div>', unsafe_allow_html=True)
        auto_delivery(res, mem)

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
            if state.get("cr_type") not in engine.REPORT_TYPES:   # first time, or a type that no longer exists
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
                        if key == "merchants":                         # a merchant chart: the top 10 merchants
                            v3.caption("Top 10 merchants")
                        else:
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
                        data, ext, mime = pdf_report.custom_pdf(content), "pdf", "application/pdf"
                    else:
                        data, ext, mime = engine.custom_excel(content), "xlsx", XLSX_MIME
                    e2.download_button("⬇ Generate & Download Report", data=data,
                                       file_name=f"{content['file_stem']}.{ext}", mime=mime, type="primary",
                                       width="stretch", key="cr_dl")
                except Exception as err:
                    e2.error(f"The custom report could not be made: {err}", icon="⚠️")

    if state.run_error:
        st.error(state.run_error, icon="⚠️")
    c1, c2, c3 = st.columns([1, 1.3, 1.5])
    if c1.button("← Back", width="stretch"):
        go(5 if res.not_reviewed else 3)
    if c2.button("↻  Run again", width="stretch",
                help="Reconcile this bank and period again (e.g. after adding files to the folder)"):
        try:
            with st.spinner(f"Reconciling {period.name}…"):
                state.result = run_reconciliation(files_for_run())
            state.run_error = None
            go(5 if state.result.stopped_for_review else 6)
        except engine.ReconError as err:
            state.run_error = str(err)
            st.rerun()
        except Exception as err:
            state.run_error = f"Something unexpected went wrong: {err}"
            st.rerun()
    if c3.button("🔄  New report", width="stretch"):
        state.files, state.read_errors, state.result = {}, [], None
        state.customize = state.cr_preview = False
        state.report_mem = state.drive_save = None          # the report leaves memory (it is on Drive if saved)
        state.session_folder.clear()                        # only this session's own temporary folder
        go(1)
