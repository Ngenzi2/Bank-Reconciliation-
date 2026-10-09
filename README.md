# Bank Reconciliation – User guide

Compares the **bank statements** (I&M, Ecobank; RWF and USD) with the **IremboPay payments** and makes the
reconciliation report (Excel or PDF) for one **day**, **week** or **month**. Choose the bank in step ①.

You only upload the files. There are **no folders to create and no files to rename**; the app reads the files
and does not keep them. It keeps only the reports.

---

## First time only: setup
Double-click **`Setup (run once).command`** in the Bank Reconciliation folder and wait for *Setup finished*.

> If macOS says the file "cannot be opened because it is from an unidentified developer":
> right-click the file → **Open** → **Open** (only once).

---

## Make a report: 6 steps
Double-click **`Start Reconciliation.command`**. The app opens in your browser (keep the small black window open).

**① Period:** click *Day*, *Week* or *Month* and pick the date.
- Day = that day · Week = Monday to Sunday · Month = the whole month
- **Only a period that has ended can be reconciled** (it ends at midnight after its last day). Today, yesterday is
  the latest day, last week the latest week, last month the latest month. A period that has not ended is blocked
  with the day it opens, e.g. *⛔ The week 05/10/2026 – 11/10/2026 has not ended yet. It can be reconciled from
  Monday 12/10/2026.*
- Optional: tick **Compare with previous period** to add a *VS prev* column. The app finds the previous report
  by itself; if it does not exist, upload it or click *Make the … report first*.

**② Upload:** drop the files from your Downloads folder (names as downloaded):
- the **I&M bank statements**, RWF and USD, for every day of the period
- if the period ends on a **Saturday or Sunday**, also **Monday's** statement (late weekend payments are posted on Monday).
  Run waits until it is uploaded: Monday's statement exists from Tuesday
- recommended: the RWF statement of **the day before** the period (it can hold payments of the first day)
- the **IremboPay payment file** (.csv)

Wrong files, the same file twice, or files outside the period are refused with a message. Remove a file with 🗑.

**③ Check:** a list shows what is covered (✅) and what is missing (❌), e.g.
*❌ Bank RWF: Tuesday 08/09 missing → upload the RWF statement for that day*.
**Run** works only when every bank day is covered. Then click **▶ Run reconciliation**.

**④ Reconcile:** the result of the matching, *Reconciled on <day> at <time>*: the bank rows, how many were
**matched** in IremboPay (with %), how many were **not found**, and the amounts, then one line per currency.
If every row matched, **Next: download** skips step ⑤ and goes to ⑥; otherwise click **Review & Resolve … Exceptions →**.

**⑤ Exception Resolution** (only when invoices are missing in IremboPay): for each one choose the **Report table**,
set **Reviewed = Yes**, then click **💾 Save decisions**. The **Result** line under the table counts the choices
(e.g. *9 Paid with other channel · 6 Pending notification*).

| Report table | Use it when |
|---|---|
| Notification received | paid on I&M and recorded in IremboPay |
| Pending notification | the invoice is still NEW in IremboPay |
| Paid with other channel | IremboPay says it was paid another way (type the provider if known) |
| Failed notification | the invoice is not in IremboPay at all |
| Exclude | not an Irembo payment |

Invoices already decided in an earlier report are filled in as a suggestion: check them and set *Reviewed = Yes*.

**⑥ Report & Export:** the summary of the run (processed, matched, exceptions, rate, one line per currency; the
**Summary tables** fold out). Then:
- **⚡ Quick Export:** the standard report as it is, no configuration: **⬇ Download Excel** (the full report) or
  **⬇ Download PDF** (the summary only).
- **⚙ Custom Report** (optional, click **Customize Report →**): choose a **report type** (*Full Report*,
  *Management Summary* or *Exceptions Report*: it ticks the sections for you), the **sections**, how each one is
  shown (**Table**, **Bar chart** or **Line chart**, grouped by week, day or currency), and the **layout**
  (landscape or portrait, standard or compact). **👁 Preview Report** shows the result; pick **Excel** or **PDF** and
  click **⬇ Generate & Download Report**. Charts are in the PDF only (the Excel shows the same numbers as tables).
  A custom report is only downloaded: the standard report saved in `output/` never changes.

---

## Good to know
- **Two dates per report:** the *period reconciled* (e.g. Week 28/09/2026 – 04/10/2026) and the day it was
  *reconciled on* (the day of the run). Both are in the report header (Excel, PDF, saved-report view) and the run log.
- **Transactions are counted on the day they happened** (I&M: Value Date; Ecobank: Request Date in Kigali time),
  not on the date of the file, which is the next day because files are collected at midnight.
- **The SLA** compares the bank **Value Date** with the IremboPay **paid_at date** (by day), for I&M payments only.
- **Saved on this computer:** the reports (`output/day`, `output/week`, `output/month`, also listed in the side panel),
  the exception decisions (`output/.review_store.csv`, no bank data, also listed in the side panel) and the run log (`output/run_log.csv`).
- **Private:** the app runs only on this computer; nothing is sent to the internet.
- The month folders (`2026-08/…`) are used only by the developer's test notebook, never by the app.

---

## For the developer: where the code is
Each file starts with a short explanation and a numbered **CONTENTS**; the sections in the file follow those numbers.

| File | What it holds |
|---|---|
| `reconciliation_app/app.py` | the screens: side panel + the 6 steps (no reconciliation rules) |
| `reconciliation_app/engine.py` | the rules that are the same for every bank: periods, reading files, checks, review, the run, the Excel report, the PDF |
| `reconciliation_app/banks/__init__.py` | the list of banks, what a bank file must provide, the optional settings |
| `reconciliation_app/banks/im.py` | I&M: statement format, SLA, calculations (the template for a new bank) |
| `reconciliation_app/banks/ecobank.py` | Ecobank: the same sections + its own rules (one row per invoice, period by paid_at, extra sheets) |
| `.streamlit/config.toml` | app settings: colours, font, local-only access |

Start with the top of `engine.py`: it lists what happens in one reconciliation and where each part is.

## For the developer: add a new bank
The rules that are the same for every bank (periods, IremboPay file, checks, review, report layout, comparison,
run log) are in `reconciliation_app/engine.py`. Everything specific to one bank is in one file in
`reconciliation_app/banks/`: I&M is `banks/im.py`.

1. Copy `banks/im.py` to a new file, e.g. `banks/bk.py`.
2. Change its values: `CODE` (used in the report file name, e.g. `BK_recon_summary_…`), `NAME`, `FULL_NAME`,
   `PROVIDER` (its name in the IremboPay `payment_provider` column), `CURRENCIES`, `FILE_TYPES`, `STATEMENT_SHEET`,
   and `COLUMNS`: which statement column is the **posted date**, the **payment date** (used for the period and the
   SLA), the description, the bank reference and the amount.
3. Change its functions: `read_statement` (statement format), `clean_statement` (invoice number), `channel`,
   `sla`, and the calculations `received`, `pending`, `duplicate`.
4. Set `MONDAY_AFTER_WEEKEND` and `DAY_BEFORE_CURRENCIES` to match how this bank posts late payments.
5. Add it to `ACTIVE` in `banks/__init__.py`.

When the app starts it checks every bank file and names anything missing. With more than one bank, step 1 shows a
**Bank** choice; reports, exception decisions and runs are kept per bank.
