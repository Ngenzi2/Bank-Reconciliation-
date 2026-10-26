# BPR Bank Rwanda — bank configuration (draft)

Status: **built on 16/10/2026** (`reconciliation_app/banks/bpr.py`, tests in `tests/test_bpr.py`). Everything below was read from the August 2026 files: 23 BPR statements in
`August/` (including `JULY 30 TO AUG 2 frw.xlsx`, added 16/10/2026) and the IremboPay payment file
`August/Payment details.csv` (618 payments). The September statements
(`Sept statements/`) and the September payment file were used to check the same points. The decisions are in section 9.

## 1. Identity

| Setting | Value | Evidence |
|---|---|---|
| CODE | `BPR` | short code for file names (`BPR_recon_summary_…`) |
| NAME / FULL_NAME | `BPR` / `BPR Bank Rwanda` | the payment files are in "BPR Payment Details" |
| PROVIDER (IremboPay `payment_provider`) | `BPR_RW` | 523 of 618 August payments; all three BPR channels |
| Other provider value | `BPR` (95 payments, all `BANK_TRANSFER`) | see decision 9.2 |
| CURRENCIES | RWF, USD | 540 RWF credits, 5 USD credits in August |
| FILE_TYPES | `xlsx` | every statement is .xlsx |
| Paid statuses | SETTLED, PAID, PAYOUT_INITIATED (same as I&M) | August: SETTLED 572, PAID 46 |
| Pending statuses | NEW (20/10/2026: EXPIRED is not a payment status; an expired invoice is NEW with a note) | none in August |

IremboPay channels of `BPR_RW`: `BPR_RW_AGENT` (398 · agents), `BPR_RW_IB` (112 · internet banking),
`BPR_RW_BRANCH` (13 · branch).

## 2. The statement

An export of the bank's table `BNR.ACCOUNT_STATEMENT_BALANCE` (the query is in the files' "SQL" sheet), one file per
day or span of days, one currency per file.

| Column | Meaning |
|---|---|
| BOOKING_DATE | business day of the credit (`YYYY-MM-DD`) |
| BANK_REFERENCE | bank reference, `FT…` (transfers) or `TT…` |
| DESCRIPTION | **the IremboPay `provider_reference`** (see 4), or a branch/other text |
| DOC ID | the IremboPay invoice number — **added by hand** after the export (not in the SQL query; "Sheet1" is the lookup list used to fill it) |
| VALUE_DATE | date **and time** of the payment at the bank |
| DEBIT | money out, as a **negative** number |
| CREDIT | money in |
| CLOSING_BALANCE | balance after the row; the first row is "Opening balance" |

Differences between the files (the reader must accept all of them):
- the DOC ID title is `DOC ID`, `DOCID`, `Doc ID`, `doc id`, `doc ID` or **blank** (Aug 4, Aug 12);
- some USD files have **no DOC ID** column; the 22 Sept – 13 Oct RWF file has **no VALUE_DATE**;
- the data sheet is usually "Export Worksheet" (once "Sheet1"); extra sheets "Sheet1/3/5" (lookup lists) and "SQL";
- one file has no "Opening balance" row (`august 7 to 09 rwf.xlsx`);
- **no currency column**: the currency is only in the file name (`… USD …`) — see decision 9.4.

## 3. What is kept

- **Credits** only (CREDIT > 0) are reconciled, like I&M and Bank of Africa.
- Debits are counted, not reconciled. They are one per day, e.g. RWF 1,973,518 on 25/09: the money collected moved
  out of the account. A credit cancelled by a debit of the same description and amount on the same day is not kept
  (USD 6,000 on 05/10).
- The "Opening balance" row is not a payment.

## 4. Matching with IremboPay

**The DESCRIPTION is the IremboPay `provider_reference`.** In August:

| Way | Statement credits matched | Note |
|---|---|---|
| DESCRIPTION = `provider_reference` | 497 of 545 | the same invoice as DOC ID for **497 of 497** |
| DOC ID = `invoice_number` | 518 of 545 | DOC ID is filled by hand |
| DESCRIPTION = `invoice_number` | 2 | |
| any of the three | **519 of 545** | amounts equal for **525 of 525** matched rows |

The kind of DESCRIPTION tells the channel:

| DESCRIPTION looks like | IremboPay channel | Matched by |
|---|---|---|
| `SZABAB…` (12 characters) | `BPR_RW_AGENT` (414) | provider_reference |
| a number, e.g. `20145323788` | `BPR_RW_IB` (99) | provider_reference |
| `FTCM…`, other text | `BPR_RW_BRANCH` (8) | DOC ID only |

**Decided (16/10/2026): the invoice is the DOC ID.** When there is no DOC ID, the invoice is looked for in IremboPay by
BANK_REFERENCE or DESCRIPTION (= `provider_reference`) or DESCRIPTION (= invoice number), and used only when IremboPay
has it (August: 7 of 21 found, e.g. TT26226M0NPS = USD 750 for RFDA; September: 7 of 45). The DESCRIPTION =
`provider_reference` link also tells which credit IremboPay recorded when an invoice is credited more than once.

Not matched in August (26 credits): `FTCM…` branch credits without DOC ID (e.g. RWF 1,000,000 on 14/08,
RWF 285,000 and RWF 260,000 on 06/08), 2 `TT…` credits with no description, and 6 credits whose DOC ID is not in the
August payment file. One of these is in the September file (`886236656500`, credited 26/08, `paid_at` 23/09, which is
late). These go to Exception Resolution like any bank.

## 5. Dates: period and SLA

| Comparison (525 matched August payments) | Result |
|---|---|
| `paid_at` − VALUE_DATE | median **0.7 minute**; 95 % within **1.3 minutes**; largest 5 days |
| `paid_at` day − BOOKING_DATE | same day 514 · the day before 3 · 1 to 5 days after 8 |

- **The date that counts is VALUE_DATE** (decided 16/10/2026): it sets the **period** of a credit and is the date
  the **SLA** compares with `paid_at`.
- VALUE_DATE and BOOKING_DATE are the same day for almost every credit (August 555 of 558, September 634 of 643).
  The others were paid **late in the evening and booked the next morning** (e.g. paid 30/08 21:07, booked 31/08;
  paid 10/09 23:45, booked 11/09). So:
  - a credit belongs to the day of its VALUE_DATE, even when it is in the next day's file;
  - **the statement of the day after the period is needed** for its last evening's payments (e.g. 31/08 evening
    payments are in the 1/09 statement). Step 3 asks for it, the way it asks I&M for the Monday after a weekend.
- **A statement without VALUE_DATE cannot be used as it is:** `sept 22 to oct 13.xlsx` (465 credits) has only
  BOOKING_DATE. See decision 9.7.
- **Weekends:** BPR books on Saturdays (rows on 22/08); 23/08 (Sunday) had no transaction and the balance continues.
  There is no "Monday after the weekend" rule (`MONDAY_AFTER_WEEKEND = False`).

## 6. Checks before Run

Every file should be checked against its own balances, like Bank of Africa: the opening balance + each row = the next
balance, and each file opens where the one before closed. In August this finds:
- two files whose **balance column is shifted** while their rows are complete (the day before + their rows = the day
  after): **3 Aug** (+RWF 20,000: 2 Aug closes 12,000, 3 Aug opens 32,000; 3 Aug closes 162,000, 4 Aug opens
  142,000) and **10 Aug** (−RWF 33,800). So the check must compare the rows, not trust a file's own opening balance;
- the RWF balance can go **below zero** within a day: the daily debit is booked before the day's credits (2 Aug:
  −RWF 198,000 after the 03:25 debit, back to +12,000 by the evening).

1–2 Aug are covered by `JULY 30 TO AUG 2 frw.xlsx` (30 Jul – 2 Aug, balances consistent): 4 RWF credits on 1–2 Aug,
RWF 240,000, **all 4 matched** (DESCRIPTION = `provider_reference`, and DOC ID). The fifth IremboPay payment of those
days is **USD 150** (`886212617782`, internet banking, 1 Aug): it needs a **USD statement for 1 Aug**, which is
missing (the USD files cover 6, 14 and 20 Aug only).

Files with a wrong name: `august 11 RWF.xlsx` contains **10 Aug**; `aug 21 to 23 RWF.xlsx` contains 21–22 Aug
(23 Aug had nothing). The reader uses the dates inside the file, not its name.

## 7. Summary tables (same as I&M)

| Table | Rule |
|---|---|
| Processed | the invoices on both sides, each once = received + pending + paid with another channel (20/10/2026) |
| Received by Irembo | found, `payment_provider` = BPR_RW, paid status |
| SLA violated | from the SLA rule (decision 9.3) |
| Pending | NEW in IremboPay (an expired invoice is NEW) |
| Paid with another channel | found, but IremboPay says another provider |

**Credited twice:** 11 invoice numbers were credited twice or three times in August, each time with a different bank
reference and its own DESCRIPTION. For example, `886232212512` was credited three times RWF 10,000 on 20/08. These
are flagged as extra credits, as I&M does.

## 8. August in figures

| | RWF | USD |
|---|---|---|
| Credits on the statements (1–31 Aug) | 544 · RWF 15,500,675 | 5 · USD 3,380 |
| Debits (1–31 Aug) | 17 · RWF 5,524,250 out | 0 |
| RWF days covered | every day 1–31 Aug (23 Aug: no transaction) | |
| IremboPay payments (`BPR_RW`) | 523 | |
| `BPR_RW` payments on no statement | 12 (11 on other days + USD 150 on 1 Aug, no USD statement) | |

## 9. Decisions (16/10/2026) — what was built

The flow is the same as the other banks: every credit is checked transaction by transaction in the IremboPay file;
what cannot be settled automatically goes to Exception Resolution.

| # | Decision | In the app |
|---|---|---|
| 1 | Name "BPR Bank Rwanda", short name "BPR", code `BPR` | step 1 bank button; reports `BPR_recon_summary_…` |
| 2 | **DOC ID is the invoice**; a credit **without DOC ID** is found in IremboPay by (1) BANK_REFERENCE = `provider_reference`, (2) DESCRIPTION = `provider_reference`, (3) DESCRIPTION = invoice number — **only when IremboPay has it** | found: matched like a DOC ID ("Invoice found by" in Payment details); **no invoice number: straight to Pending notification** (not in Exception Resolution; currency from the file name), unless cancelled by a debit (then reviewed); no extra table or sheet |
| 3 | IremboPay `payment_provider` **BPR_RW** is BPR | a credit IremboPay shows under another provider (e.g. `BPR`) goes to review |
| 4 | **VALUE_DATE** is the date (period and SLA) | the statement of the **day after** the period is recommended (evening payments) |
| 5 | **SLA in whole minutes, 1 minute allowed** (changeable in Bank Management) | paid_at's minute = VALUE_DATE's minute or the next one: on time; more: SLA violated |
| 6 | **Currency from IremboPay** | a credit **not in the IremboPay file** (DOC ID unknown, or no DOC ID): currency from the **file name** (`USD` in it = USD, else RWF), changeable in step 5 (**Choose currency**); step 5's IremboPay check looks an unknown DOC ID up on checkout (**Paid → Paid with other channel**, confirmed by the reviewer) |
| 7 | Rows without VALUE_DATE go to review | BOOKING_DATE only decides which period's review they appear in; a credit also in a file with VALUE_DATE (same BANK_REFERENCE and amount) is kept once, with its VALUE_DATE (19/10/2026) |
| 8 | An invoice credited more than once | the credit IremboPay recorded (DESCRIPTION = provider_reference) counts; the others go to review, each saying whether IremboPay **also recorded** it (a second IremboPay transaction) or **has no record** of it; in step 5 they are in their own table "Invoice credited more than once", next to the credit counted (19/10/2026) |
| 9 | A credit cancelled by a debit | goes to review |
| 10 | Nothing may be missing | every file checked row by row (refused if it does not add up) and against the files around it: money in with no row = a payment missing (shown, **Run is not blocked**); days with no statement: shown, Run allowed, named in the Slack message; money out with no row = a debit missing (shown); a shifted balance column with complete rows = accepted with a note; a file inside another = counted once; a file that starts by repeating the last row(s) of the file before (same BANK_REFERENCE) = checked after those rows (19/10/2026) |
| 11 | IremboPay `BPR_RW` payments on no statement | listed in step 3 |
| 12 | Pending = NEW (an expired invoice is NEW, with a note) | same as I&M |

Checked on the real files: August — 550 credits, every day covered (23/08: no transaction), 40 to review (21 no DOC
ID, 13 credited again, 6 DOC ID not in IremboPay); September — 855 credits, 271 to review (210 of them from
`sept 22 to oct 13.xlsx`, which has no VALUE_DATE).
