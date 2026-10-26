# BPR Bank Rwanda: how the reconciliation works

This is the full path of a BPR reconciliation in the app, from the files you upload to the Slack message. It covers
what is read from each file, every check, and every rule. The code is in `reconciliation_app/banks/bpr.py` (the
BPR rules) and `reconciliation_app/engine.py` (the steps shared by all banks).

**In one sentence:** every **credit** on the BPR statements is matched, one by one, with its invoice in the
IremboPay payment file. A credit is either settled automatically or sent to Exception Resolution for someone to
decide. Then the report is written, saved to Drive and announced on Slack.

---

## 1. The two kinds of files

### 1.1 The BPR statement (.xlsx, one day or a few days per file)

This is an export of BPR's table `BNR.ACCOUNT_STATEMENT_BALANCE`. The app looks for the sheet that has a
`BOOKING_DATE` header, wherever that row is. It ignores the other sheets (the lookup lists "Sheet1/3/5" and "SQL").

| Column read | What the app uses it for |
|---|---|
| `BOOKING_DATE` | The bank's business day. Used only when `VALUE_DATE` is missing, to decide the period. |
| `BANK_REFERENCE` | The bank's reference (`FT…`, `TT…`). Used to find an invoice, to find the same row in two files, and to name the credit in messages. |
| `DESCRIPTION` | Usually the IremboPay **provider_reference** (agent `SZ…`, internet-banking numbers). Used to find an invoice, to tell which credit IremboPay recorded, and to tell the channel. |
| `DOC ID` | **The IremboPay invoice number**, filled in by the team. Also accepted as `DOCID`, `Doc ID`, `doc id`, or a blank title. |
| `VALUE_DATE` | **The date and time of the payment.** It decides the period and the SLA. |
| `DEBIT` | Money out (negative). Not reconciled. Used for the balance checks and to spot a cancelled credit. |
| `CREDIT` | Money in. **Only the credits are reconciled.** |
| `CLOSING_BALANCE` | The balance after each row. Used to prove nothing is missing. |

- **No currency column:** the currency comes from IremboPay (see 4.3).
- **"Opening balance" row:** it is not a payment.

### 1.2 The IremboPay payment file (.csv)

This is the export of IremboPay's BPR payments. The app requires these columns: `invoice_number`, `amount`,
`currency`, `payment_status`, `paid_at`, `payment_provider`, `transaction_status`, `created_at` and `name` (the
merchant). It also reads `provider_reference`, `payment_channel`, `transaction_reference` and `settled_at`.

- **A file with no `BPR_RW` payments** is refused.
- **An invoice with several transactions** is kept once: the **successful** transaction, otherwise the latest one.
- **Every `provider_reference` of the invoice's transactions** is still kept. This shows whether IremboPay
  recorded a second credit for the same invoice.

---

## 2. Steps 1 and 2: choosing the period and uploading

1. **Step 1: choose BPR and the period** (a day, a week Monday to Sunday, or a month). Only a period that has ended
   can be reconciled.
2. **Step 2: upload the files.** As each statement is read, it is checked on its own:
   - **The balances must add up row by row:** the balance before + the credit + the debit = the balance after. If
     one row does not add up, the file is **refused** with the place where it breaks ("ask BPR for the complete
     export").
   - **The days a file covers** are read from the dates inside it, not from its name (`august 11 RWF.xlsx` holds 10 Aug).
   - **Accepted range:** a statement is accepted when it covers any day from the **day before** the period to the
     **day after** it.
   - **The same file twice** is refused.
   - **A credit cancelled by a debit** (same DESCRIPTION, same amount) is marked "Cancelled by <debit reference>".
   - **A file without `VALUE_DATE`:** its credits are marked "VALUE_DATE missing", and `BOOKING_DATE` stands in for
     the period.

---

## 3. Step 3: the check before Run (it never blocks BPR)

Every line is **shown**. For BPR nothing stops the Run button. The only exceptions, true of every bank, are a
period that has not ended yet and no file at all.

| Check | How | What you see |
|---|---|---|
| **Files follow each other** | RWF files and USD files (by "USD" in the name) are put in date order. Each file must open where the one before closed. | ✅ "every file follows the one before" |
| A file that **repeats the last rows** of the file before | The rows with the same `BANK_REFERENCE` are skipped, and the check starts after them. *(Added 19/10/2026, e.g. OCT 15 RWF.xlsx repeating the 14 Oct debit.)* | ✅ "starts with 1 row(s) already in …; counted once" |
| A file whose **balance column is shifted**, though its rows are complete | The file before + its rows = the file after | ✅ "has its balance column shifted by …" |
| A file **inside another** (same rows) | Counted once | ✅ "is already inside …" |
| **More money** at the start of a file than the file before ended with | A credit is missing between the two files | ❌ "… came in with no row in the files: a payment is missing" (not blocking) |
| **Less money** than expected | A debit is missing | ❌ "… went out with no row (a debit is missing)" (not blocking) |
| **Every day of the period covered** | By the files' dates and by balances that follow each other (a day with no payment has no file) | ❌ "16/10 to 18/10 missing", with a warning that Run is allowed. **These days are also named in the Slack message.** |
| **Day after the period** | Evening payments are booked the next morning | Recommended, never required |
| **IremboPay `BPR_RW` payments on no statement** | Paid payments of the period whose invoice is on none of the uploaded statements | ❌ "43 of 179 are on no statement", with the days (not blocking) |
| **Payment file** | Covers the period | ✅ with the number of payments; a warning if it looks short |

---

## 4. Run: what happens to each credit

### 4.1 Joining the statements

- **All the statements are put together.** A row that appears in two files (same currency, reference, dates,
  description and amount) is **kept once**.
- **A credit in a file without VALUE_DATE that is also in a file with VALUE_DATE** (same BANK_REFERENCE and
  amount) is kept once, with its VALUE_DATE. For example, `sept 22 to oct 13.xlsx` together with the day files of
  12 and 13 Oct.
- **A credit belongs to the period of its `VALUE_DATE`**, even when it is in the next day's file. Credits outside
  the period are left out.

### 4.2 Finding the invoice of each credit

1. **DOC ID** is the invoice number.
2. **No DOC ID:** the invoice is looked for in IremboPay, in this order:
   1. `BANK_REFERENCE` = an IremboPay `provider_reference`
   2. `DESCRIPTION` = an IremboPay `provider_reference`
   3. `DESCRIPTION` = an IremboPay invoice number

   It is used **only when IremboPay has it**, and only when that reference belongs to one invoice. How it was found
   is written in the column **"Invoice found by"**.
3. **Still no invoice:** the credit goes **straight to Pending notification**. It is not reviewed, and its currency
   is taken from the file name. The only exception is a credit cancelled by a debit, which is still reviewed.

### 4.3 Matching with IremboPay

- **The match:** each credit is joined with its IremboPay payment by invoice number.
- **Currency:**
  - **Found in IremboPay:** the IremboPay currency of the invoice.
  - **Not found:** taken from the statement's file name ("USD" in it = USD, otherwise RWF). The reviewer can change
    it in step 5 ("Choose currency").
- **Channel**, read from the DESCRIPTION: `SZ…` = AGENT, a number = INTERNET BANKING, `FTCM…` = BRANCH, otherwise
  OTHER.

### 4.4 What goes to Exception Resolution (step 5)

**A credit not found in IremboPay:**

| Reason shown | Meaning |
|---|---|
| DOC ID not in the IremboPay payment file | The invoice is unknown in this file. Step 5's IremboPay check looks it up on checkout: "Paid" means Paid with other channel, which the reviewer confirms. |

**A credit found in IremboPay that still needs a person:**

| Reason shown | Rule |
|---|---|
| Paid through `<provider>` in IremboPay (not BPR_RW) | IremboPay shows another provider, e.g. `BPR` / BANK_TRANSFER |
| Invoice credited again: IremboPay also recorded this credit | The same invoice was credited more than once, and IremboPay has a second transaction for it |
| Invoice credited again: IremboPay has no record of this credit | The same invoice was credited more than once, but only one credit is in IremboPay. The one counted is the credit whose DESCRIPTION = the IremboPay `provider_reference`. |
| Cancelled by the debit … | The money did not stay |
| No VALUE_DATE in the statement | The SLA cannot be checked |

**The extra credits of an invoice credited more than once** are in a table of their own, **"Invoice credited more
than once"**, under the main one (added 19/10/2026). Each row is the extra credit (its bank reference and agent /
description) next to **the credit the report counted** (its bank reference, agent / description and value date). A
double payment is usually *Exclude* with a comment (refund); a wrong DOC ID is fixed in *Correct invoice*.

**In step 5,** the reviewer chooses for each row:
- the report table: Notification received, Pending notification, Paid with other channel, Failed notification, or
  Exclude
- if needed: the SLA, the corrected invoice number, the status, the provider, paid_at, the merchant, the currency
  and a comment

**Decisions are saved** (`output/.review_store.csv`). An invoice already decided in an earlier review is suggested
again. The report is written only when every row is reviewed.

### 4.5 SLA (in minutes)

- **On time:** `paid_at` (cut to the minute) is within **1 minute** of `VALUE_DATE`. The number of minutes is
  changeable in Bank Management (up to 60).
- **SLA violated:** a larger difference.
- **Cannot check:** no `VALUE_DATE`.
- **Not processed:** paid through another provider.
- **A reviewer's SLA decision** always wins.

### 4.6 The five summary tables

| Table | Rule |
|---|---|
| Overall processed payments | The invoices **on both sides** (statement and payment details, or placed by the reviewer), **each invoice once** |
| Overall notifications received by Irembo | Processed, provider `BPR_RW`, and status SETTLED / PAID / PAYOUT_INITIATED |
| SLA Violated | Processed and late (4.5) |
| Pending notifications | Processed with IremboPay payment_status NEW (an expired invoice is NEW, with the note "invoice expired") |
| Duplicate Payment | Processed, but paid through another provider (including invoices the checkout shows as Paid) |

**Processed = Notification received + Pending + Duplicate Payment**, with no overlap (all banks, 20/10/2026). The
invoice is the unique identifier, so these are **not counted**:
- credits with no invoice number (still **listed** in the Pending sheet)
- invoices only on the statement, unless the reviewer places them in a table
- the extra credits of an invoice credited more than once
- invoices whose IremboPay status is neither paid nor pending (e.g. FAILED)

- **Per currency:** each table has one row per currency (RWF, USD) and a total. The % is the share of the processed
  payments of that currency.
- **Payment per merchant:** a sixth table shows the notifications received, per merchant and currency.
- **Blocks:** a month gets one block per week plus one for the month, and the app checks that **the weeks add up to
  the month**.
- **Comparison (optional):** compares with the previous period's report, even a No-Transaction Report.

---

## 5. Step 6: the report and its delivery

**The Excel report** `BPR_recon_summary_<period>.xlsx` has these sheets:

| Sheet | Content |
|---|---|
| Summary-Daily / Weekly / Monthly | The five tables plus Payment per merchant |
| Pending notification | The rows of that table |
| Duplicate payment other channel | The rows of that table |
| BPR statement | The credits of the period with their invoice and source file |
| Payment details | Every credit with its IremboPay data, "Invoice found by", "VALUE_DATE missing", "Cancelled by", SLA minutes and status, review decision and reason |

**Automatic delivery:**
1. A **PDF** of the report is made.
2. Both files are **saved to Google Drive**. A report already there is replaced.
3. **Once Drive confirms,** the **Slack message** is sent once for that report, with:
   - the rates
   - the summary table
   - the top merchants
   - the days with no statement
   - links to the Excel and PDF

---

## 6. The whole flow at a glance

```
Statements (.xlsx) ─┐  read + balance check per file (refused if it does not add up)
                    ├─ Step 3: files follow each other · days covered · day after · BPR_RW payments on no
IremboPay (.csv) ───┘          statement (all shown, nothing blocks)
        │
        ▼
Join statements (repeated rows once) → keep the credits whose VALUE_DATE is in the period
        │
        ▼
Invoice = DOC ID, else BANK_REFERENCE/DESCRIPTION → IremboPay (only if IremboPay has it)
        │                                    └─ no invoice → Pending notification
        ▼
Match with IremboPay → currency (IremboPay, else the file name)
        │
        ├─ clean match ──────────────► Notification received / Pending / Duplicate
        └─ not found, other provider, credited again,
           cancelled, no VALUE_DATE ──► Step 5 Exception Resolution (a person decides)
        │
        ▼
SLA (1 minute) → 5 tables + merchants → Excel + PDF → Google Drive → Slack
```
