# Bank Reconciliation

A local reconciliation application for comparing bank statements with IremboPay payment data and generating daily, weekly, or monthly reconciliation reports in Excel and PDF format.

## Overview

This project helps reconcile:
- I&M and Ecobank bank statements
- RWF and USD transactions
- IremboPay payment exports
- exception review and reporting

The app runs locally on the computer. Uploaded files are read in memory and are not kept after processing. Generated reports are saved on Google Drive (optional) or downloaded; only the exception decisions and the run log are kept on this computer.

## Features
- Reconcile by day, week, or month
- Validate whether all required bank statement files are present
- Match payment rows against IremboPay data
- Review missing or unresolved exceptions
- Export standard and custom reconciliation reports
- Save reports on Google Drive (by bank and month) and find them again in Report History

## Project structure

```text
Bank Reconciliation/
├── README.md
├── Setup (run once).command
├── Start Reconciliation.command
├── .streamlit/
├── Bank of Africe/
│   ├── Payment details.csv
│   └── ...bank statements...
├── output/
│   ├── day/
│   ├── month/
│   ├── run_log.csv
│   └── ...generated reports...
├── reconciliation_app/
│   ├── app.py
│   ├── engine.py
│   ├── assets/
│   ├── banks/
│   └── static/
└── ...supporting files
```

## Requirements

- macOS
- Python 3.10+
- Streamlit
- pandas
- openpyxl / Excel export support

## First-time setup

1. Open the project folder.
2. Double-click `Setup (run once).command`.
3. Wait for the setup process to finish.

> If macOS blocks the script as an unidentified developer app, right-click the file, choose Open, and then confirm once.

## Run the app

You can start the app in either of these ways:

1. Double-click `Start Reconciliation.command`
2. Or run the app manually:

```bash
cd "/Users/divinngenzi/Documents/Bank Reconciliation"
python3 -m streamlit run reconciliation_app/app.py
```

## How to use it

1. Select the bank and the reconciliation period.
2. Upload the bank statement files from your computer. For the IremboPay payment file, choose **This computer**
   (upload) or **Google Drive** (pick a file in the IremboPay folder or its month subfolders, then **☁ Add from
   Drive**). Both go through the same checks: a file of another bank or outside the period is refused.
3. Check that all required files are covered.
4. Run the reconciliation.
5. Review any exceptions if transactions are missing or unresolved.
6. Download the report (Excel or PDF) and/or click **☁ Save to Google Drive**. If the report is already on Drive,
   the app asks: **Replace the existing report** or **Save as a new version** (`_v2`, `_v3`, …). A file is shown as
   saved only after Google Drive confirms it; if an upload fails, the report stays in the app and **↻ Retry** sends
   only the files that failed.

Saved reports are found again with **🗂 Report History** in the side panel: choose the bank and the month, select a
report, then **👁 View** or **⬇ Get file** (a report is downloaded only when you ask). Older reports still in
`output/` are listed under *This computer (older reports)*; they are not copied anywhere.

## Notes
- Reconciliation periods must be complete before they can be processed.
- Generated reports are **not kept on this computer**: they stay in the app's memory until you download them or
  save them on Google Drive. While a report is made, a temporary file is used in the session's own private
  temporary folder and deleted at once.
- Exception decisions (`output/.review_store.csv`) and the run log (`output/run_log.csv`) stay on this computer.
- Nothing is sent to the internet unless Google Drive is used.

## Google Drive setup

Google Drive is optional: without it the app works as before (IremboPay files uploaded from the computer, reports
downloaded). The app uses Drive **as you** (OAuth, with your Irembo Google account), so reports are created in your
own **My Drive → Bank Reconciliation** folders. A service account can be used instead (fallback, see below).

1. **Google Cloud** (once; done for project `bank-reconciliation-511109`, Irembo organization):
   Google Auth Platform with **Audience = Internal**; **Data access**: `drive.readonly` and `drive.file`;
   **Clients → Desktop app** client, its JSON saved as `~/.config/bank-reconciliation/oauth-client.json`
   (outside the project; `chmod 600`). If Irembo's Workspace policy blocks the app, the sign-in page shows
   *Access blocked / admin_policy_enforced*: IT must then trust the OAuth client (Admin console → Security →
   API controls).
2. **Settings**: copy `drive_settings.example.toml` to `drive_settings.toml` (same folder; ignored by Git) and fill in
   the two folder IDs (the last part of a folder's link, `https://drive.google.com/drive/folders/<ID>`):
   - `reports_root_folder_id`: **My Drive → Bank Reconciliation → Reconciliation Reports**. Reports are saved in
     `<bank>/<YYYY-MM>/` inside it (YYYY-MM = the month the period starts in); the app makes these folders.
   - `irembopay_folder_id`: **My Drive → Bank Reconciliation → IremboPay File**. Files in it and in its subfolders
     (up to 3 levels, e.g. `I&M / 09/2026`) are listed with their folder path.

   Environment variables can be used instead: `BANK_RECON_AUTH`, `BANK_RECON_OAUTH_CLIENT`, `BANK_RECON_TOKEN`,
   `BANK_RECON_CREDENTIALS`, `BANK_RECON_REPORTS_FOLDER_ID`, `BANK_RECON_IREMBOPAY_FOLDER_ID`.
3. **Sign in**: side panel → **Google Drive** → **Sign in with Google**. Your browser opens Google's page: choose your
   Irembo account and allow. The sign-in is kept in `~/.config/bank-reconciliation/token.json` (readable only by
   you, outside the project) and renewed automatically. **Sign out** revokes and deletes it; you can also remove the
   app at *myaccount.google.com → Security → Third-party apps*. If the sign-in expires or is revoked, the app says
   *Sign in again*, and nothing is uploaded or shown as saved until you do.
4. **Permissions asked** (the narrowest that work):

   | Permission | Why |
   |---|---|
   | `drive.readonly` | list and download the IremboPay exports you put in your folders |
   | `drive.file` | create, verify and replace the reports (and their folders) the app itself creates |

   With `drive.file` the app can only change files it created: a report put on Drive by hand cannot be replaced by
   the app (it shows an error; nothing is lost).
5. **Test it**: side panel → **Google Drive** → **Test connection** (what your account may do in each folder) and
   **Upload test**: a small file with a unique name is uploaded to *Reconciliation Reports*, checked (size and MD5
   checksum), read back and moved to the bin; only that file is touched. If a check fails, the file is kept for
   investigation. Do not rely on Drive until both are ✅.

**Service account (fallback):** set `auth = "service_account"` and put the key at `credentials_file`. Share the
folders with the service account's e-mail. A service account has no storage of its own, so it **cannot save reports
in a My Drive folder** (only in a Shared Drive, as *Content manager*); it can still read the IremboPay folder.

## Developer notes

The main logic is split across the application files:

- `reconciliation_app/app.py`: Streamlit UI and steps
- `reconciliation_app/engine.py`: reconciliation logic, validation, and report generation
- `reconciliation_app/banks/`: bank-specific parsing and reconciliation rules
- `reconciliation_app/config.py`: Google Drive settings (paths and folder IDs only, never the credentials)
- `reconciliation_app/integrations/google_drive.py`: every Google Drive call (list, download, upload, access check)
- `reconciliation_app/services/report_history.py`: saving reports on Drive (duplicates, versions) and Report History
- `reconciliation_app/services/session_files.py`: the private temporary folder of each browser session
## License

This project is intended for internal use and local reconciliation workflows.

## Support

For changes to the workflow or new bank integrations, update the bank-specific logic in the `reconciliation_app/banks/` folder and extend the reconciliation rules in `reconciliation_app/engine.py`.
