# Bank Reconciliation

A local reconciliation application for comparing bank statements with IremboPay payment data and generating daily, weekly, or monthly reconciliation reports in Excel and PDF format.

## Overview

This project helps reconcile:
- I&M and Ecobank bank statements
- RWF and USD transactions
- IremboPay payment exports
- exception review and reporting

The app runs locally on the computer and only saves the generated reconciliation outputs and review decisions. Uploaded files are read in memory and are not kept after processing.

## Features
- Reconcile by day, week, or month
- Validate whether all required bank statement files are present
- Match payment rows against IremboPay data
- Review missing or unresolved exceptions
- Export standard and custom reconciliation reports
- Save reports under the local `output/` directory

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
2. Upload the required bank statement files and the IremboPay payment CSV.
3. Check that all required files are covered.
4. Run the reconciliation.
5. Review any exceptions if transactions are missing or unresolved.
6. Export the report to Excel or PDF.

## Notes
- Reconciliation periods must be complete before they can be processed.
- The app keeps reports locally in the `output/` folder.
- Reports and exception decisions are saved on the same machine where the app runs.
- Nothing is sent to the internet during normal use.

## Developer notes

The main logic is split across the application files:

- `reconciliation_app/app.py`: Streamlit UI and steps
- `reconciliation_app/engine.py`: reconciliation logic, validation, and report generation
- `reconciliation_app/banks/`: bank-specific parsing and reconciliation rules

## License

This project is intended for internal use and local reconciliation workflows.

## Support

For changes to the workflow or new bank integrations, update the bank-specific logic in the `reconciliation_app/banks/` folder and extend the reconciliation rules in `reconciliation_app/engine.py`.
