#!/bin/bash
# One-time setup for the Bank Reconciliation app: installs the Python packages it needs.
cd "$(dirname "$0")"
PY=/opt/anaconda3/bin/python
[ -x "$PY" ] || PY=$(command -v python3)
if [ -z "$PY" ]; then
  echo "Python was not found. Please install Anaconda (https://www.anaconda.com/download) and run this again."
  read -p "Press Enter to close..." ; exit 1
fi
echo "Using Python: $PY"
echo "Installing the packages (this takes a minute)..."
"$PY" -m pip install --upgrade streamlit pandas openpyxl python-calamine reportlab google-api-python-client google-auth google-auth-oauthlib
if [ $? -eq 0 ]; then
  echo ""
  echo "Setup finished. From now on, double-click 'Start Reconciliation' to open the app."
else
  echo ""
  echo "Setup did not finish. Check the internet connection and try again."
fi
read -p "Press Enter to close..."
