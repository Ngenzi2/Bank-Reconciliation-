#!/bin/bash
# Opens the Bank Reconciliation app (I&M, Ecobank, ...) in the browser. Keep this window open while you use the app.
cd "$(dirname "$0")"
URL="http://localhost:8501"
PY=/opt/anaconda3/bin/python
[ -x "$PY" ] || PY=$(command -v python3)

if lsof -iTCP:8501 -sTCP:LISTEN >/dev/null 2>&1; then
  echo "The app is already running - opening it in the browser."
  open "$URL"; exit 0
fi
if ! "$PY" -c "import streamlit, python_calamine, openpyxl, reportlab" 2>/dev/null; then
  echo "The app is not set up yet. Double-click 'Setup (run once)' first."
  read -p "Press Enter to close..."; exit 1
fi

echo "Starting Bank Reconciliation..."
echo "The app opens in your browser. Keep this window open; close it to stop the app."
( sleep 3; open "$URL" ) &
"$PY" -m streamlit run reconciliation_app/app.py
