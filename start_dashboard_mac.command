#!/bin/bash
# Crisdel Toll Dashboard launcher (Mac)
# Double-click this file in Finder to set up (first run only) and open the dashboard.
# If macOS blocks it the first time: right-click -> Open, then confirm.

cd "$(dirname "$0")"

if ! command -v python3 &> /dev/null; then
    echo "Python 3 is not installed."
    echo "Install it from https://www.python.org/downloads/ then double-click this file again."
    read -p "Press Enter to close..."
    exit 1
fi

if [ ! -d ".venv" ]; then
    echo "First-time setup: creating a private Python environment for this app..."
    python3 -m venv .venv
fi

echo "Checking dependencies..."
".venv/bin/pip" install --quiet --upgrade pip
".venv/bin/pip" install --quiet -r requirements.txt

echo ""
echo "Starting the Crisdel Toll Dashboard..."
echo "Your browser will open automatically. Close this window when you're done."
echo ""

".venv/bin/streamlit" run app.py
