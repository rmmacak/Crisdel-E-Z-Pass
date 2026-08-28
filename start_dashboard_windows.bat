@echo off
REM Crisdel Toll Dashboard launcher (Windows)
REM Double-click this file to set up (first run only) and open the dashboard.

cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
    echo Python is not installed, or not on your PATH.
    echo Install it from https://www.python.org/downloads/
    echo IMPORTANT: during install, check the box "Add python.exe to PATH".
    echo Then double-click this file again.
    pause
    exit /b 1
)

if not exist ".venv\" (
    echo First-time setup: creating a private Python environment for this app...
    python -m venv .venv
)

echo Checking dependencies...
".venv\Scripts\pip.exe" install --quiet --upgrade pip
".venv\Scripts\pip.exe" install --quiet -r requirements.txt

echo.
echo Starting the Crisdel Toll Dashboard...
echo Your browser will open automatically. Close this window when you're done.
echo.

".venv\Scripts\streamlit.exe" run app.py

pause
