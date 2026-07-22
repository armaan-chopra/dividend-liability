@echo off
REM ============================================================================
REM run.bat — one-click launcher for the two portfolio tools (Windows)
REM
REM Double-click this file (or right-click it and choose "Run") to:
REM   1. Find Python on this computer.
REM   2. Create a private virtual environment in .venv (first run only).
REM   3. Install all required packages into it.
REM   4. Launch the Streamlit dashboard (opens in your web browser).
REM   5. Launch the Dividend Liability Dashboard (opens as a desktop window).
REM ============================================================================

cd /d "%~dp0"

echo ============================================================
echo  Portfolio Dashboards - Setup
echo ============================================================

where python >nul 2>nul
if errorlevel 1 (
    echo.
    echo ERROR: Python was not found on this computer.
    echo Install it from https://www.python.org/downloads/
    echo IMPORTANT: on the first install screen, check the box that says
    echo "Add python.exe to PATH" - then re-run this file.
    pause
    exit /b 1
)

echo Using: 
python --version

if not exist ".venv" (
    echo Creating virtual environment in .venv (first run only)...
    python -m venv .venv
)

call .venv\Scripts\activate.bat

echo Installing/updating required packages (this can take a few minutes the first time)...
python -m pip install --quiet --upgrade pip
pip install --quiet -r requirements.txt

echo.
echo Starting the Portfolio True Exposure Terminal (Streamlit)...
echo It will open automatically in your web browser. Leave that black
echo window open in the background - closing it will stop the dashboard.
start "True Exposure Terminal" cmd /k streamlit run True_Exposure.py

echo Waiting a few seconds for Streamlit to boot...
timeout /t 3 /nobreak >nul

echo Starting the Dividend Liability Dashboard (desktop window)...
python dividend_liability_dashboard.py

echo.
echo Dividend Liability Dashboard closed.
echo You can close the Streamlit window (the other black window) manually now.
pause
