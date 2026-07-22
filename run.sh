#!/usr/bin/env bash
# ============================================================================
# run.sh — one-click launcher for the two portfolio tools (macOS / Linux)
#
# What this does, step by step:
#   1. Finds a working Python 3 install (or tells you how to get one).
#   2. Creates a private virtual environment in ./.venv (first run only) so
#      nothing is installed system-wide.
#   3. Installs every package listed in requirements.txt into that venv.
#   4. Launches the Streamlit dashboard (True_Exposure.py) — this opens in
#      your web browser automatically.
#   5. Launches the Tkinter desktop dashboard (dividend_liability_dashboard.py)
#      — this opens as its own window.
#   6. When you close the Tkinter window, the Streamlit server is shut down
#      automatically.
#
# How to run it:
#   1. Open Terminal.
#   2. Drag this folder into the Terminal window (or `cd` into it).
#   3. Run:  ./run.sh
#      (If you get a "permission denied" error, run:  chmod +x run.sh
#       once, then try again.)
# ============================================================================

set -e

# Move into the folder this script lives in, so it works no matter where
# it's launched from.
cd "$(dirname "${BASH_SOURCE[0]}")"

echo "============================================================"
echo " Portfolio Dashboards — Setup"
echo "============================================================"

# ----------------------------------------------------------------------
# 1. Find Python 3
# ----------------------------------------------------------------------
PYTHON_BIN=""
for candidate in python3 python; do
    if command -v "$candidate" >/dev/null 2>&1; then
        if "$candidate" -c 'import sys; sys.exit(0 if sys.version_info >= (3,8) else 1)' 2>/dev/null; then
            PYTHON_BIN="$candidate"
            break
        fi
    fi
done

if [ -z "$PYTHON_BIN" ]; then
    echo ""
    echo "ERROR: Python 3.8+ was not found on this computer."
    echo "Install it from https://www.python.org/downloads/ (check the box"
    echo "'Add Python to PATH' during install on Windows) and then re-run"
    echo "this script."
    exit 1
fi

echo "Using Python: $($PYTHON_BIN --version)"

# ----------------------------------------------------------------------
# 2. Create the virtual environment (only if it doesn't already exist)
# ----------------------------------------------------------------------
if [ ! -d ".venv" ]; then
    echo "Creating virtual environment in .venv (first run only)..."
    "$PYTHON_BIN" -m venv .venv
fi

# Activate it
# shellcheck disable=SC1091
source .venv/bin/activate

# ----------------------------------------------------------------------
# 3. Install dependencies
# ----------------------------------------------------------------------
echo "Installing/updating required packages (this can take a few minutes"
echo "the first time)..."
pip install --quiet --upgrade pip
pip install --quiet -r requirements.txt

# tkinter check — it's part of the standard library, but some Linux
# distros ship Python without it.
if ! python -c "import tkinter" >/dev/null 2>&1; then
    echo ""
    echo "WARNING: The 'tkinter' module is missing. The Dividend Liability"
    echo "Dashboard needs it. On Debian/Ubuntu, install it with:"
    echo "    sudo apt-get install python3-tk"
    echo "then re-run this script."
fi

# ----------------------------------------------------------------------
# 4. Launch the Streamlit app (True Exposure Terminal) in the background
# ----------------------------------------------------------------------
echo ""
echo "Starting the Portfolio True Exposure Terminal (Streamlit)..."
echo "It will open automatically in your web browser."
streamlit run True_Exposure.py &
STREAMLIT_PID=$!

# Make sure the Streamlit server is stopped when this script exits
# (whether normally, via Ctrl+C, or by closing the dashboard window).
cleanup() {
    echo ""
    echo "Shutting down the Streamlit server..."
    kill "$STREAMLIT_PID" 2>/dev/null || true
}
trap cleanup EXIT

# Give Streamlit a moment to boot before opening the second app.
sleep 3

# ----------------------------------------------------------------------
# 5. Launch the Tkinter app (Dividend Liability Dashboard) in the foreground
# ----------------------------------------------------------------------
echo "Starting the Dividend Liability Dashboard (desktop window)..."
python dividend_liability_dashboard.py

echo ""
echo "Dividend Liability Dashboard closed."
