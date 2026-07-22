# Portfolio Dashboards

Two standalone tools for analyzing a stock/ETF portfolio:

1. **`True_Exposure.py`** — Portfolio True Exposure Terminal
   A browser-based (Streamlit) dashboard that computes your *look-through*
   exposure to individual stocks held indirectly through a portfolio of
   ETFs — i.e., what you actually own once ETF holdings are unwrapped and
   combined. Upload a CSV or Excel file and it produces exposure tables,
   an ETF → holding Sankey diagram, a network graph, a treemap, and
   overlap heatmaps between ETFs.

2. **`dividend_liability_dashboard.py`** — Dividend Liability Dashboard
   A desktop (Tkinter) app that tracks your portfolio's cumulative
   dividend yield and each holding's contribution to it. You enter each
   holding's portfolio weight (and optionally its yields), and it can
   auto-fetch missing dividend yields from the web (via Yahoo Finance /
   stockanalysis.com). Produces stat cards, an allocation pie chart, and a
   top-contributors bar chart, and can import/export Excel files.

Neither app requires PyCharm, VS Code, or any other IDE — everything
needed to run them is handled by the scripts in this folder.

---

## Quick start (no coding experience needed)

### macOS or Linux
1. Download/unzip this folder onto the computer.
2. Open the **Terminal** app.
3. Drag the folder into the Terminal window (this fills in the path), or
   type `cd ` followed by the folder path, then press Enter.
4. Make the launcher runnable (only needed once):
   ```
   chmod +x run.sh
   ```
5. Run it:
   ```
   ./run.sh
   ```

### Windows
1. Download/unzip this folder onto the computer.
2. Double-click **`run.bat`**.
   (If Windows shows a "Windows protected your PC" warning because the
   file was downloaded from the internet, click **More info** →
   **Run anyway**.)

Either way, the script will:
- Check that Python 3 is installed (and tell you exactly how to install
  it if it isn't).
- Create a private, self-contained environment (`.venv` folder) so
  nothing is installed system-wide and it won't conflict with anything
  else on the computer.
- Install all required packages automatically.
- Open the **Portfolio True Exposure Terminal** in your default web
  browser.
- Open the **Dividend Liability Dashboard** as its own desktop window.

You only need to double-click/run the script once per session. The
first run takes a few minutes (installing packages); every run after
that is much faster.

To stop everything: close the Dividend Liability Dashboard window, and
(on Windows) also close the black Streamlit command-window that opened
alongside it. On macOS/Linux, pressing `Ctrl+C` in Terminal, or simply
closing the Dividend Liability Dashboard window, shuts down the
Streamlit server automatically.

---

## Requirements

- **Python 3.8 or newer.** If it's not already installed, get it from
  [python.org/downloads](https://www.python.org/downloads/).
  - **Windows:** on the first install screen, check the box **"Add
    python.exe to PATH"** before clicking Install — this is what lets
    `run.bat` find it.
  - **macOS:** the installer from python.org works out of the box.
  - **Linux:** also make sure the `tkinter` GUI toolkit is installed
    (it's needed for the Dividend Liability Dashboard's window and is
    not installed by pip):
    ```
    sudo apt-get install python3-tk
    ```
- Everything else (Streamlit, pandas, plotly, yfinance, etc.) is listed
  in `requirements.txt` and gets installed automatically by the launcher
  script — no manual `pip install` needed.
- An internet connection is needed the first time you run the script
  (to install packages), and any time you use the Dividend Liability
  Dashboard's "Refresh Yields from Web" feature (it's optional — you can
  also type yields in manually).

---

## Folder contents

| File | Purpose |
|---|---|
| `True_Exposure.py` | The Streamlit true-exposure app |
| `dividend_liability_dashboard.py` | The Tkinter dividend dashboard app |
| `requirements.txt` | List of Python packages both apps need |
| `run.sh` | One-click launcher for macOS/Linux |
| `run.bat` | One-click launcher for Windows |
| `README.md` | This file |

---

## Using the True Exposure Terminal

Upload a CSV or Excel file with these columns (spacing/case doesn't
matter):

```
SL. NO.
ETF IN FUND
WEIGHT OF ETF IN FUND
ETF'S HOLDING
WEIGHT IN ETF
TRUE EXPOSURE IN PORTFOLIO
```

The app will merge and aggregate this into your true look-through
exposure per underlying stock, across all your ETFs.

## Using the Dividend Liability Dashboard

Either type in holdings manually (Ticker + Portfolio Weight %, with
optional TTM/SEC yields), or click **Import Excel** with a sheet that
has at least a Ticker (or Symbol) column and a Weight (or Weight %)
column. Any missing yields are automatically looked up from the web once
import finishes; you can also click **Refresh Yields from Web** anytime
to re-fetch them.

---

## Troubleshooting

- **"python3: command not found" / "'python' is not recognized"** —
  Python isn't installed or isn't on your PATH. Reinstall from
  python.org and, on Windows, make sure "Add python.exe to PATH" is
  checked.
- **Tkinter window doesn't open on Linux** — install it with
  `sudo apt-get install python3-tk`, then re-run the script.
- **Streamlit opens but the browser tab is blank / "connection
  refused"** — wait a few seconds and refresh; the server can take a
  moment to start on the first run.
- **Yield lookups keep failing** — this uses free, unofficial public
  data sources (Yahoo Finance via `yfinance`, and a scrape of
  stockanalysis.com as a fallback); they occasionally rate-limit or lag.
  You can always type yields in manually instead.
