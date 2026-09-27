# VA Hub — all scraping engines in one window

## Start it
```
pip install -r requirements.txt
python va_hub.py
```

## Use it
1. Click an engine (Ctrl/Shift-click to pick several).
2. Press **Run**. Each engine opens its own date/county window, same as before.
3. Several engines can run at the same time. The **Status** column shows Running / Done / Failed.
4. The box at the bottom shows live output for the selected engine. Every run is also saved in `logs/`.
5. **Open Sheet** opens the Google Sheet that engine writes to.

## Engines
| Engine | Input | Google Sheet |
|---|---|---|
| TN Ledger – Probate (Court Notices) | one date | `1KIcOYi8…` |
| TN Ledger – Foreclosures | one date | `1DInbhKw…` |
| TN Ledger – Foreclosure Comparison | date range | `1k0j25pN…` |
| TN Public Notice – Probate & Foreclosure | date range + county | `1278HJam…` |
| TN Courts – Civil Cases (tncrtinfo) | date range + county + court, optional login | `1UJp8UvQ…` |

## Add a new engine
Put the script in `engines/<name>/main.py` and add an entry to `engines.json`
(name, category, description, cwd, script, sheet_id). It shows up in the hub next launch.
Engines get the Google key path in the `VA_KEYS_FILE` environment variable.
