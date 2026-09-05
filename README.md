# IPO Shareholding Excel Generator

A small local web app: pick a **start** and **end** date, click **Generate Excel**,
and get a two-sheet workbook of IPO companies (SEBI "Final Offer Documents filed with ROC")
with market cap **> 3000 cr**, plus their detailed promoter and public shareholding.

## What it produces

**Sheet 1 – `IPO Companies >3000cr`**
`S.No | Company Name | BSE Ticker | BSE Scrip Code | NSE Ticker | ISIN | Mcap Full (Cr.) | Latest Shareholding Quarter | Promoter statement link | Public statement link`

**Sheet 2 – `Detailed Shareholding`**
For every company:
- Each **Promoter / Promoter Group** entity whose "% of (A+B+C2)" > 0.
- Each **non-bold public shareholder** whose "% of (A+B+C2)" > 0, tagged with the
  bold category heading it appears under (e.g. *Mutual Funds*, *Foreign Companies*).

## How it works

The whole pipeline runs inside a real headless Chromium (via Playwright) because:
- SEBI blocks plain HTTP POSTs (WAF), so the date-filtered list is read in a browser.
- The "bold vs non-bold" public-shareholder distinction only exists in the rendered page.

Data sources: SEBI filings page, BSE JSON APIs (search, `StockTrading`, shareholding
quarter, statement pages) and the NSE `globalSearch` API for the NSE ticker.

## Setup (one time)

```bash
cd ipo-tool
pip install -r requirements.txt
```

BSE blocks Playwright's bundled Chromium (403 Access Denied). The app needs
**Google Chrome** (or Edge) and only continues if BSE returns 200.

- **This PC:** install Chrome, then `python app.py`.
- **Render / Docker:** the `Dockerfile` installs real Chrome. The Render
  service must use the **Docker** environment (not native Python) and
  **2 GB RAM**. Starter/512 MB dies when Chrome opens BSE. Redeploy after
  pulling this repo.

## Run

```bash
python app.py
```

Open <http://127.0.0.1:5000>, choose the dates, and click **Generate Excel**.
The finished file is saved under `ipo_tool/output/` and offered as a download.

To look up specific listed companies (skipping the SEBI date range), type one
name or several separated by semicolons:

```
Delhivery; Lodha Developers; Vedanta
```

Each result shows live BSE market cap, shares outstanding, and each holder's
latest-quarter shares / % / rupee holding. After the quarter-end date, NSE and
BSE block/bulk trades are applied in share counts (sells subtract, buys add);
adjusted % is adjusted shares ÷ outstanding. A June SHP is as-of 30 June, so
the trade window is 1 July through today — the same rule for every quarter.

> A full run visits every qualifying company's pages live, so expect a few minutes.
