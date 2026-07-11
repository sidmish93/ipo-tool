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
cd ipo_tool
pip install -r requirements.txt
python -m playwright install chromium
```

## Run

```bash
python app.py
```

Open <http://127.0.0.1:5000>, choose the dates, and click **Generate Excel**.
The finished file is saved under `ipo_tool/output/` and offered as a download.

> A full run visits every qualifying company's pages live, so expect a few minutes.
