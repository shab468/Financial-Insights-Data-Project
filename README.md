# Financial Market Insights Dashboard (Automated)

A fast, resume-ready project inspired by Bloomberg Analytics workflows. It demonstrates:
- Data ingestion (auto-fetch via yfinance OR swap CSVs)
- Core analytics (daily % change, 10/30-day moving averages, simple volatility, 5D return)
- Client-ready visuals and summary insights in Excel
- A one-command automation script

## Project Structure
```
market_insights_dashboard/
├─ data/
│  ├─ AAPL.csv
│  ├─ JPM.csv
│  └─ SPY.csv
├─ Financial_Market_Insights_Dashboard.xlsx   # generated
├─ update_dashboard.py                         # builds Excel from CSVs
├─ fetch_and_build.py                          # fetches via yfinance, then builds
└─ requirements.txt
```

## Option A: Full Automation (One Command)

1) Install Python 3.10+ and dependencies:
```bash
cd market_insights_dashboard
pip install -r requirements.txt
```

2) Run end-to-end fetch + build:
```bash
python fetch_and_build.py --tickers AAPL,SPY,JPM --period 3mo --interval 1d
```
This will download real market data via **yfinance**, save CSVs into `./data`, and rebuild the Excel dashboard automatically.

## Option B: Manual CSVs → Build

1) Replace sample CSVs with real ones (Yahoo Finance → Historical Data → Download).
   Keep columns **Date, Close** and add a **Ticker** column, e.g. `AAPL`.
2) Build the dashboard:
```bash
python update_dashboard.py
```

## What the Dashboard Includes
- **RawData**: Combined time series (Date, Ticker, Close)
- **Metrics**: Daily % change, MA10, MA30, 10-day rolling volatility (annualized approx)
- **Summary**: Total return, 5-day return, latest close, volatility per ticker
- **Dashboard**: Line chart (price trend) + bar chart (total return)

## Resume Bullet
- Built an **automated market insights dashboard** in Excel; fetched real equity data with Python (yfinance), computed **daily returns, moving averages, rolling volatility**, and generated **client-ready charts and KPIs**.

## Scheduling (Optional)
- **macOS/Linux (cron)**: Run `crontab -e` and add (weekdays 6pm):
```
0 18 * * 1-5 /usr/bin/python3 /path/to/market_insights_dashboard/fetch_and_build.py >> /path/to/market_insights_dashboard/run.log 2>&1
```
- **Windows Task Scheduler**:
  - Program: `python`
  - Args: `C:\path\to\market_insights_dashboard\fetch_and_build.py`
  - Start in: `C:\path\to\market_insights_dashboard\`

