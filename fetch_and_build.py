#!/usr/bin/env python3
"""
Automate the entire project in one command:
- Fetch real OHLC data for tickers using yfinance
- Save per-ticker CSVs with columns [Date, Ticker, Close] into ./data
- Rebuild the Excel dashboard by calling update_dashboard.py

Usage examples:
    python fetch_and_build.py
    python fetch_and_build.py --tickers AAPL,SPY,JPM --period 3mo --interval 1d
    python fetch_and_build.py --tickers MSFT,GOOGL --period 6mo

Requirements:
    pip install -r requirements.txt
"""
import argparse
import sys
import os
from datetime import datetime
from pathlib import Path

try:
    import pandas as pd
except Exception as e:
    print("Missing pandas. Please run: pip install -r requirements.txt")
    sys.exit(1)

try:
    import yfinance as yf
except Exception as e:
    print("Missing yfinance. Please run: pip install -r requirements.txt")
    sys.exit(1)

# Allow importing update_dashboard.py as a module
BASE_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE_DIR))
try:
    import update_dashboard as builder
except Exception as e:
    print("Could not import update_dashboard.py. Make sure it exists in the same folder.")
    raise

def fetch_csvs(tickers, period="3mo", interval="1d", out_dir=None):
    out = Path(out_dir or (BASE_DIR / "data"))
    out.mkdir(parents=True, exist_ok=True)
    for t in tickers:
        t = t.strip().upper()
        print(f"[{datetime.now().strftime('%H:%M:%S')}] Fetching {t} ({period}, {interval}) ...")
        df = yf.download(t, period=period, interval=interval, progress=False)
        if df.empty:
            print(f"  Warning: No data returned for {t}. Skipping.")
            continue
        df = df.reset_index()
        date_col = "Date" if "Date" in df.columns else "Datetime"
        df["Date"] = pd.to_datetime(df[date_col]).dt.date
        df["Ticker"] = t
        if "Close" not in df.columns:
            if "Adj Close" in df.columns:
                df["Close"] = df["Adj Close"]
            else:
                raise RuntimeError(f"{t}: No Close/Adj Close column found.")
        out_fp = out / f"{t}.csv"
        df[["Date","Ticker","Close"]].to_csv(out_fp, index=False)
        print(f"  Saved -> {out_fp}")
    return out

def main():
    parser = argparse.ArgumentParser(description="Fetch real prices and rebuild the dashboard")
    parser.add_argument("--tickers", type=str, default="AAPL,SPY,JPM", help="Comma-separated tickers")
    parser.add_argument("--period", type=str, default="3mo", help="yfinance period (e.g., 1mo,3mo,6mo,1y,2y,max)")
    parser.add_argument("--interval", type=str, default="1d", help="yfinance interval (e.g., 1d,1wk,1mo)")
    args = parser.parse_args()

    tickers = [x for x in args.tickers.split(",") if x.strip()]
    fetch_csvs(tickers, period=args.period, interval=args.interval, out_dir=BASE_DIR/"data")

    print(f"[{datetime.now().strftime('%H:%M:%S')}] Building Excel dashboard ...")
    builder.main()
    print(f"[{datetime.now().strftime('%H:%M:%S')}] Done. Open: {BASE_DIR/'Financial_Market_Insights_Dashboard.xlsx'}")

if __name__ == "__main__":
    main()
