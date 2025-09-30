
"""
Rebuilds the Financial Market Insights Dashboard Excel file from CSVs in ./data

Usage:
    python update_dashboard.py

Expectations:
- ./data contains one or more CSVs with columns: Date, Ticker, Close
- Produces ./Financial_Market_Insights_Dashboard.xlsx
- Creates sheets: RawData, Metrics, Summary, Dashboard (with charts)
"""
import os
import pandas as pd
from datetime import datetime
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.chart import LineChart, Reference, BarChart
from openpyxl.chart.axis import DateAxis

BASE_DIR = os.path.dirname(__file__)
DATA_DIR = os.path.join(BASE_DIR, "data")
OUT_XLSX = os.path.join(BASE_DIR, "Financial_Market_Insights_Dashboard.xlsx")

def load_data():
    frames = []
    for fname in os.listdir(DATA_DIR):
        if fname.lower().endswith(".csv"):
            fp = os.path.join(DATA_DIR, fname)
            df = pd.read_csv(fp)
            # basic normalization
            df.columns = [c.strip().title() for c in df.columns]
            # Expect Date, Ticker, Close
            if not set(["Date","Ticker","Close"]).issubset(df.columns):
                continue
            df["Date"] = pd.to_datetime(df["Date"]).dt.date
            df = df.sort_values(["Ticker","Date"])
            frames.append(df[["Date","Ticker","Close"]])
    if not frames:
        raise RuntimeError("No CSVs with columns Date,Ticker,Close found in ./data")
    return pd.concat(frames, ignore_index=True).sort_values(["Ticker","Date"])

def compute_metrics(prices):
    # group by ticker and compute daily % change, 10/30 MA, rolling 10-day volatility
    prices = prices.copy()
    prices["Date"] = pd.to_datetime(prices["Date"])
    metrics = []
    for t, g in prices.groupby("Ticker"):
        g = g.sort_values("Date").reset_index(drop=True)
        g["Pct_Change"] = g["Close"].pct_change()
        g["MA_10"] = g["Close"].rolling(window=10).mean()
        g["MA_30"] = g["Close"].rolling(window=30).mean()
        g["Vol_10"] = g["Pct_Change"].rolling(window=10).std() * (252 ** 0.5)  # annualized approx
        metrics.append(g)
    m = pd.concat(metrics, ignore_index=True)
    m["Date"] = m["Date"].dt.date
    return m

def build_workbook(raw, metrics):
    # If exists, overwrite
    if os.path.exists(OUT_XLSX):
        os.remove(OUT_XLSX)
    wb = Workbook()
    ws_raw = wb.active
    ws_raw.title = "RawData"

    # Write RawData
    for r in dataframe_to_rows(raw, index=False, header=True):
        ws_raw.append(r)

    # Metrics sheet
    ws_met = wb.create_sheet("Metrics")
    for r in dataframe_to_rows(metrics, index=False, header=True):
        ws_met.append(r)

    # Summary: period performance and latest stats per ticker
    perf = []
    for t, g in metrics.groupby("Ticker"):
        g = g.sort_values("Date")
        last_close = g["Close"].iloc[-1] if not g["Close"].empty else None
        if len(g) >= 2:
            first_close = g["Close"].iloc[0]
            total_ret = (last_close / first_close) - 1 if first_close else None
        else:
            total_ret = None
        if len(g) >= 6:
            ret_5d = (g["Close"].iloc[-1] / g["Close"].iloc[-6]) - 1
        else:
            ret_5d = None
        vol10 = g["Vol_10"].iloc[-1] if not g["Vol_10"].dropna().empty else None
        perf.append({"Ticker": t, "Total_Return": total_ret, "Return_5D": ret_5d, "Vol_10": vol10, "Last_Close": last_close})
    summary = pd.DataFrame(perf)

    ws_sum = wb.create_sheet("Summary")
    for r in dataframe_to_rows(summary, index=False, header=True):
        ws_sum.append(r)

    # Dashboard sheet with charts
    ws_dash = wb.create_sheet("Dashboard")
    ws_dash["A1"] = "Financial Market Insights Dashboard"
    ws_dash["A2"] = f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}"

    # Line chart of Close over time for all tickers
    pivot = metrics.pivot(index="Date", columns="Ticker", values="Close").reset_index()
    start_row = 5
    for i, col in enumerate(pivot.columns, start=1):
        ws_dash.cell(row=start_row, column=i, value=col)
    for r_i in range(len(pivot)):
        for c_i, col in enumerate(pivot.columns, start=1):
            ws_dash.cell(row=start_row + 1 + r_i, column=c_i, value=pivot.iloc[r_i, c_i-1])

    chart = LineChart()
    chart.title = "Price Trend (Close)"
    chart.y_axis.title = "Price"
    chart.x_axis = DateAxis()
    chart.x_axis.title = "Date"
    chart.x_axis.number_format = "yyyy-mm-dd"
    chart.x_axis.majorTimeUnit = "days"

    data_ref = Reference(ws_dash, min_col=2, min_row=start_row, max_col=1 + len(pivot.columns)-1,
                         max_row=start_row + 1 + len(pivot) - 1)
    cats_ref = Reference(ws_dash, min_col=1, min_row=start_row + 1, max_row=start_row + 1 + len(pivot) - 1)
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)
    chart.height = 15
    chart.width = 30
    ws_dash.add_chart(chart, "A4")

    # Bar chart of Total Return by ticker
    sum_start_row = start_row + len(pivot) + 4
    sum_cols = ["Ticker","Total_Return","Return_5D","Vol_10","Last_Close"]
    for j, col in enumerate(sum_cols, start=1):
        ws_dash.cell(row=sum_start_row, column=j, value=col)
    for i in range(len(summary)):
        for j, col in enumerate(sum_cols, start=1):
            val = summary.iloc[i][col]
            # robust write: numbers as floats, strings as strings, NaN -> None
            try:
                from pandas.api.types import is_number
                is_num = is_number(val)
            except Exception:
                is_num = isinstance(val, (int, float))
            if is_num:
                ws_dash.cell(row=sum_start_row + 1 + i, column=j, value=float(val))
            else:
                # allow Ticker string, handle NaN
                if pd.isna(val):
                    ws_dash.cell(row=sum_start_row + 1 + i, column=j, value=None)
                else:
                    ws_dash.cell(row=sum_start_row + 1 + i, column=j, value=str(val))

    bar = BarChart()
    bar.title = "Total Return (Period)"
    bar.y_axis.title = "Return"
    data_ref2 = Reference(ws_dash, min_col=2, min_row=sum_start_row, max_col=2, max_row=sum_start_row + len(summary))
    cats_ref2 = Reference(ws_dash, min_col=1, min_row=sum_start_row + 1, max_row=sum_start_row + len(summary))
    bar.add_data(data_ref2, titles_from_data=True)
    bar.set_categories(cats_ref2)
    bar.height = 12
    bar.width = 20
    ws_dash.add_chart(bar, "A25")

    wb.save(OUT_XLSX)

def main():
    raw = load_data()
    metrics = compute_metrics(raw)
    build_workbook(raw, metrics)
    print(f"Dashboard written to: {OUT_XLSX}")

if __name__ == "__main__":
    main()
