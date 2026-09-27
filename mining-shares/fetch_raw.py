"""Download raw daily data into a dataset's raw folder (needs internet).

Usage: python fetch_raw.py [majors|gold]

<raw>/prices.csv   Yahoo Finance daily open/high/low/close (split-adjusted, as
                   Yahoo serves them), dividends and stock splits per ticker
<raw>/fx.csv       Yahoo daily FX closes, units of currency per 1 USD
<raw>/targets.csv  current 12-month analyst price targets (Yahoo)
<raw>/stooq.csv    second source: Stooq daily closes (US, LSE, HK tickers)
<raw>/nasdaq.csv   second source: Nasdaq daily closes (US tickers)
<raw>/status.csv   which source returned how many rows for each ticker
"""

import io
import sys
from pathlib import Path

import pandas as pd
import requests
import yfinance as yf

HERE = Path(__file__).parent
START = "2016-12-01"
DATASETS = {
    "majors": ("companies.csv", "raw"),
    "gold": ("companies_gold.csv", "raw_gold"),
    "etf": ("companies_etf.csv", "raw_etf"),
}
FX = {"CAD": "CAD=X", "AUD": "AUD=X", "GBP": "GBP=X", "MXN": "MXN=X", "BRL": "BRL=X",
      "SAR": "SAR=X", "HKD": "HKD=X", "CNY": "CNY=X", "IDR": "IDR=X", "ZAR": "ZAR=X",
      "SEK": "SEK=X", "EUR": "EUR=X"}
UA = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/126.0 Safari/537.36", "Accept": "application/json, text/plain, */*"}


def yahoo(ticker: str) -> pd.DataFrame:
    try:
        h = yf.Ticker(ticker).history(start=START, auto_adjust=False, actions=True)
    except Exception as e:
        print(f"  yahoo {ticker}: {e}")
        return pd.DataFrame()
    if h.empty:
        return h
    h.index = h.index.tz_localize(None).date
    h.index.name = "date"
    return h


def stooq_symbol(ticker: str) -> str | None:
    if "." not in ticker and "=" not in ticker:
        return ticker.lower() + ".us"
    base, suffix = ticker.rsplit(".", 1)
    if suffix == "L":
        return base.lower() + ".uk"
    if suffix == "HK":
        return base.lstrip("0").lower() + ".hk"
    return None


def stooq(ticker: str) -> pd.DataFrame:
    sym = stooq_symbol(ticker)
    if not sym:
        return pd.DataFrame()
    try:
        r = requests.get(f"https://stooq.com/q/d/l/?s={sym}&i=d", headers=UA, timeout=30)
        df = pd.read_csv(io.StringIO(r.text))
        if "Close" not in df:
            print(f"  stooq {sym}: {r.text[:80]!r}")
            return pd.DataFrame()
        df = df[df.Date >= START]
        return pd.DataFrame({"date": df.Date, "close": df.Close, "ticker": ticker, "source_symbol": sym})
    except Exception as e:
        print(f"  stooq {sym}: {e}")
        return pd.DataFrame()


def nasdaq(ticker: str) -> pd.DataFrame:
    if "." in ticker or "=" in ticker:
        return pd.DataFrame()
    url = (f"https://api.nasdaq.com/api/quote/{ticker}/historical?assetclass=stocks"
           f"&fromdate={START}&todate=2030-01-01&limit=9999")
    try:
        rows = requests.get(url, headers=UA, timeout=30).json()["data"]["tradesTable"]["rows"]
        df = pd.DataFrame(rows)
        return pd.DataFrame({"date": pd.to_datetime(df.date).dt.date,
                             "close": df.close.str.replace(r"[$,]", "", regex=True).astype(float),
                             "ticker": ticker})
    except Exception as e:
        print(f"  nasdaq {ticker}: {e}")
        return pd.DataFrame()


def main() -> None:
    name = sys.argv[1] if len(sys.argv) > 1 else "majors"
    companies_file, raw_dir = DATASETS[name]
    raw = HERE / raw_dir
    raw.mkdir(exist_ok=True)
    companies = pd.read_csv(HERE / companies_file)
    extra = []
    if "cross_check" in companies:
        for v in companies.cross_check.dropna():
            extra += [t.strip() for t in str(v).split(";") if t.strip()]
    tickers = list(dict.fromkeys(companies.ticker.dropna().tolist() + extra))

    frames, targets, st, nd, status = [], [], [], [], []
    for t in tickers:
        h = yahoo(t)
        if not h.empty:
            cols = [c for c in ["Open", "High", "Low", "Close", "Dividends", "Stock Splits"] if c in h]
            frames.append(h[cols].assign(ticker=t).reset_index())
            try:
                pt = yf.Ticker(t).analyst_price_targets or {}
            except Exception:
                pt = {}
            targets.append({"ticker": t, **{k: pt.get(k) for k in ("current", "mean", "median", "low", "high")}})
        s, n = stooq(t), nasdaq(t)
        st.append(s)
        nd.append(n)
        status.append({"ticker": t, "yahoo_rows": len(h), "stooq_rows": len(s), "nasdaq_rows": len(n)})
        print(status[-1])

    pd.concat(frames).round(6).to_csv(raw / "prices.csv", index=False)
    pd.DataFrame(targets).to_csv(raw / "targets.csv", index=False)
    pd.DataFrame(status).to_csv(raw / "status.csv", index=False)
    for fname, parts in (("stooq.csv", st), ("nasdaq.csv", nd)):
        parts = [p for p in parts if not p.empty]
        if parts:
            pd.concat(parts).to_csv(raw / fname, index=False)

    fx = pd.DataFrame({cur: yahoo(sym)["Close"] for cur, sym in FX.items()})
    fx.index.name = "date"
    fx.sort_index().round(6).to_csv(raw / "fx.csv")
    print(f"Wrote {raw_dir}/")


if __name__ == "__main__":
    main()
