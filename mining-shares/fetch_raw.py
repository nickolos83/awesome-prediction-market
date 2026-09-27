"""Download raw daily data from Yahoo Finance into raw/ (needs internet).

raw/prices.csv  daily open/high/low/close (split-adjusted, as Yahoo serves
                them), dividends and stock splits for every ticker
raw/fx.csv      daily FX closes, units of currency per 1 USD
raw/targets.csv current 12-month analyst price targets
"""

from pathlib import Path

import pandas as pd
import yfinance as yf

HERE = Path(__file__).parent
RAW = HERE / "raw"
START = "2016-12-01"

# Second listings used only to cross-check the primary ticker.
CROSS_CHECK = ["ABX.TO", "NGLOY", "601899.SS", "BHP", "RIO", "VALE", "SCCO"]
FX = {"CAD": "CAD=X", "AUD": "AUD=X", "GBP": "GBP=X", "MXN": "MXN=X", "BRL": "BRL=X",
      "SAR": "SAR=X", "HKD": "HKD=X", "CNY": "CNY=X", "IDR": "IDR=X"}


def history(ticker: str) -> pd.DataFrame:
    h = yf.Ticker(ticker).history(start=START, auto_adjust=False, actions=True)
    if h.empty:
        return h
    h.index = h.index.tz_localize(None).date
    h.index.name = "date"
    return h


def main() -> None:
    RAW.mkdir(exist_ok=True)
    companies = pd.read_csv(HERE / "companies.csv")
    tickers = list(dict.fromkeys(companies.ticker.dropna().tolist() + CROSS_CHECK))

    frames, targets = [], []
    for t in tickers:
        print(t)
        h = history(t)
        if h.empty:
            continue
        cols = [c for c in ["Open", "High", "Low", "Close", "Dividends", "Stock Splits"] if c in h]
        frames.append(h[cols].assign(ticker=t).reset_index())
        try:
            pt = yf.Ticker(t).analyst_price_targets or {}
        except Exception:
            pt = {}
        targets.append({"ticker": t, **{k: pt.get(k) for k in ("current", "mean", "median", "low", "high")}})
    pd.concat(frames).round(6).to_csv(RAW / "prices.csv", index=False)
    pd.DataFrame(targets).to_csv(RAW / "targets.csv", index=False)

    fx = pd.DataFrame({cur: history(sym)["Close"] for cur, sym in FX.items()})
    fx.index.name = "date"
    fx.sort_index().round(6).to_csv(RAW / "fx.csv")
    print("Wrote raw/")


if __name__ == "__main__":
    main()
