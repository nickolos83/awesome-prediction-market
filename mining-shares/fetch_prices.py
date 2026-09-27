"""Yearly share prices (2017-2026) and 2-year forward ("futures") prices
for the listed parents of the mining entities in companies.csv.

Usage:  pip install yfinance pandas openpyxl && python fetch_prices.py
Output: mining_shares_2017_2026.xlsx and .csv next to this script.

Forward price: F = S * exp((r - q) * 2), where S is the year-end close,
r is the 2-year government yield of the quote currency and q the trailing
12-month dividend yield. This is the no-arbitrage price a 2-year single-stock
future would trade at. It reflects carry, not a view on where the price goes.
"""

import math
from pathlib import Path

import pandas as pd
import yfinance as yf

HERE = Path(__file__).parent
FIRST_YEAR, LAST_YEAR = 2017, 2026
HORIZON_YEARS = 2

# Approximate year-end 2-year government yields in % (for Saudi Arabia, a
# SAIBOR-style proxy). Replace these with exact figures if you need precision.
# 2026 uses the 2025 values until a real year-end figure exists.
TWO_YEAR_YIELD = {
    "USD": [1.89, 2.49, 1.58, 0.13, 0.73, 4.43, 4.25, 4.24, 3.48],
    "CAD": [1.69, 1.86, 1.70, 0.20, 0.95, 4.05, 3.88, 2.93, 2.60],
    "AUD": [2.00, 1.90, 0.90, 0.10, 0.60, 3.50, 3.70, 3.85, 3.90],
    "GBP": [0.45, 0.75, 0.55, -0.15, 0.70, 3.60, 4.00, 4.40, 3.75],
    "MXN": [7.60, 8.60, 6.90, 4.30, 7.20, 10.50, 10.30, 9.90, 7.30],
    "BRL": [7.50, 7.20, 5.30, 4.30, 11.30, 13.40, 10.30, 15.40, 13.90],
    "SAR": [2.00, 2.90, 2.30, 0.90, 1.20, 5.00, 5.40, 5.20, 4.60],
    "HKD": [1.30, 1.80, 1.70, 0.20, 0.60, 4.00, 3.60, 3.40, 2.60],
    "CNY": [3.50, 2.70, 2.60, 2.70, 2.30, 2.30, 2.20, 1.10, 1.40],
    "IDR": [5.80, 7.70, 5.70, 4.10, 3.90, 5.90, 6.40, 6.90, 5.00],
}


def rate(currency: str, year: int) -> float:
    series = TWO_YEAR_YIELD[currency.replace("GBp", "GBP")]
    return series[min(year - FIRST_YEAR, len(series) - 1)] / 100


def yearly_rows(ticker: str, currency: str) -> list[dict]:
    t = yf.Ticker(ticker)
    hist = t.history(start=f"{FIRST_YEAR}-01-01", auto_adjust=False, actions=True)
    if hist.empty:
        return []
    hist.index = hist.index.tz_localize(None)
    rows = []
    for year in range(FIRST_YEAR, LAST_YEAR + 1):
        y = hist[hist.index.year == year]
        if y.empty:
            continue
        close = y["Close"].iloc[-1]
        divs = y["Dividends"].sum()
        q = divs / close if close else 0.0
        r = rate(currency, year)
        rows.append({
            "year": year,
            "last_trading_day": y.index[-1].date(),
            "open": round(y["Open"].iloc[0], 4),
            "high": round(y["High"].max(), 4),
            "low": round(y["Low"].min(), 4),
            "close": round(close, 4),
            "average_close": round(y["Close"].mean(), 4),
            "dividends_in_year": round(divs, 4),
            "dividend_yield_pct": round(q * 100, 2),
            "rate_2y_pct": round(r * 100, 2),
            f"forward_{HORIZON_YEARS}y_price": round(close * math.exp((r - q) * HORIZON_YEARS), 4),
            f"forward_for_year": year + HORIZON_YEARS,
        })
    by_year = {row["year"]: row for row in rows}
    for row in rows:
        later = by_year.get(row["year"] + HORIZON_YEARS)
        actual = later["close"] if later else None
        row["actual_close_in_forward_year"] = actual
        fwd = row[f"forward_{HORIZON_YEARS}y_price"]
        row["forward_error_pct"] = round((actual / fwd - 1) * 100, 1) if actual else None
    return rows


def analyst_targets(ticker: str) -> dict:
    """Current 12-month analyst consensus (no free source keeps the history)."""
    try:
        t = yf.Ticker(ticker).analyst_price_targets or {}
    except Exception:
        t = {}
    return {k: t.get(k) for k in ("current", "mean", "median", "low", "high")}


def main() -> None:
    companies = pd.read_csv(HERE / "companies.csv")
    listed = companies.dropna(subset=["ticker"]).drop_duplicates("ticker")

    prices, targets = [], []
    for _, c in listed.iterrows():
        print(f"{c.ticker:12} {c.listed_parent}")
        for row in yearly_rows(c.ticker, c.currency):
            prices.append({"ticker": c.ticker, "company": c.listed_parent,
                           "currency": c.currency, **row})
        targets.append({"ticker": c.ticker, "company": c.listed_parent,
                        "currency": c.currency, **analyst_targets(c.ticker)})

    prices_df = pd.DataFrame(prices)
    wide = prices_df.pivot(index=["ticker", "company", "currency"],
                           columns="year", values="close").reset_index()
    fwd_col = f"forward_{HORIZON_YEARS}y_price"
    fwd_wide = prices_df.pivot(index=["ticker", "company", "currency"],
                               columns="forward_for_year", values=fwd_col).reset_index()

    prices_df.to_csv(HERE / "mining_shares_2017_2026.csv", index=False)
    with pd.ExcelWriter(HERE / "mining_shares_2017_2026.xlsx") as xl:
        companies.to_excel(xl, sheet_name="Entities", index=False)
        wide.to_excel(xl, sheet_name="Year-end close", index=False)
        fwd_wide.to_excel(xl, sheet_name=f"{HORIZON_YEARS}y forward (by target yr)", index=False)
        prices_df.to_excel(xl, sheet_name="Detail", index=False)
        pd.DataFrame(targets).to_excel(xl, sheet_name="Analyst targets (now)", index=False)
    print("Wrote mining_shares_2017_2026.xlsx / .csv")


if __name__ == "__main__":
    main()
