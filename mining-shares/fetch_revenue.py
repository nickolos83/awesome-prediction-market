"""Download annual revenue for both company lists (needs internet).

Usage: python fetch_revenue.py
Writes revenue/<dataset>_revenue.csv: ticker, source, fiscal_year_end, revenue, currency

Sources (both free, last ~5 fiscal years):
  stockanalysis  StockAnalysis.com annual income statement
  yahoo          Yahoo Finance annual income statement
"""

import re
from pathlib import Path

import pandas as pd
import requests
import yfinance as yf

HERE = Path(__file__).parent
OUT = HERE / "revenue"
DATASETS = {"majors": "companies.csv", "gold": "companies_gold.csv"}
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/126.0 Safari/537.36", "Accept-Language": "en-US,en;q=0.9"}
SA_EXCHANGE = {"TO": "tsx", "V": "tsxv", "CN": "cse", "AX": "asx", "L": "lon", "HK": "hkg", "SS": "sha",
               "SZ": "she", "JK": "idx", "SR": "tadawul", "MX": "bmv", "SA": "bvmf", "JO": "jse"}


def sa_url(ticker: str) -> str | None:
    base, _, suffix = ticker.partition(".")
    if not suffix:
        return f"https://stockanalysis.com/stocks/{base.lower()}/financials/"
    if suffix not in SA_EXCHANGE:
        return None
    if suffix == "HK":
        base = base.lstrip("0")
    return f"https://stockanalysis.com/quote/{SA_EXCHANGE[suffix]}/{base}/financials/"


def stockanalysis_revenue(ticker: str) -> list[dict]:
    url = sa_url(ticker)
    if not url:
        return []
    r = requests.get(url, headers=UA, timeout=30)
    if r.status_code != 200:
        raise ValueError(f"HTTP {r.status_code}")
    dates = re.search(r"datekey:\[([^\]]*)\]", r.text)
    revenue = re.search(r"revenue:\[([^\]]*)\]", r.text)
    currency = re.search(r'currency:"([A-Z]{3})"', r.text)
    if not dates or not revenue:
        raise ValueError("no revenue table on page")
    ds = [d.strip('"') for d in dates.group(1).split(",")]
    vs = revenue.group(1).split(",")
    return [{"ticker": ticker, "source": "stockanalysis", "fiscal_year_end": d,
             "revenue": float(v) if v not in ("", "null") else None,
             "currency": currency.group(1) if currency else None, "url": url}
            for d, v in zip(ds, vs)]


def yahoo_revenue(ticker: str) -> list[dict]:
    t = yf.Ticker(ticker)
    try:
        cur = (t.info or {}).get("financialCurrency")
    except Exception:
        cur = None
    stmt = t.income_stmt
    if stmt is None or stmt.empty:
        return []
    for label in ("Total Revenue", "Operating Revenue"):
        if label in stmt.index:
            s = stmt.loc[label].dropna()
            return [{"ticker": ticker, "source": "yahoo", "fiscal_year_end": str(d.date()), "revenue": float(v),
                     "currency": cur, "url": ""} for d, v in s.items()]
    return []


def main() -> None:
    OUT.mkdir(exist_ok=True)
    for name, file in DATASETS.items():
        companies = pd.read_csv(HERE / file, dtype=str, keep_default_na=False)
        tickers = [t for t in dict.fromkeys(companies.ticker) if t]
        rows, log = [], []
        for t in tickers:
            entry = {"ticker": t}
            for source, fn in (("stockanalysis", stockanalysis_revenue), ("yahoo", yahoo_revenue)):
                try:
                    got = fn(t)
                    rows += got
                    entry[source] = len(got)
                except Exception as e:
                    entry[source] = f"error: {str(e)[:80]}"
            log.append(entry)
            print(entry, flush=True)
        pd.DataFrame(rows).to_csv(OUT / f"{name}_revenue.csv", index=False)
        pd.DataFrame(log).to_csv(OUT / f"{name}_revenue_log.csv", index=False)


if __name__ == "__main__":
    main()
