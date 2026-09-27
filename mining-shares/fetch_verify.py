"""Download year-end closes from sources independent of Yahoo (needs internet).

Usage: python fetch_verify.py [gold|majors]
Writes <raw>/verify.csv: ticker, source, year, date, close (as the source serves it).

Sources:
  nasdaq       api.nasdaq.com historical quotes (US listings)
  tmx          TMX Money GraphQL (TSX / TSXV listings)
  cnbc         CNBC chart API (ASX, LSE, HKEX)
  eastmoney    Eastmoney kline API, unadjusted (Shanghai, Hong Kong)
"""

import io
import sys
import time
from pathlib import Path

import pandas as pd
import requests

HERE = Path(__file__).parent
DATASETS = {"majors": ("companies.csv", "raw"), "gold": ("companies_gold.csv", "raw_gold")}
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/126.0 Safari/537.36",
      "Accept": "application/json, text/plain, */*", "Accept-Language": "en-US,en;q=0.9"}
YEARS = range(2016, 2027)
MW_COUNTRY = {"AX": "au", "TO": "ca", "V": "ca", "L": "uk", "HK": "hk", "CN": "ca"}


def year_end(df: pd.DataFrame, ticker: str, source: str) -> list[dict]:
    df = df.dropna().sort_values("date")
    df["year"] = pd.to_datetime(df.date).dt.year
    last = df.groupby("year").tail(1)
    return [{"ticker": ticker, "source": source, "year": int(r.year), "date": str(r.date)[:10],
             "close": float(r.close)} for r in last.itertuples()]


def nasdaq(ticker: str, symbol: str) -> list[dict]:
    url = (f"https://api.nasdaq.com/api/quote/{symbol}/historical?assetclass=stocks"
           f"&fromdate=2016-12-01&todate=2026-12-31&limit=9999")
    rows = requests.get(url, headers=UA, timeout=30).json()["data"]["tradesTable"]["rows"]
    df = pd.DataFrame(rows)
    df = pd.DataFrame({"date": pd.to_datetime(df.date).dt.date,
                       "close": df.close.str.replace(r"[$,]", "", regex=True).astype(float)})
    return year_end(df, ticker, "nasdaq")


TMX_QUERY = """query getTimeSeriesData($symbol: String!, $freq: String, $interval: Int, $start: String,
  $end: String, $startDateTime: Int, $endDateTime: Int) {
  getTimeSeriesData(symbol: $symbol, freq: $freq, interval: $interval, start: $start, end: $end,
    startDateTime: $startDateTime, endDateTime: $endDateTime) { dateTime open high low close volume }
}"""


def tmx(ticker: str, symbol: str) -> list[dict]:
    body = {"operationName": "getTimeSeriesData", "query": TMX_QUERY,
            "variables": {"symbol": symbol, "freq": "day", "start": "2016-12-01", "end": "2026-12-31"}}
    headers = {**UA, "Content-Type": "application/json", "Origin": "https://money.tmx.com",
               "Referer": "https://money.tmx.com/"}
    r = requests.post("https://app-money.tmx.com/graphql", json=body, headers=headers, timeout=60).json()
    data = (r.get("data") or {}).get("getTimeSeriesData") or []
    if not data:
        raise ValueError(str(r)[:200])
    df = pd.DataFrame(data)
    df["date"] = pd.to_datetime(df.dateTime, utc=True).dt.tz_convert("America/Toronto").dt.date
    return year_end(df[["date", "close"]], ticker, "tmx")


def marketwatch(ticker: str, symbol: str, country: str | None) -> list[dict]:
    out = []
    for y in YEARS:
        url = (f"https://www.marketwatch.com/investing/stock/{symbol}/downloaddatapartial"
               f"?startdate=12/01/{y}%2000:00:00&enddate=12/31/{y}%2023:59:59&daterange=d30"
               f"&frequency=p1d&csvdownload=true&downloadpartial=false&newdates=false"
               + (f"&countrycode={country}" if country else ""))
        r = requests.get(url, headers=UA, timeout=30)
        if r.status_code != 200 or "Date" not in r.text[:50]:
            if y == YEARS[0]:
                raise ValueError(f"HTTP {r.status_code}: {r.text[:120]!r}")
            continue
        df = pd.read_csv(io.StringIO(r.text))
        df = pd.DataFrame({"date": pd.to_datetime(df.Date, format="%m/%d/%Y").dt.date,
                           "close": pd.to_numeric(df.Close.astype(str).str.replace(",", ""), errors="coerce")})
        out += year_end(df, ticker, "marketwatch")
        time.sleep(0.5)
    return out


def cnbc(ticker: str, symbol: str) -> list[dict]:
    out, err = [], ""
    for mode in ("unadjusted", "adjusted"):
        url = (f"https://ts-api.cnbc.com/harmony/app/bars/{symbol}/1D/20161201000000/20261231000000/"
               f"{mode}/EST5EDT.json")
        try:
            bars = requests.get(url, headers=UA, timeout=30).json()["barData"]["priceBars"]
            df = pd.DataFrame(bars)
            df = pd.DataFrame({"date": pd.to_datetime(df.tradeTime.astype(str).str[:8]).dt.date,
                               "close": df.close.astype(float)})
            out += [{**r, "source": f"cnbc_{mode}"} for r in year_end(df, ticker, "cnbc")]
        except Exception as e:
            err += f"{mode}: {str(e)[:80]} "
    if not out:
        raise ValueError(err)
    return out


def eastmoney(ticker: str, secid: str) -> list[dict]:
    url = ("https://push2his.eastmoney.com/api/qt/stock/kline/get?secid=" + secid +
           "&fields1=f1,f2,f3,f4,f5,f6&fields2=f51,f52,f53,f54,f55&klt=101&fqt=0&beg=20161201&end=20261231")
    klines = requests.get(url, headers=UA, timeout=30).json()["data"]["klines"]
    df = pd.DataFrame([k.split(",")[:3] for k in klines], columns=["date", "open", "close"])
    df = pd.DataFrame({"date": pd.to_datetime(df.date).dt.date, "close": df.close.astype(float)})
    return year_end(df, ticker, "eastmoney")


CNBC_SUFFIX = {"AX": "-AU", "L": "-GB", "HK": "-HK"}


def main() -> None:
    name = sys.argv[1] if len(sys.argv) > 1 else "gold"
    companies_file, raw_dir = DATASETS[name]
    companies = pd.read_csv(HERE / companies_file, dtype=str, keep_default_na=False)
    tickers = [t for t in dict.fromkeys(companies.ticker) if t]
    rows, log = [], []
    for t in tickers:
        base, _, suffix = t.partition(".")
        jobs = []
        if not suffix:
            jobs.append(("nasdaq", lambda t=t, b=base: nasdaq(t, b)))
        if suffix in ("TO", "V"):
            jobs.append(("tmx", lambda t=t, b=base: tmx(t, b)))
        if suffix in CNBC_SUFFIX:
            sym = (base.lstrip("0") if suffix == "HK" else base) + CNBC_SUFFIX[suffix]
            jobs.append(("cnbc", lambda t=t, s=sym: cnbc(t, s)))
        if suffix == "SS":
            jobs.append(("eastmoney", lambda t=t, b=base: eastmoney(t, "1." + b)))
        if suffix == "HK":
            jobs.append(("eastmoney", lambda t=t, b=base: eastmoney(t, "116." + b.zfill(5))))
        for source, job in jobs:
            try:
                got = job()
                rows += got
                log.append({"ticker": t, "source": source, "years": len(got), "error": ""})
            except Exception as e:
                log.append({"ticker": t, "source": source, "years": 0, "error": str(e)[:200]})
            print(log[-1], flush=True)
    raw = HERE / raw_dir
    pd.DataFrame(rows).to_csv(raw / "verify.csv", index=False)
    pd.DataFrame(log).to_csv(raw / "verify_log.csv", index=False)


if __name__ == "__main__":
    main()
