"""Download annual revenue for both company lists (needs internet).

Usage: python fetch_revenue.py
Writes revenue/<dataset>_revenue.csv with one row per ticker, source and fiscal year:
  ticker, source, fiscal_year_end, revenue, currency, form

Sources:
  sec    SEC EDGAR XBRL company facts (full history for US / SEC filers, incl. 20-F and 40-F)
  yahoo  Yahoo Finance income statement (last ~4 fiscal years)
"""

from pathlib import Path

import pandas as pd
import requests
import yfinance as yf

HERE = Path(__file__).parent
OUT = HERE / "revenue"
SEC_UA = {"User-Agent": "MiningSharesResearch github-actions@users.noreply.github.com",
          "Accept-Encoding": "gzip, deflate"}
DATASETS = {"majors": "companies.csv", "gold": "companies_gold.csv"}
# US listing used to find a company in SEC EDGAR when the primary ticker is not a US one.
SEC_TICKER = {
    "BHP.AX": "BHP", "RIO.L": "RIO", "VALE3.SA": "VALE", "B": "B",
    "AGI.TO": "AGI", "BTO.TO": "BTG", "CG.TO": "CGAU", "ELD.TO": "EGO", "EQX.TO": "EQX", "FVI.TO": "FSM",
    "IMG.TO": "IAG", "K.TO": "KGC", "IAU.TO": "IAUX", "NFGC.TO": "NFGC", "OGG.V": "OGG", "OGC.TO": "OGC",
    "SSRM": "SSRM", "AU": "AU", "GFI": "GFI", "HMY": "HMY", "MUX": "MUX", "AAUC.TO": "AAUC",
}
REVENUE_TAGS = [
    ("ifrs-full", "Revenue"),
    ("ifrs-full", "RevenueFromContractsWithCustomers"),
    ("us-gaap", "Revenues"),
    ("us-gaap", "RevenueFromContractWithCustomerExcludingAssessedTax"),
    ("us-gaap", "SalesRevenueNet"),
    ("us-gaap", "SalesRevenueGoodsNet"),
]


def sec_cik_map() -> dict:
    errors = []
    r = requests.get("https://www.sec.gov/files/company_tickers.json", headers=SEC_UA, timeout=60)
    if r.status_code == 200 and r.text.lstrip().startswith("{"):
        return {v["ticker"].upper(): int(v["cik_str"]) for v in r.json().values()}
    errors.append(f"company_tickers.json HTTP {r.status_code}: {r.text[:300]!r}")
    r = requests.get("https://www.sec.gov/include/ticker.txt", headers=SEC_UA, timeout=60)
    if r.status_code == 200 and "\t" in r.text[:200]:
        return {t.upper(): int(c) for t, c in (line.split("\t") for line in r.text.split("\n") if "\t" in line)}
    errors.append(f"ticker.txt HTTP {r.status_code}: {r.text[:300]!r}")
    raise ValueError(" | ".join(errors))


def sec_revenue(ticker: str, cik: int) -> list[dict]:
    facts = requests.get(f"https://data.sec.gov/api/xbrl/companyfacts/CIK{cik:010d}.json",
                         headers=SEC_UA, timeout=60).json()["facts"]
    rows = []
    for ns, tag in REVENUE_TAGS:
        for unit, items in facts.get(ns, {}).get(tag, {}).get("units", {}).items():
            for it in items:
                if it.get("form", "").split("/")[0] not in ("10-K", "20-F", "40-F") or "start" not in it:
                    continue
                days = (pd.Timestamp(it["end"]) - pd.Timestamp(it["start"])).days
                if not 350 <= days <= 380:
                    continue
                rows.append({"ticker": ticker, "source": "sec", "fiscal_year_end": it["end"],
                             "revenue": float(it["val"]), "currency": unit, "form": it["form"],
                             "filed": it["filed"], "tag": f"{ns}:{tag}"})
    if not rows:
        return []
    df = pd.DataFrame(rows)
    # Prefer the most recent filing for each period, then the first tag in REVENUE_TAGS order.
    order = {f"{ns}:{tag}": i for i, (ns, tag) in enumerate(REVENUE_TAGS)}
    df["tag_rank"] = df.tag.map(order)
    df = df.sort_values(["fiscal_year_end", "tag_rank", "filed"], ascending=[True, True, False])
    return df.groupby("fiscal_year_end").head(1).drop(columns="tag_rank").to_dict("records")


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
                     "currency": cur, "form": "", "filed": "", "tag": label} for d, v in s.items()]
    return []


def main() -> None:
    OUT.mkdir(exist_ok=True)
    try:
        ciks = sec_cik_map()
    except Exception as e:
        print("SEC ticker map failed:", e)
        (OUT / "sec_error.txt").write_text(str(e))
        ciks = {}
    for name, file in DATASETS.items():
        companies = pd.read_csv(HERE / file, dtype=str, keep_default_na=False)
        tickers = [t for t in dict.fromkeys(companies.ticker) if t]
        rows, log = [], []
        for t in tickers:
            us = SEC_TICKER.get(t, t if "." not in t else None)
            got_sec, got_y = [], []
            if us and us.upper() in ciks:
                try:
                    got_sec = sec_revenue(t, ciks[us.upper()])
                except Exception as e:
                    print(t, "sec", e)
            try:
                got_y = yahoo_revenue(t)
            except Exception as e:
                print(t, "yahoo", e)
            rows += got_sec + got_y
            log.append({"ticker": t, "sec_ticker": us, "sec_cik": ciks.get((us or "").upper()),
                        "sec_years": len(got_sec), "yahoo_years": len(got_y)})
            print(log[-1], flush=True)
        pd.DataFrame(rows).to_csv(OUT / f"{name}_revenue.csv", index=False)
        pd.DataFrame(log).to_csv(OUT / f"{name}_revenue_log.csv", index=False)


if __name__ == "__main__":
    main()
