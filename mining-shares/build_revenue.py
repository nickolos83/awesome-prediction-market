"""Build annual revenue workbooks from revenue/<dataset>_revenue.csv (no internet needed).

Usage: python build_revenue.py [majors|gold]
  majors -> mining_revenue_2017_2026.xlsx
  gold   -> gold_miners_revenue_2017_2026.xlsx

Primary source is StockAnalysis (last 5 fiscal years); Yahoo Finance fills gaps and is
the cross-check. Revenue is in millions of the currency the company reports in. The
USD tab divides by the average FX rate over the 12 months of each fiscal year.
A fiscal year is labelled by the calendar year it ends in (a year ending in January
counts as the previous year).
"""

import sys
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter

from build_workbook import (BOLD, CORRECTED, FONT, GREEN, Font, set_widths, store_cached_values,
                            style_header)

HERE = Path(__file__).parent
YEARS = list(range(2017, 2027))
DATASETS = {
    "majors": ("companies.csv", "mining_revenue_2017_2026"),
    "gold": ("companies_gold.csv", "gold_miners_revenue_2017_2026"),
}
# Differences between the two sources that were looked into.
NOTES = {
    "VALE3.SA": "Vale reports in USD. StockAnalysis converts to BRL (used here, converted back on the USD tab); "
                "Yahoo shows the USD figures but labels them BRL.",
    "AAUC.TO": "Yahoo has 2024 and 2025 swapped; StockAnalysis matches Allied Gold's reported figures.",
}


def fiscal_label(end: str) -> int:
    d = pd.Timestamp(end)
    return d.year - 1 if d.month == 1 else d.year


def load_fx() -> pd.DataFrame:
    frames = [pd.read_csv(HERE / f, parse_dates=["date"]).set_index("date") for f in ("raw/fx.csv", "raw_gold/fx.csv")]
    fx = pd.concat(frames, axis=1)
    fx = fx.T.groupby(level=0).first().T.sort_index().ffill()
    fx["USD"] = 1.0
    return fx


def avg_rate(fx: pd.DataFrame, currency: str, end: str) -> float | None:
    if currency not in fx:
        return None
    e = pd.Timestamp(end)
    window = fx.loc[e - pd.DateOffset(years=1) + pd.Timedelta(days=1): e, currency].dropna()
    return float(window.mean()) if len(window) else None


def main() -> None:
    name = sys.argv[1] if len(sys.argv) > 1 else "majors"
    companies_file, out_name = DATASETS[name]
    companies = pd.read_csv(HERE / companies_file, dtype=str, keep_default_na=False)
    listed = (companies[companies.ticker != ""].drop_duplicates("ticker")
              .sort_values("ticker").reset_index(drop=True))
    rev = pd.read_csv(HERE / "revenue" / f"{name}_revenue.csv")
    rev = rev.dropna(subset=["revenue"])
    rev["year"] = rev.fiscal_year_end.map(fiscal_label)
    fx = load_fx()

    # Pick one figure per ticker and fiscal year: StockAnalysis first, Yahoo otherwise.
    chosen, checks = {}, []
    for (t, y), g in rev.groupby(["ticker", "year"]):
        by = {r.source: r for r in g.itertuples()}
        sa, yh = by.get("stockanalysis"), by.get("yahoo")
        pick = sa or yh
        cur = pick.currency if isinstance(pick.currency, str) else (yh.currency if yh else None)
        chosen[(t, y)] = (pick.revenue, cur, pick.fiscal_year_end, pick.source)
        if sa and yh:
            diff = sa.revenue / yh.revenue - 1 if yh.revenue else float("inf")
            status = "match" if abs(diff) <= 0.01 else ("small difference (1-5%)" if abs(diff) <= 0.05 else
                                                        "DIFFERENT")
            if t in NOTES and status != "match":
                status = "explained: " + NOTES[t]
            checks.append([t, y, sa.fiscal_year_end, round(sa.revenue / 1e6, 1), round(yh.revenue / 1e6, 1),
                           round(diff * 100, 1), status])

    wb = Workbook()
    ws = wb.active
    ws.title = "Entities"
    ws.append(list(companies.columns))
    style_header(ws)
    for row in companies.itertuples(index=False):
        ws.append(list(row))
    set_widths(ws, {"A": 40, "B": 34, "C": 13, "D": 14, "E": 9, "F": 30, "G": 60})

    n = len(listed)
    year_col = {y: get_column_letter(5 + i) for i, y in enumerate(YEARS)}

    def ticker_sheet(title: str) -> "Worksheet":
        s = wb.create_sheet(title)
        s.append(["ticker", "company", "currency", "fiscal year ends"] + [str(y) for y in YEARS])
        style_header(s)
        set_widths(s, {"A": 13, "B": 32, "C": 9, "D": 14, **{year_col[y]: 11 for y in YEARS}})
        s.freeze_panes = "E2"
        return s

    local = ticker_sheet("Revenue")
    usd = ticker_sheet("Revenue USD")
    rates = ticker_sheet("FX rate used")
    for i, c in listed.iterrows():
        r = 2 + i
        picks = {y: chosen[(c.ticker, y)] for y in YEARS if (c.ticker, y) in chosen}
        cur = next((p[1] for p in reversed(list(picks.values())) if p[1]), "")
        fy_end = pd.Timestamp(list(picks.values())[-1][2]).strftime("%d %b") if picks else ""
        for s in (local, usd, rates):
            s.cell(r, 1, c.ticker)
            s.cell(r, 2, c.listed_parent)
            s.cell(r, 3, "USD" if s is usd else cur)
            s.cell(r, 4, fy_end)
        if not picks:
            local.cell(r, 1).comment = Comment("No revenue data found (delisted, or no free source).", "Claude")
        for y, (value, vcur, end, source) in picks.items():
            cell = local[f"{year_col[y]}{r}"]
            cell.value = round(value / 1e6, 2)
            cell.number_format = "#,##0.0"
            note = f"Fiscal year ended {end}. Source: {'StockAnalysis' if source == 'stockanalysis' else 'Yahoo Finance'}."
            check = next((k for k in checks if k[0] == c.ticker and k[1] == y), None)
            if check and check[-1] == "DIFFERENT":
                cell.fill = CORRECTED
                note += f" Yahoo shows {check[4]:,.1f} ({check[5]:+.1f}%): likely a restatement, check the annual report."
            cell.comment = Comment(note, "Claude")
            rate = avg_rate(fx, vcur, end)
            if rate:
                rates[f"{year_col[y]}{r}"] = round(rate, 6)
                rates[f"{year_col[y]}{r}"].number_format = "0.0000"
                ref = f"{year_col[y]}{r}"
                usd[ref] = f"=IF(OR('Revenue'!{ref}=\"\",'FX rate used'!{ref}=\"\"),\"\",'Revenue'!{ref}/'FX rate used'!{ref})"
                usd[ref].number_format = "#,##0.0"
                usd[ref].font = GREEN
    notes = {
        local: "Annual revenue in millions of the reporting currency. Column = calendar year the fiscal year ends in "
               "(fiscal years ending in January count as the previous year). Blank = no free data "
               "(free sources keep only the last ~5 fiscal years). Hover a cell for its exact period and source.",
        usd: "Revenue / average FX rate over that fiscal year (from the 'FX rate used' tab), in USD millions.",
        rates: "Average units of reporting currency per 1 USD over the 12 months of each fiscal year (Yahoo FX closes).",
    }
    for s, text in notes.items():
        s.cell(n + 3, 1, text).font = FONT

    ver = wb.create_sheet("Verification")
    df = pd.DataFrame(checks, columns=["ticker", "fiscal_year", "period_end", "stockanalysis_m", "yahoo_m",
                                       "difference_%", "status"])
    counts = df.status.str.split(":").str[0].value_counts()
    ver.append(["StockAnalysis vs Yahoo Finance, same fiscal year (millions, reporting currency)"])
    ver.append([f"{len(df)} comparisons across {df.ticker.nunique()} tickers: "
                + ", ".join(f"{k}: {v}" for k, v in counts.items())])
    ver.append(["DIFFERENT usually means one source restated the year (e.g. a sold mine moved to discontinued "
                "operations); the Revenue tab shows StockAnalysis and flags the cell."])
    ver.append([])
    ver.append(list(df.columns))
    style_header(ver, 5)
    for row in df.sort_values(["status", "ticker", "fiscal_year"]).itertuples(index=False):
        ver.append(list(row))
        if row.status == "DIFFERENT":
            for c in ver[ver.max_row]:
                c.fill = CORRECTED
    ver.cell(1, 1).font = BOLD
    set_widths(ver, {"A": 13, "C": 12, "D": 16, "E": 12, "G": 60})

    for s in wb:
        for row in s.iter_rows():
            for c in row:
                if c.font == Font():
                    c.font = FONT
    out = HERE / f"{out_name}.xlsx"
    wb.save(out)
    store_cached_values(out, [s.title for s in wb])
    print(f"Wrote {out.name}")


if __name__ == "__main__":
    main()
