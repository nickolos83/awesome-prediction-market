# Mining company share prices 2017-2026 + 2-year forward prices

- `companies.csv`: maps each entity to its listed parent and Yahoo Finance ticker.
  Codelco, LKAB and Servicio Geológico Colombiano are state-owned or government bodies, so they have no shares.
- `fetch_raw.py`: downloads daily prices, splits, FX and analyst targets into `raw/` (needs internet;
  runs on GitHub Actions via `.github/workflows/mining-shares.yml` when this file or `companies.csv` changes).
- `build_workbook.py`: builds `mining_shares_2017_2026.xlsx` from `raw/` (offline; needs pandas, openpyxl, pycel).

Workbook tabs:

| Tab | Content |
|---|---|
| Year-end close | Last close of each year in the trading currency (2026 = latest close). Highlighted cells were corrected; hover to see Yahoo's figure |
| Year-end close USD | The same prices converted at the year-end FX rate (formulas) |
| 2y forward | Price a 2-year future would have had, by target year: `close × e^((r − q)·2)` (formulas) |
| 2y forward USD | Local forward ÷ 2-year FX forward (formulas) |
| Dividend yield, Rates 2y, FX year-end, FX 2y forward | Inputs for the formulas. Rates are approximate; edit them and the workbook updates |
| Detail | Open, high, low, close, average, dividends, forward, the actual price 2 years later and the forward error |
| Analyst targets (now) | Today's 12-month consensus. Free sources don't keep historical targets |

Prices are **as traded**. Yahoo rescales history after splits, bonus issues, stock dividends and demergers.
The build undoes that: it uses Yahoo's split data plus the demerger factors in `HIDDEN_ADJUSTMENTS`
(BHP–Woodside 2022, Anglo American–Thungela 2021 and Valterra 2025). LSE prices are in pence (GBp).
The forward price reflects carry (interest minus dividends) only. It is not a forecast.
