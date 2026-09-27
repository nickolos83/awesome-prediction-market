# Mining company share prices 2017-2026 + 2-year forward prices

- `companies.csv`: maps each entity to its listed parent and Yahoo Finance ticker.
  Codelco, LKAB and Servicio Geológico Colombiano are state-owned or government bodies, so they have no shares.
- `fetch_prices.py`: downloads daily prices from Yahoo Finance (via `yfinance`) and writes
  `mining_shares_2017_2026.xlsx` with these sheets:
  - **Year-end close**: one row per company, one column per year.
  - **2y forward (by target yr)**: the price a 2-year future would have implied for that year, set two years earlier.
    For example, the 2019 column is `close_2017 × e^((r − q)·2)`.
  - **Detail**: open, high, low, close, average, dividends, rate, forward price, the actual price 2 years later and the forward error.
  - **Analyst targets (now)**: today's 12-month consensus. Free sources don't keep historical targets.

```
pip install yfinance pandas openpyxl
python fetch_prices.py
```

Notes: prices are in the quote currency, and LSE prices are in pence (GBp). The forward price reflects carry
(interest minus dividends) only. It is not a forecast. The 2-year yields in the script are approximate year-end values.
