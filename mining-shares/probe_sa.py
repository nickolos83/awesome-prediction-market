"""Temporary probe: save StockAnalysis financials pages to inspect their format."""
from pathlib import Path
import requests
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/126.0 Safari/537.36", "Accept-Language": "en-US,en;q=0.9"}
out = Path(__file__).parent / "revenue" / "probe"
out.mkdir(parents=True, exist_ok=True)
for name, url in {"nem": "https://stockanalysis.com/stocks/nem/financials/",
                  "rms": "https://stockanalysis.com/quote/asx/RMS/financials/",
                  "agi": "https://stockanalysis.com/quote/tsx/AGI/financials/",
                  "gor": "https://stockanalysis.com/quote/asx/GOR/financials/"}.items():
    r = requests.get(url, headers=UA, timeout=30)
    (out / f"{name}.html").write_text(f"<!-- HTTP {r.status_code} -->\n" + r.text)
    print(name, r.status_code, len(r.text))
