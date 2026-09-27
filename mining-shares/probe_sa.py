"""Temporary probe: old Wayback Machine snapshots of StockAnalysis financials pages."""
import re
from pathlib import Path
import requests
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/126.0 Safari/537.36"}
out = Path(__file__).parent / "revenue" / "probe"
out.mkdir(parents=True, exist_ok=True)
pages = {"nem": "stockanalysis.com/stocks/nem/financials/", "rms": "stockanalysis.com/quote/asx/RMS/financials/",
         "agi": "stockanalysis.com/quote/tsx/AGI/financials/"}
for name, url in pages.items():
    for ts in ("20210601", "20220601", "20230601"):
        try:
            a = requests.get(f"https://archive.org/wayback/available?url={url}&timestamp={ts}", headers=UA,
                             timeout=60).json()
            snap = a.get("archived_snapshots", {}).get("closest", {})
            if not snap:
                print(name, ts, "no snapshot"); continue
            raw = re.sub(r"/web/(\d+)/", r"/web/\1id_/", snap["url"])
            r = requests.get(raw, headers=UA, timeout=60)
            (out / f"wb_{name}_{ts}.html").write_text(f"<!-- {snap['timestamp']} HTTP {r.status_code} -->\n" + r.text)
            print(name, ts, snap["timestamp"], r.status_code, len(r.text))
        except Exception as e:
            print(name, ts, "error", e)
