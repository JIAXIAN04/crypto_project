import requests, os, pandas as pd
from datetime import datetime, timezone
# === 設定 ===
CMC_API_KEY = "e34b7dd5-b2e3-42c7-864b-b9d84f491f20"
HEADERS = {"X-CMC_PRO_API_KEY": CMC_API_KEY}
OUTPUT_DIR = r"C:\Users\Administrator\Desktop\論文\cryptodata"
os.makedirs(OUTPUT_DIR, exist_ok=True)
OUT_FILE = os.path.join(OUTPUT_DIR, "cmc_top200_with_date_added.xlsx")

def cmc_get(url, params=None):
    r = requests.get(url, headers=HEADERS, params=params, timeout=30)
    r.raise_for_status()
    return r.json()

# 抓市值前200
def fetch_cmc_top200():
    url = "https://pro-api.coinmarketcap.com/v1/cryptocurrency/listings/latest"
    params = {"start": 1, "limit": 200, "convert": "USD", "sort": "market_cap"}
    js = cmc_get(url, params)
    data = js["data"]

    records = []
    for d in data:
        records.append({
            "cmc_id": d["id"],
            "name": d["name"],
            "symbol": d["symbol"],
            "date_added": d["date_added"],
            "last_updated": d["last_updated"],
        })
    return pd.DataFrame(records)

print("抓取 CMC 市值前200 …")
df = fetch_cmc_top200()
# 算資產年齡（天）
df["date_added"] = pd.to_datetime(df["date_added"], utc=True).dt.date
today = datetime.now(timezone.utc).date()
df["asset_age_days"] = (today - df["date_added"]).apply(lambda x: x.days)
# 存檔
df.to_excel(OUT_FILE, index=False, engine="openpyxl")
print("完成，已輸出：", OUT_FILE)

