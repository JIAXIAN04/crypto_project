import requests, pandas as pd, os

OUTPUT_DIR = r"C:\Users\Administrator\Desktop\論文\cryptodata"
os.makedirs(OUTPUT_DIR, exist_ok=True)
OUT_FILE = os.path.join(OUTPUT_DIR, "cg_top200_eventday.xlsx")

def cg_get(url, params=None):
    r = requests.get(url, params=params, timeout=30)
    r.raise_for_status()
    return r.json()

def fetch_eventday_top200(event_date="17-07-2025"):
    url = "https://api.coingecko.com/api/v3/coins/markets"
    params = {
        "vs_currency": "usd",
        "order": "market_cap_desc",
        "per_page": 200,
        "page": 1,
        "date": event_date  # 格式 dd-mm-yyyy
    }
    js = cg_get(url, params)
    records = []
    for d in js:
        records.append({
            "id": d["id"],
            "name": d["name"],
            "symbol": d["symbol"],
            "market_cap": d["market_cap"],
            "price": d["current_price"]
        })
    return pd.DataFrame(records)

df = fetch_eventday_top200()
df.to_excel(OUT_FILE, index=False, engine="openpyxl")
print("完成，已輸出事件日 Top200 名單：", OUT_FILE)