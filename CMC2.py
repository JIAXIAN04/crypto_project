import requests, os, math, time
import pandas as pd
from datetime import datetime, timedelta, timezone

# === 設定 ===
CMC_API_KEY = "你的CMC Pro API Key"
HEADERS = {"X-CMC_PRO_API_KEY": CMC_API_KEY}
OUTPUT_DIR = r"C:\Users\Administrator\Desktop\cryptodata"
os.makedirs(OUTPUT_DIR, exist_ok=True)
OUT_XLSX = os.path.join(OUTPUT_DIR, "cmc_top100_timeseries.xlsx")

EVENT_DATE = datetime(2025, 7, 17, tzinfo=timezone.utc)
START_DATE = EVENT_DATE - timedelta(days=150)
END_DATE   = EVENT_DATE + timedelta(days=30)

# === 工具 ===
def cmc_get(url, params=None):
    r = requests.get(url, headers=HEADERS, params=params, timeout=30)
    r.raise_for_status()
    return r.json()

def fetch_top100(event_date):
    """抓事件日市值前100幣"""
    url = "https://pro-api.coinmarketcap.com/v1/cryptocurrency/listings/historical"
    params = {
        "date": event_date.strftime("%Y-%m-%d"),
        "limit": 100,
        "sort": "market_cap",
        "convert": "USD"
    }
    js = cmc_get(url, params)
    data = js["data"]
    records = []
    for d in data:
        records.append({
            "id": d["id"],
            "name": d["name"],
            "symbol": d["symbol"],
            "date_added": d["date_added"]
        })
    return pd.DataFrame(records)

def fetch_historical_quotes(coin_id, start, end):
    """抓單一幣在指定區間的每日行情"""
    url = f"https://pro-api.coinmarketcap.com/v1/cryptocurrency/ohlcv/historical"
    params = {
        "id": coin_id,
        "time_start": start.strftime("%Y-%m-%d"),
        "time_end": end.strftime("%Y-%m-%d"),
        "interval": "daily",
        "convert": "USD"
    }
    js = cmc_get(url, params)
    rows = []
    for d in js["data"]["quotes"]:
        rows.append({
            "date": pd.to_datetime(d["time_open"]).date(),
            "price": d["quote"]["USD"]["close"],
            "market_cap": d["quote"]["USD"]["market_cap"],
            "log_market_cap": math.log(d["quote"]["USD"]["market_cap"]) if d["quote"]["USD"]["market_cap"] else None,
            "volume": d["quote"]["USD"]["volume"]
        })
    return pd.DataFrame(rows)

# === 主流程 ===
print("抓事件日 Top100 幣 …")
sample_df = fetch_top100(EVENT_DATE)

rows = []
today = datetime.now(timezone.utc).date()
for i, row in sample_df.iterrows():
    cid, name, sym, date_added = row["id"], row["name"], row["symbol"], row["date_added"]
    print(f"[{i+1}/100] 抓 {name} ({sym}) …")
    try:
        df = fetch_historical_quotes(cid, START_DATE, END_DATE)
        # 算資產年齡
        first_date = pd.to_datetime(date_added).date()
        df["asset_age_days"] = (pd.to_datetime(df["date"]) - pd.Timestamp(first_date)).dt.days
        df.insert(0, "name", name)
        df.insert(1, "symbol", sym)
        rows.append(df)
    except Exception as e:
        print(f"  [跳過] {name}: {e}")
    time.sleep(1)  # 避免限流

# 合併 & 輸出
data = pd.concat(rows, ignore_index=True)
cols = ["date","name","symbol","price","market_cap","log_market_cap","volume","asset_age_days"]
data = data[cols].sort_values(["date","symbol"])

with pd.ExcelWriter(OUT_XLSX, engine="openpyxl") as writer:
    data.to_excel(writer, index=False, sheet_name="data")
    sample_df.to_excel(writer, index=False, sheet_name="top100_ids")

print("完成，已輸出：", OUT_XLSX)
