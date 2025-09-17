import requests, os, math, time
import pandas as pd
from datetime import datetime, timedelta, timezone

# === 基本設定 ===
API_KEY = "CG-38G8j12guYuvo8ksrALeu9Du"  # 換成你的 CoinGecko demo key
VS = "usd"
START_DATE = datetime(2025, 2, 17)
END_DATE   = datetime(2025, 8, 16)

OUTPUT_DIR = r"C:\Users\Administrator\Desktop\論文\cryptodata"
os.makedirs(OUTPUT_DIR, exist_ok=True)
OUT_XLSX = os.path.join(OUTPUT_DIR, "eventday_meme_timeseries.xlsx")

# === 事件日 defi 幣 IDs ===
coin_ids = [
    "dogecoin", "shiba-inu", "pepe", "bonk", "pudgy-penguins",
    "pump-fun", "spx6900", "sky", "fartcoin", "floki",
    "dogwifcoin", "official-trump", "gala", "the-sandbox"
]

# === 基本工具 ===
def cg_get(url, params=None):
    p = params.copy() if params else {}
    p["x_cg_demo_api_key"] = API_KEY
    r = requests.get(url, params=p, timeout=30)
    r.raise_for_status()
    return r.json()

def fetch_range_daily(coin_id, vs, start_dt, end_dt):
    """抓某幣區間的日資料 (價格、市值、成交量)"""
    url = f"https://api.coingecko.com/api/v3/coins/{coin_id}/market_chart/range"
    start_unix = int(start_dt.replace(tzinfo=timezone.utc).timestamp())
    end_unix   = int(end_dt.replace(tzinfo=timezone.utc).timestamp())
    js = cg_get(url, {"vs_currency": vs, "from": start_unix, "to": end_unix})

    def arr_to_df(key, col):
        arr = js.get(key, [])
        if not arr: return pd.DataFrame(columns=["date", col])
        df = pd.DataFrame(arr, columns=["ts", col])
        df["date"] = pd.to_datetime(df["ts"], unit="ms", utc=True).dt.date
        return df.groupby("date", as_index=False).last()

    price = arr_to_df("prices", "price")
    mcap  = arr_to_df("market_caps", "market_cap")
    vol   = arr_to_df("total_volumes", "total_volume")
    out = price.merge(mcap, on="date", how="outer").merge(vol, on="date", how="outer")
    return out


def fetch_fear_greed():
    """抓 Alternative.me Fear & Greed 指數"""
    js = requests.get("https://api.alternative.me/fng/", params={"limit": 0}, timeout=30).json()
    data = js.get("data", [])
    fg = pd.DataFrame(data)
    fg["timestamp"] = pd.to_numeric(fg["timestamp"], errors="coerce")
    fg["date"] = pd.to_datetime(fg["timestamp"], unit="s", utc=True).dt.date
    fg["fear_greed"] = pd.to_numeric(fg["value"], errors="coerce")
    return fg[["date","fear_greed"]]

# === 主流程 ===
print("抓 Fear & Greed 指數 …")
fg_df = fetch_fear_greed()

rows = []
for i, cid in enumerate(coin_ids, 1):
    print(f"[{i}/{len(coin_ids)}] 抓 {cid} …")
    try:
        df = fetch_range_daily(cid, VS, START_DATE, END_DATE)
        df["log_market_cap"] = df["market_cap"].apply(lambda x: math.log(x) if x and x>0 else None)

        # merge Fear & Greed
        df = df.merge(fg_df, on="date", how="left")

        df.insert(0, "coin_id", cid)
        rows.append(df)
    except Exception as e:
        print(f"  [跳過] {cid}: {e}")
    time.sleep(5)  # 防止被限流

data = pd.concat(rows, ignore_index=True)

# 排序並輸出
cols = ["date","coin_id","price","market_cap","log_market_cap","total_volume","fear_greed"]
data = data[cols].sort_values(["date","coin_id"])

import openpyxl
with pd.ExcelWriter(OUT_XLSX, engine="openpyxl") as writer:
    data.to_excel(writer, index=False, sheet_name="data")

print("完成，已輸出：", OUT_XLSX)