import requests, time, os
import pandas as pd
from datetime import datetime

 # === 設定 ===
API_KEY = "import requests, time, os
import pandas as pd
from datetime import datetime
EVENT_DATE = datetime(2025, 7, 17)   # 事件日
MAIN_EXCHANGES = {"Binance", "Coinbase Exchange", "Kraken", "OKX"}  # 主流交易所清單
OUTPUT_DIR = r"C:\Users\Administrator\Desktop\論文\cryptodata"
os.makedirs(OUTPUT_DIR, exist_ok=True)

# === 基本工具 ===
def cg_get(url, params=None):
    p = params.copy() if params else {}
    p["x_cg_demo_api_key"] = API_KEY
    r = requests.get(url, params=p, timeout=30)
    r.raise_for_status()
    return r.json()

def has_major_exchange(coin_id):
    """檢查該幣是否有在主流交易所掛牌"""
    url = f"https://api.coingecko.com/api/v3/coins/{coin_id}/tickers"
    try:
        js = cg_get(url)
        markets = {t["market"]["name"] for t in js.get("tickers", [])}
        return bool(markets & MAIN_EXCHANGES)
    except Exception:
        return False

def fetch_eventday_mcap(coin_id, event_dt):
    """抓事件日市值"""
    url = f"https://api.coingecko.com/api/v3/coins/{coin_id}/history"
    ddmmyyyy = event_dt.strftime("%d-%m-%Y")
    try:
        js = cg_get(url, {"date": ddmmyyyy, "localization": "false"})
        md = js.get("market_data") or {}
        return (md.get("market_cap") or {}).get("usd")
    except Exception:
        return None

# === Step 1: 抓市值排名前 200 的幣 ===
print("抓市值排名前 200 的幣…")
markets = cg_get("https://api.coingecko.com/api/v3/coins/markets", {
    "vs_currency": "usd",
    "order": "market_cap_desc",
    "per_page": 200,
    "page": 1
})

candidates = [m["id"] for m in markets]
print(f"候選數量: {len(candidates)}")

# === Step 2: 篩選主流交易所 & 抓事件日市值 ===
records = []
for i, cid in enumerate(candidates, 1):
    try:
        if has_major_exchange(cid):
            mcap = fetch_eventday_mcap(cid, EVENT_DATE)
            if mcap:
                records.append({
                    "coin_id": cid,
                    "symbol": markets[i-1]["symbol"],
                    "name": markets[i-1]["name"],
                    "eventday_market_cap_usd": mcap
                })
    except Exception as e:
        print(f"[跳過] {cid} {e}")
    time.sleep(0.7)  # Demo key 防限流
    if i % 20 == 0:
        print(f"已處理 {i}/{len(candidates)} …")

# === Step 3: 輸出結果 ===
df = pd.DataFrame(records).sort_values("eventday_market_cap_usd", ascending=False)
out_path = os.path.join(OUTPUT_DIR, "eventday_major_exchange_top200.csv")
df.to_csv(out_path, index=False, encoding="utf-8-sig")
print("完成，已輸出：", out_path)
