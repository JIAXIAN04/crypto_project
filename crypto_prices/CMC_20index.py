import requests
import time
import pandas as pd
import numpy as np
from datetime import datetime, timedelta

# === 1. 參數設定 ===
API_KEY = "e34b7dd5-b2e3-42c7-864b-b9d84f491f20"

event_date = datetime(2025, 7, 17)
start_date = event_date - timedelta(days=150)
end_date   = event_date + timedelta(days=30)

url = "https://pro-api.coinmarketcap.com/v3/index/cmc20-historical"
headers = {
    "Accepts": "application/json",
    "X-CMC_PRO_API_KEY": API_KEY,
}

def fetch_chunk(s_date, e_date):
    params = {
        "time_start": s_date.strftime("%Y-%m-%dT%H:%M:%SZ"),
        "time_end": e_date.strftime("%Y-%m-%dT%H:%M:%SZ"),
        "interval": "daily",
    }
    resp = requests.get(url, headers=headers, params=params)
    data = resp.json()
    if "data" not in data:
        print("API 回傳錯誤：", data)
        return []
    return data["data"]

# === 2. 用迴圈分段抓取 ===
all_records = []
chunk_size = 4  # API 限制一次最多回 10 筆
current = start_date
counter = 0

while current < end_date:
    # 計算這次的結束時間
    chunk_end = min(current + timedelta(days=chunk_size), end_date)

    # 呼叫 API 抓資料
    quotes = fetch_chunk(current, chunk_end)
    for q in quotes:
        date = q["update_time"]  # e.g. "2025-07-01"
        price = q["value"]
        all_records.append({"date": date, "price": price})

    # 更新迴圈時間，避免重複
    current = chunk_end + timedelta(days=1)

    # 控制呼叫次數，避免 rate limit
    counter += 1
    if counter % 4 == 0:
        print("⏸ 已呼叫 10 次 API，暫停 60 秒避免超額...")
        time.sleep(120)

# === 3. 建 DataFrame ===
df = pd.DataFrame(all_records)
df["date"] = pd.to_datetime(df["date"]).dt.tz_localize(None)

# === 4. 算報酬 ===
df = df.sort_values("date").reset_index(drop=True)
df["return"] = df["price"].pct_change()
df["log_return"] = np.log(df["price"] / df["price"].shift(1))

# === 5. 輸出 Excel ===
output_path = r"C:\Users\Administrator\Desktop\論文\cryptodata\CMC20.xlsx"
df.to_excel(output_path, index=False)

print(f"完成！共抓到 {len(df)} 筆資料，輸出到 {output_path}")
