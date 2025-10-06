import requests
import time
import pandas as pd
import numpy as np
import os
from datetime import datetime, timedelta, timezone

# === 基本設定 ===
API_KEY = "CG-38G8j12guYuvo8ksrALeu9Du"  # 你的 Pro API key
VS = "usd"
EVENT_DATE = datetime(2025, 7, 17, tzinfo=timezone.utc)  # 事件日

# 日期範圍：-150 天到 +10 天 (總 161 天，包括事件日)
START_DATE = EVENT_DATE - timedelta(days=150)  # 2025-02-17
END_DATE = EVENT_DATE + timedelta(days=10) + timedelta(days=1)  # 2025-07-28

# 輸入/輸出路徑
INPUT_DIR = r"C:\Users\Administrator\Desktop\論文\crypto_data2"
INPUT_XLSX = os.path.join(INPUT_DIR, "coins_list_event_day_2025-07-17_top300.xlsx")
OUTPUT_DIR = INPUT_DIR
os.makedirs(OUTPUT_DIR, exist_ok=True)
OUT_HIST_XLSX = os.path.join(OUTPUT_DIR, "historical_data_2025-07-17_2.xlsx")

# === 基本工具 ===
def cg_get(url, params=None):
    p = params.copy() if params else {}
    p["x_cg_demo_api_key"] = API_KEY
    try:
        r = requests.get(url, params=p, timeout=30)
        r.raise_for_status()
    except requests.exceptions.RequestException as e:
        print(f"API Error: {e}, URL: {url}, Params: {p}")
        return {}  # 返回空字典，跳過該幣
    return r.json()

# 獲取單一幣的歷史數據
def get_coin_historical_data(coin_id):
    start_unix = int(START_DATE.timestamp())
    end_unix = int(END_DATE.timestamp())
    url = f"https://api.coingecko.com/api/v3/coins/{coin_id}/market_chart/range"
    params = {
        "vs_currency": VS,
        "from": start_unix,
        "to": end_unix,
        "precision": "full"  # 高精度
    }
    data = cg_get(url, params)
    if not data:
        return None, "No data"

    # 提取數據
    prices = data.get("prices", [])
    market_caps = data.get("market_caps", [])
    volumes = data.get("total_volumes", [])

    if len(prices) == 0:
        return None, "No price data"

    # 轉換為 DataFrame
    df = pd.DataFrame({
        "timestamp_ms": [p[0] for p in prices],
        "price": [p[1] for p in prices],
        "market_cap": [m[1] for m in market_caps],
        "volume": [v[1] for v in volumes]
    })

    # 轉換 timestamp 為日期 (UTC)
    df["date"] = pd.to_datetime(df["timestamp_ms"], unit="ms", utc=True).dt.date

    # 檢查數據起始日期（正負 3 天範圍）
    earliest_date = df["date"].min()
    start_date_min = (START_DATE - timedelta(days=3)).date()  # 2025-02-14
    start_date_max = (START_DATE + timedelta(days=3)).date()  # 2025-02-20
    if earliest_date < start_date_min or earliest_date > start_date_max:
        return None, f"Data starts outside allowed range: {earliest_date}"

    # 檢查缺失天數
    expected_dates = pd.date_range(start=START_DATE, end=END_DATE - timedelta(days=1), freq='D').date
    missing_dates = len(expected_dates) - len(df["date"].unique())
    if missing_dates > 2:
        return None, f"Missing {missing_dates} days of data"

    # 檢查市值和成交量缺失
    missing_market_cap = len(df[df["market_cap"].isna() | (df["market_cap"] == 0)])
    missing_volume = len(df[df["volume"].isna() | (df["volume"] == 0)])
    if missing_market_cap > 5:
        return None, f"Missing {missing_market_cap} market cap entries"
    if missing_volume > 5:
        return None, f"Missing {missing_volume} volume entries"

    # 計算對數報酬
    df["log_return"] = np.log(df["price"] / df["price"].shift(1))

    # 計算對數成交量
    df["log_volume"] = np.where(df["volume"] > 0, np.log(df["volume"]), np.nan)

    # 計算對數市值
    df["log_market_cap"] = np.where(df["market_cap"] > 0, np.log(df["market_cap"]), np.nan)

    # 刪除 timestamp_ms
    df = df.drop(columns=["timestamp_ms"])

    return df, None

# 主程式
# 步驟1: 讀取篩選後的 Excel
df_coins = pd.read_excel(INPUT_XLSX)
print(f"Loaded {len(df_coins)} coins from {INPUT_XLSX}")

# 結果列表和排除列表
results = []
excluded_coins = []

for i, row in df_coins.iterrows():
    coin_id = row["ID"]
    symbol = row["Symbol"]
    name = row["Name"]
    categories = row["Categories"]
    classification = row["Classification"]

    print(f"Processing coin {i + 1}/{len(df_coins)}: {symbol} ({coin_id})")

    df_hist, reason = get_coin_historical_data(coin_id)
    if df_hist is None:
        excluded_coins.append({"coin_id": coin_id, "symbol": symbol, "name": name, "reason": reason})
        continue

    # 添加幣資訊
    df_hist["coin_id"] = coin_id
    df_hist["symbol"] = symbol
    df_hist["name"] = name
    df_hist["categories"] = categories
    df_hist["classification"] = classification

    # 重新排列欄位
    df_hist = df_hist[
        ["coin_id", "symbol", "name", "categories", "date", "price", "log_return", "volume", "log_volume", "market_cap",
         "log_market_cap", "classification"]]

    results.append(df_hist)

    time.sleep(10)  # 間隔 10 秒，避免 API 限制

# 列出被排除的幣種
if excluded_coins:
    print("\n被排除的幣種:")
    df_excluded = pd.DataFrame(excluded_coins)
    print(df_excluded[["coin_id", "symbol", "name", "reason"]])
else:
    print("\n無幣種被排除。")

# 合併所有 DataFrame
if results:
    df_all = pd.concat(results, ignore_index=True)
    df_all.to_excel(OUT_HIST_XLSX, index=False)
    print(f"\n輸出到 {OUT_HIST_XLSX}")
    print(f"總資料行數: {len(df_all)} (約 {len(df_all) / len(results)} 天/幣)")

    # 統計最終分類分佈
    print("\n最終分類分佈 (每個分類的幣數):")
    final_classification_counts = df_all[["coin_id", "classification"]].drop_duplicates()["classification"].value_counts()
    print(final_classification_counts)
else:
    print("\nNo data fetched.")