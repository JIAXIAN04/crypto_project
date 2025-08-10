import requests
import os
import pandas as pd
import numpy as np
import statsmodels.api as sm
from datetime import datetime, timedelta

# === 設定區 ===
api_key = "CG-38G8j12guYuvo8ksrALeu9Du"
coins = ["bitcoin", "ethereum"]
vs_currency = "usd"
event_date = datetime(2025, 7, 17)
start_date = event_date - timedelta(days=150)
end_date = event_date + timedelta(days=30)
total_days = (end_date - start_date).days

# === 抓資料 ===
all_data = {}
for coin in coins:
    print(f"抓取 {coin} 資料中...")
    url = f"https://api.coingecko.com/api/v3/coins/{coin}/market_chart"
    params = {
        'vs_currency': vs_currency,
        'days': total_days,
        'x_cg_demo_api_key': api_key
    }
    r = requests.get(url, params=params)
    r.raise_for_status()
    prices = r.json()['prices']

    df = pd.DataFrame(prices, columns=["timestamp", "price"])
    df['datetime'] = pd.to_datetime(df['timestamp'], unit='ms', utc=True)
    df['date'] = df['datetime'].dt.date
    df = df[['date', 'price']].drop_duplicates(subset='date').reset_index(drop=True)

    # 篩選日期區間
    df = df[(df['date'] >= start_date.date()) & (df['date'] <= end_date.date())]
    df = df.reset_index(drop=True)  # 這行會讓 index 從 0 開始
    all_data[coin] = df

for coin, df in all_data.items():
    df['log_return'] = np.log(df['price'] / df['price'].shift(1))
    all_data[coin] = df  # 更新回原資料結構中

# === 顯示範例 ===
print("BTC 前幾天價格：")
print(all_data["bitcoin"].tail())

print("\nETH 前幾天價格：")
print(all_data["ethereum"].head())

# 建立輸出資料夾
output_dir = "crypto_prices"
os.makedirs(output_dir, exist_ok=True)

# 將每個幣種的價格資料輸出為 Excel
for coin, df in all_data.items():
    filename = os.path.join(output_dir, f"{coin}_price.csv")
    df.to_csv(filename, index=False)
    print(f"{coin} 資料已儲存到：{filename}")

# 抓出資料
btc_df = all_data["bitcoin"].copy()
eth_df = all_data["ethereum"].copy()

# 確保依日期排序，避免因排序錯誤造成錯位
btc_df = btc_df.sort_values("date")
eth_df = eth_df.sort_values("date")

# 設定欄位名稱好合併
btc_df = btc_df[['date', 'log_return']].rename(columns={'log_return': 'btc_return'})
eth_df = eth_df[['date', 'log_return']].rename(columns={'log_return': 'eth_return'})

# 合併資料（根據 date）
merged = pd.merge(eth_df, btc_df, on='date', how='inner')

# 先把 ±inf 轉成 NaN，再一次丟掉
merged = merged.replace([np.inf, -np.inf], np.nan)
merged = merged.dropna(subset=["eth_return", "btc_return"]).copy()

# 設定估計期範圍（2025/2/17 ～ 2025/7/7）
start_est = datetime(2025, 2, 17).date()
end_est = datetime(2025, 7, 7).date()

# 選出估計期的資料
estimation_data = merged[(merged['date'] >= start_est) & (merged['date'] <= end_est)]

# 設定 X、Y
X = estimation_data['btc_return']
Y = estimation_data['eth_return']
X = sm.add_constant(X)  # 加上截距項

# 跑 OLS 回歸
model = sm.OLS(Y, X).fit()

# 顯示結果
print(model.summary())

# === 0) 事件日與事件窗 ===
win_start = (event_date - timedelta(days=10)).date()   # 事件窗起：-10
win_end   = (event_date + timedelta(days=30)).date()   # 事件窗迄：+30
event_date = datetime(2025, 7, 17).date()
print(win_start)

# === 1) 取 α、β（來自你剛剛的回歸結果 model）===
alpha = model.params['const']
beta  = model.params['btc_return']
print(f"alpha={alpha:.6g}, beta={beta:.6g}")

# === 2) 準備完整期間的報酬資料（ETH、BTC）===
btc_full = all_data["bitcoin"][["date", "log_return"]].rename(columns={"log_return":"btc_return"}).copy()
eth_full = all_data["ethereum"][["date", "log_return"]].rename(columns={"log_return":"eth_return"}).copy()

print(btc_full.head())

# 依日期排序、內連接對齊
df = pd.merge(eth_full.sort_values("date"), btc_full.sort_values("date"), on="date", how="inner")

print(df.head())


# 只留事件窗（避免含估計期的 NaN 影響）
df = df[(df["date"] >= win_start)& (df["date"] <= win_end)].copy()

print(df.head())

# 清理極端值/遺漏
df.replace([np.inf, -np.inf], np.nan, inplace=True)
df.dropna(subset=["eth_return","btc_return"], inplace=True)

print(df.head())

# === 3) 期望報酬、異常報酬 AR、相對日 ===
df["exp_return"] = alpha + beta * df["btc_return"]           # E[R_it] = α + β * R_mt
df["AR"] = df["eth_return"] - df["exp_return"]               # AR = R_it - E[R_it]
df["rel_day"] = (df["date"] - event_date).apply(lambda x: x.days)           # 相對事件日（0 為 7/17）

print(df.head())

# 若你也要對 BTC 自己做 AR（通常不需要），可類推建立 exp_return_btc 等

# === 4) 建立 CAR 計算函式（可計任意視窗）===
def car_over_window(frame: pd.DataFrame, t1: int, t2: int) -> float:
    """在相對日介於 [t1, t2] 之間的 AR 進行加總，回傳 CAR。"""
    sel = frame[(frame["rel_day"] >= t1) & (frame["rel_day"] <= t2)]
    return float(sel["AR"].sum())

# 你可以自訂想看的事件窗清單（以下舉例）
windows = [(-1, 1), (-3, 3), (-5, 5), (0, 1), (0, 3), (0, 5), (-10, 10), (-10, 30)]

# 計算各視窗 CAR
car_results = []
for (a,b) in windows:
    car_results.append({"window": f"[{a},{b}]", "CAR": car_over_window(df, a, b)})

car_df = pd.DataFrame(car_results).sort_values("window")
print("\n=== CAR by window (ETH vs BTC Market Model) ===")
print(car_df)

# === 5) 如需輸出事件窗逐日的 AR/CAR 累加軌跡 ===
# 這裡示範累積到當日的 CAR 路徑（從最小 rel_day 開始累積）
df = df.sort_values("rel_day")
df["CAR_running"] = df["AR"].cumsum()

# 檢視前幾列
print("\n=== Event-window daily AR & CAR_running (head) ===")
print(df[["date","rel_day","eth_return","btc_return","exp_return","AR","CAR_running"]].head())

# 如需匯出：
df.to_excel("crypto_prices/ETH_event_AR_CAR.xlsx", index=False)
car_df.to_excel("crypto_prices/ETH_event_CAR_windows.xlsx", index=False)

from scipy import stats

# === 6) 每日 AR 單樣本 t 檢定（事件窗內）===
# 因為現在只有一個幣種，所以每日 AR 就是當天的值，t-test 沒有意義
# 但如果你以後有多個幣種，可以 groupby 每天做檢定
# 以下先模擬多幣種的檢定結構（目前只有單幣，p-value 無法反映統計意義）

daily_t_results = []
for day, group in df.groupby("rel_day"):
    # 單樣本 t 檢定：檢查該天 AR 是否顯著不等於 0
    t_stat, p_value = stats.ttest_1samp(group["AR"], popmean=0)
    daily_t_results.append({
        "rel_day": day,
        "date": group["date"].iloc[0],
        "mean_AR": group["AR"].mean(),
        "t_stat": t_stat,
        "p_value": p_value
    })

daily_t_df = pd.DataFrame(daily_t_results).sort_values("rel_day")

print("\n=== Daily AR t-test results ===")
print(daily_t_df)

# 輸出成 Excel
daily_t_df.to_excel("crypto_prices/ETH_event_AR_ttest.xlsx", index=False)