import requests
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from scipy import stats
import statsmodels.api as sm
import os

# === 0) 設定 ===
api_key = "CG-38G8j12guYuvo8ksrALeu9Du"
coins = ["bitcoin", "ethereum", "ripple", "avalanche-2"]  # 幣種清單 (id來自CoinGecko)
vs_currency = "usd"

event_date = datetime(2025, 7, 17)
start_date = event_date - timedelta(days=150)
end_date = event_date + timedelta(days=30)
total_days = (end_date - start_date).days

output_dir = "crypto_prices"
os.makedirs(output_dir, exist_ok=True)

# === 1) 抓資料 + 計算 log return ===
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
    df['date'] = pd.to_datetime(df['timestamp'], unit='ms', utc=True).dt.date
    df = df[['date', 'price']].drop_duplicates(subset='date').reset_index(drop=True)
    df['log_return'] = np.log(df['price'] / df['price'].shift(1))
    all_data[coin] = df

print("資料抓取完成。")

# === 2) 準備市場指數 (BTC) ===
market_df = all_data["bitcoin"].rename(columns={"log_return": "mkt_return"})[["date", "mkt_return"]]

# === 3) 估計期範圍 ===
start_est = (event_date - timedelta(days=150)).date()
end_est = (event_date - timedelta(days=10)).date()

# === 4) 儲存所有幣的 AR 資料 ===
ar_all = []

for coin in coins:
    if coin == "bitcoin":
        continue  # BTC 當市場指數，不計算 AR/CAR

    df = all_data[coin][["date", "log_return"]].rename(columns={"log_return": "ret"})
    df = pd.merge(df, market_df, on="date", how="inner")

    # === 4.1) 估計期回歸 (Market Model) ===
    est_df = df[(df["date"] >= start_est) & (df["date"] <= end_est)].dropna()
    X = sm.add_constant(est_df["mkt_return"])
    Y = est_df["ret"]
    model = sm.OLS(Y, X).fit()

    alpha = model.params['const']
    beta = model.params['mkt_return']

    # === 4.2) 事件窗資料 ===
    win_start = (event_date - timedelta(days=10)).date()
    win_end = (event_date + timedelta(days=30)).date()
    win_df = df[(df["date"] >= win_start) & (df["date"] <= win_end)].copy()

    win_df["exp_return"] = alpha + beta * win_df["mkt_return"]
    win_df["AR"] = win_df["ret"] - win_df["exp_return"]
    win_df["rel_day"] = (pd.to_datetime(win_df["date"]) - pd.Timestamp(event_date)).dt.days
    win_df["coin"] = coin

    ar_all.append(win_df)

# 合併所有幣的 AR
ar_df = pd.concat(ar_all, ignore_index=True)

# === 5) 每日截面 AR t 檢定 ===
daily_t_results = []
for day, group in ar_df.groupby("rel_day"):
    t_stat, p_value = stats.ttest_1samp(group["AR"], popmean=0, nan_policy='omit')
    daily_t_results.append({
        "rel_day": day,
        "mean_AR": group["AR"].mean(),
        "t_stat": t_stat,
        "p_value": p_value
    })

daily_t_df = pd.DataFrame(daily_t_results).sort_values("rel_day")
daily_t_df.to_excel(os.path.join(output_dir, "Daily_AR_ttest.xlsx"), index=False)

# === 6) 多事件窗 CAR t 檢定 ===
windows = [(-1, 1), (-3, 3), (-5, 5), (0, 1), (0, 3), (0, 5), (-10, 10), (-10, 30)]
car_results = []

for (a, b) in windows:
    car_list = []
    for coin, group in ar_df.groupby("coin"):
        car_val = group[(group["rel_day"] >= a) & (group["rel_day"] <= b)]["AR"].sum()
        car_list.append(car_val)

    t_stat, p_value = stats.ttest_1samp(car_list, popmean=0, nan_policy='omit')
    car_results.append({
        "window": f"[{a},{b}]",
        "mean_CAR": np.mean(car_list),
        "t_stat": t_stat,
        "p_value": p_value
    })

car_df = pd.DataFrame(car_results)
car_df.to_excel(os.path.join(output_dir, "CAR_ttest.xlsx"), index=False)

print("分析完成，結果已輸出到資料夾：", output_dir)
