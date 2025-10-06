import pandas as pd
import numpy as np
import statsmodels.api as sm
from datetime import datetime, timedelta

# === 基本設定 ===
INPUT_DIR = r"C:\Users\Administrator\Desktop\論文\crypto_data2"
INPUT_XLSX = f"{INPUT_DIR}/historical_data_2025-07-17_2.xlsx"
CMC20_XLSX = r"C:\Users\Administrator\Desktop\論文\cryptodata\CMC20.xlsx"
OUTPUT_DIR = r"C:\Users\Administrator\Desktop\論文\crypto_data_cmc"
EVENT_DATE = pd.to_datetime("2025-07-17")
CAR_WINDOW = (-3, 3)  # CAR 窗口：t=0 到 t=5
STABLECOIN_CLASS = "5. 穩定幣"

# === 讀取數據 ===
df = pd.read_excel(INPUT_XLSX)
df["date"] = pd.to_datetime(df["date"])
df["date"] = df["date"] - timedelta(days=1)  # 平移日期
df = df.sort_values(["coin_id", "date"])
df["rel_day"] = (df["date"] - EVENT_DATE).dt.days

# 排除穩定幣
non_stable_df = df[df["classification"] != STABLECOIN_CLASS]
if non_stable_df.empty:
    print(f"\n錯誤：排除 '{STABLECOIN_CLASS}' 後無數據！")
    exit()
print(f"\n排除穩定幣後包含 {non_stable_df['coin_id'].nunique()} 個幣種")

# === 讀取 CMC20 指數 ===
cmc = pd.read_excel(CMC20_XLSX)
cmc["date"] = pd.to_datetime(cmc["date"]) - pd.Timedelta(days=1)
cmc = cmc.rename(columns={"log_return": "mkt_return"})[["date", "mkt_return"]]
print("\nCMC20 市場指數 (最近 5 天):")
print(cmc.tail())

# === 事件研究法函數（計算 CAR） ===
def run_market_model(coin_df, cmc_df, estimation_start=-150, estimation_end=-11, event_start=-3, event_end=3):
    tmp = coin_df[["date", "log_return", "rel_day"]].rename(columns={"log_return": "ret"})
    tmp = tmp.merge(cmc_df, on="date", how="inner")

    # 估計期
    est_df = tmp[(tmp["rel_day"] >= estimation_start) & (tmp["rel_day"] <= estimation_end)].dropna()
    if est_df.empty or len(est_df) < 30:
        return None, None, "Insufficient estimation data"

    # 市場模型回歸
    X = sm.add_constant(est_df["mkt_return"])
    Y = est_df["ret"]
    try:
        model = sm.OLS(Y, X).fit(cov_type="HC0")
        alpha, beta = model.params["const"], model.params["mkt_return"]
    except Exception as e:
        return None, None, f"Regression failed: {str(e)}"

    # 事件期
    evt_df = tmp[(tmp["rel_day"] >= event_start) & (tmp["rel_day"] <= event_end)].copy()
    if evt_df.empty:
        return None, None, "Empty event data"

    # 計算 AR 和 CAR
    evt_df["expected_ret"] = alpha + beta * evt_df["mkt_return"]
    evt_df["AR"] = evt_df["ret"] - evt_df["expected_ret"]
    car = evt_df["AR"].sum()

    return evt_df, car, None

# === 主程式：計算 CAR 和控制變數 ===
car_data = []
excluded_coins = []

for coin_id in non_stable_df["coin_id"].unique():
    print(f"\n處理幣種: {coin_id}")
    coin_df = non_stable_df[non_stable_df["coin_id"] == coin_id]

    # 計算 CAR
    evt_df, car, reason = run_market_model(coin_df, cmc, event_start=CAR_WINDOW[0], event_end=CAR_WINDOW[1])
    if evt_df is None:
        coin_info = coin_df[["coin_id", "symbol", "name"]].iloc[0]
        excluded_coins.append({
            "coin_id": coin_info["coin_id"],
            "symbol": coin_info["symbol"],
            "name": coin_info["name"],
            "reason": reason
        })
        continue

    # 控制變數
    # 對數市值：t=0 或 t=-1 的 log_market_cap
    event_day_data = coin_df[coin_df["rel_day"] == 0]
    if event_day_data.empty:
        event_day_data = coin_df[coin_df["rel_day"] == -1]
    if not event_day_data.empty and "log_market_cap" in event_day_data.columns:
        log_market_cap = event_day_data["log_market_cap"].iloc[0]
    else:
        log_market_cap = np.nan

    # 動能：t=-21 到 t=-1 的累積 log_return
    momentum_period = coin_df[(coin_df["rel_day"] >= -24) & (coin_df["rel_day"] <= -4)]
    momentum = momentum_period["log_return"].sum() if not momentum_period.empty else np.nan

    car_data.append({
        "coin_id": coin_id,
        "symbol": coin_df["symbol"].iloc[0],
        "name": coin_df["name"].iloc[0],
        "CAR": car,
        "log_market_cap": log_market_cap,
        "momentum": momentum
    })

# 轉為 DataFrame
car_df = pd.DataFrame(car_data)
if car_df.empty:
    print("\n無有效 CAR 數據，無法進行回歸！")
    exit()

# 儲存 CAR 和控制變數
CAR_XLSX = f"{OUTPUT_DIR}/CAR_control_variables_2025-07-17_all_coins_no_stablecoin.xlsx"
car_df.to_excel(CAR_XLSX, index=False)
print(f"\nCAR 和控制變數已儲存至 {CAR_XLSX}")
print("\n=== CAR 和控制變數概覽 ===")
print(car_df[["coin_id", "symbol", "CAR", "log_market_cap", "momentum"]].to_string(index=False))

# === Winsorize 函數 ===
def winsorize_series(series, lower=0.01, upper=0.99):
    q_low = series.quantile(lower)
    q_high = series.quantile(upper)
    series = series.copy()
    series[series < q_low] = q_low
    series[series > q_high] = q_high
    return series

# 清理數據
reg_df = car_df.dropna(subset=["CAR", "log_market_cap", "momentum"]).copy()
#reg_df["CAR"] = winsorize_series(reg_df["CAR"], lower=0.01, upper=0.99)

# === 橫斷面回歸 ===
if len(reg_df) < 5:
    print("\n有效數據過少（少於 5 個幣種），無法進行回歸！")
else:
    X = reg_df[["log_market_cap", "momentum"]]
    X = sm.add_constant(X)
    y = reg_df["CAR"]

    model = sm.OLS(y, X).fit(cov_type="HC0")
    print("\n=== 橫斷面回歸結果 ===")
    print(model.summary())

    # 儲存回歸結果
    results = {
        "variable": ["const", "log_market_cap", "momentum"],
        "coefficient": model.params,
        "t_value": model.tvalues,
        "p_value": model.pvalues,
        "r_squared": [model.rsquared] + [np.nan] * 2,
        "r_squared_adj": [model.rsquared_adj] + [np.nan] * 2,
        "N": [int(model.nobs)] + [np.nan] * 2,
        "MSE": [model.mse_resid] + [np.nan] * 2
    }
    results_df = pd.DataFrame(results)
    REG_XLSX = f"{OUTPUT_DIR}/cross_sectional_regression_2025-07-17_all_coins_no_stablecoin_cut.xlsx"
    results_df.to_excel(REG_XLSX, index=False)
    print(f"\n回歸結果已儲存至 {REG_XLSX}")

# === 列出排除的幣種 ===
if excluded_coins:
    print("\n被排除的幣種:")
    df_excluded = pd.DataFrame(excluded_coins)
    print(df_excluded[["coin_id", "symbol", "name", "reason"]])
else:
    print("\n無幣種被排除。")
