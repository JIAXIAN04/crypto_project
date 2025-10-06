import pandas as pd
import numpy as np
import statsmodels.api as sm
from scipy import stats
from datetime import datetime, timedelta

# === 基本設定 ===
INPUT_DIR = r"C:\Users\Administrator\Desktop\論文\crypto_data2"
INPUT_XLSX = f"{INPUT_DIR}/historical_data_2025-07-17_2.xlsx"
CMC20_XLSX = r"C:\Users\Administrator\Desktop\論文\cryptodata\CMC20.xlsx"  # CMC20 指數檔案
OUTPUT_DIR = INPUT_DIR
AAR_CAAR_XLSX = f"{OUTPUT_DIR}/AAR_CAAR_results_2025-07-17_all_no_stablecoins_CMC20.xlsx"
CHECK_WINDOW_XLSX = f"{OUTPUT_DIR}/AAR_CAAR_check_window_2025-07-17_all_no_stablecoins_CMC20.xlsx"
EVENT_DATE = pd.to_datetime("2025-07-17")

# === 讀取加密貨幣數據 ===
df = pd.read_excel(INPUT_XLSX)
df["date"] = pd.to_datetime(df["date"])
df["date"] = df["date"] - timedelta(days=1)  # CoinGecko 收盤日對齊
df = df.sort_values(["coin_id", "date"])
df["rel_day"] = (df["date"] - EVENT_DATE).dt.days

# 排除穩定幣分類
df = df[df["classification"] != "5. 穩定幣"]
print(f"\n排除 '5. 穩定幣' 後，剩餘 {df['coin_id'].nunique()} 個幣種")

# === 讀取 CMC20 指數 ===
cmc = pd.read_excel(CMC20_XLSX)
cmc["date"] = pd.to_datetime(cmc["date"]) - pd.Timedelta(days=1)
cmc = cmc.rename(columns={"log_return": "mkt_return"})[["date", "mkt_return"]]
print("\nCMC20 市場指數 (最近 5 天):")
print(cmc.tail())

# === 事件研究法函數 ===
def run_market_model(coin_df, cmc_df, estimation_start=-150, estimation_end=-11, event_start=-10, event_end=10):
    tmp = coin_df[["date", "log_return", "rel_day"]].rename(columns={"log_return": "ret"})
    tmp = tmp.merge(cmc_df, on="date", how="inner")

    est_df = tmp[(tmp["rel_day"] >= estimation_start) & (tmp["rel_day"] <= estimation_end)].dropna()
    if est_df.empty or len(est_df) < 30:
        return None, "Insufficient estimation data"

    X = sm.add_constant(est_df["mkt_return"])
    Y = est_df["ret"]
    try:
        model = sm.OLS(Y, X).fit()
        alpha, beta = model.params["const"], model.params["mkt_return"]
    except Exception as e:
        return None, f"Regression failed: {str(e)}"

    evt_df = tmp[(tmp["rel_day"] >= event_start) & (tmp["rel_day"] <= event_end)].copy()
    if evt_df.empty:
        return None, "Empty event data"

    evt_df["expected_ret"] = alpha + beta * evt_df["mkt_return"]
    evt_df["AR"] = evt_df["ret"] - evt_df["expected_ret"]

    return evt_df, None

# === 計算 AAR / CAAR ===
def compute_aar_caar(ar_df):
    aar_df = (
        ar_df[(ar_df["rel_day"] >= -10) & (ar_df["rel_day"] <= 10)]
        .groupby("rel_day")["AR"]
        .mean()
        .reset_index()
        .rename(columns={"AR": "AAR"})
    )
    aar_df["CAAR"] = aar_df["AAR"].cumsum()

    results = []
    for day in aar_df["rel_day"]:
        ar_group = ar_df[ar_df["rel_day"] == day]["AR"]
        aar_t_stat, aar_p_value = stats.ttest_1samp(ar_group, popmean=0, nan_policy="omit")

        car_list = []
        for coin, group in ar_df.groupby("coin_id"):
            car_val = group[(group["rel_day"] >= -10) & (group["rel_day"] <= day)]["AR"].sum()
            car_list.append(car_val)
        caar_t_stat, caar_p_value = stats.ttest_1samp(car_list, popmean=0, nan_policy="omit")

        results.append({
            "rel_day": day,
            "AAR": aar_df.loc[aar_df["rel_day"] == day, "AAR"].values[0],
            "CAAR": aar_df.loc[aar_df["rel_day"] == day, "CAAR"].values[0],
            "AAR_t_stat": aar_t_stat,
            "AAR_p_value": aar_p_value,
            "CAAR_t_stat": caar_t_stat,
            "CAAR_p_value": caar_p_value
        })

    results_df = pd.DataFrame(results).sort_values("rel_day")
    check_df = results_df[(results_df["rel_day"] >= -3) & (results_df["rel_day"] <= 3)]
    return results_df, check_df

# === 主程式 ===
ar_all = []
excluded_coins = []

for coin_id in df["coin_id"].unique():
    print(f"Processing coin: {coin_id}")
    coin_df = df[df["coin_id"] == coin_id]
    evt_df, reason = run_market_model(coin_df, cmc)

    if evt_df is None:
        coin_info = df[df["coin_id"] == coin_id][["coin_id", "symbol", "name"]].iloc[0]
        excluded_coins.append({
            "coin_id": coin_info["coin_id"],
            "symbol": coin_info["symbol"],
            "name": coin_info["name"],
            "reason": reason
        })
        continue

    evt_df["coin_id"] = coin_id
    ar_all.append(evt_df)

if ar_all:
    ar_df = pd.concat(ar_all, ignore_index=True)
    results_df, check_df = compute_aar_caar(ar_df)

    results_df.to_excel(AAR_CAAR_XLSX, index=False)
    check_df.to_excel(CHECK_WINDOW_XLSX, index=False)

    print(f"\nAAR/CAAR 結果已儲存至 {AAR_CAAR_XLSX}")
    print(f"3 天檢查結果已儲存至 {CHECK_WINDOW_XLSX}")
    print("\n=== 事件日前後 3 天檢查 (AAR / CAAR / t-test) ===")
    print(check_df.to_string(index=False))
else:
    print("\n無有效 AR 數據。")

if excluded_coins:
    print("\n被排除的幣種：")
    df_excluded = pd.DataFrame(excluded_coins)
    print(df_excluded[["coin_id", "symbol", "name", "reason"]])
