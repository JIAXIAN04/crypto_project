import pandas as pd
import numpy as np
import statsmodels.api as sm
from scipy import stats
from datetime import datetime, timedelta

# === 基本設定 ===
INPUT_DIR = r"C:\Users\Administrator\Desktop\論文\crypto_data2"
INPUT_XLSX = f"{INPUT_DIR}/historical_data_2025-07-17_2.xlsx"
OUTPUT_DIR = INPUT_DIR
AAR_CAAR_XLSX = f"{OUTPUT_DIR}/AAR_CAAR_results_2025-07-17_bootstrap.xlsx"
EVENT_DATE = pd.to_datetime("2025-07-17")

# === 讀取數據 ===
df = pd.read_excel(INPUT_XLSX)
df["date"] = pd.to_datetime(df["date"])
df["date"] = df["date"] - timedelta(days=1)  # 平移日期
df = df.sort_values(["coin_id", "date"])
df["rel_day"] = (df["date"] - EVENT_DATE).dt.days

# 提取 BTC 作為市場指數
btc_df = df[df["coin_id"] == "bitcoin"][["date", "log_return"]].rename(columns={"log_return": "market_return"})

# === 事件研究法 ===
def run_market_model(coin_df, btc_df, estimation_start=-150, estimation_end=-11, event_start=-10, event_end=10):
    tmp = coin_df[["date", "log_return", "rel_day"]].rename(columns={"log_return": "ret"})
    tmp = tmp.merge(btc_df, on="date", how="inner")

    est_df = tmp[(tmp["rel_day"] >= estimation_start) & (tmp["rel_day"] <= estimation_end)].dropna()
    if est_df.empty or len(est_df) < 30:
        return None, "Insufficient estimation data"

    X = sm.add_constant(est_df["market_return"])
    Y = est_df["ret"]
    try:
        model = sm.OLS(Y, X).fit()
        alpha, beta = model.params["const"], model.params["market_return"]
    except Exception as e:
        return None, f"Regression failed: {str(e)}"

    evt_df = tmp[(tmp["rel_day"] >= event_start) & (tmp["rel_day"] <= event_end)].copy()
    if evt_df.empty:
        return None, "Empty event data"

    evt_df["expected_ret"] = alpha + beta * evt_df["market_return"]
    evt_df["AR"] = evt_df["ret"] - evt_df["expected_ret"]

    return evt_df, None

# === Bootstrap 檢定函數 ===
def bootstrap_test(sample, B=1000):
    sample = np.array(sample.dropna())
    if len(sample) == 0:
        return np.nan, np.nan
    obs_mean = np.mean(sample)
    boot_means = []
    for _ in range(B):
        boot_sample = np.random.choice(sample, size=len(sample), replace=True)
        boot_means.append(np.mean(boot_sample))
    boot_means = np.array(boot_means)
    p_val = np.mean(np.abs(boot_means) >= np.abs(obs_mean))
    return obs_mean, p_val

# === 主程式 ===
ar_all, excluded_coins = [], []
for coin_id in df["coin_id"].unique():
    if coin_id == "bitcoin":
        continue
    coin_df = df[df["coin_id"] == coin_id]
    evt_df, reason = run_market_model(coin_df, btc_df)
    if evt_df is None:
        coin_info = df[df["coin_id"] == coin_id][["coin_id", "symbol", "name"]].iloc[0]
        excluded_coins.append({"coin_id": coin_info["coin_id"], "symbol": coin_info["symbol"], "name": coin_info["name"], "reason": reason})
        continue
    evt_df["coin_id"] = coin_id
    ar_all.append(evt_df)

if ar_all:
    ar_df = pd.concat(ar_all, ignore_index=True)

    # === 計算 AAR 和 CAAR ===
    aar_df = (
        ar_df[(ar_df["rel_day"] >= -10) & (ar_df["rel_day"] <= 10)]
        .groupby("rel_day")["AR"].mean()
        .reset_index()
        .rename(columns={"AR": "AAR"})
    )
    aar_df["CAAR"] = aar_df["AAR"].cumsum()

    results = []
    B = 1000  # bootstrap 次數

    for day in aar_df["rel_day"]:
        # --- AAR ---
        ar_group = ar_df[ar_df["rel_day"] == day]["AR"]

        # (1) 傳統 t-test
        aar_t_stat, aar_p_value = stats.ttest_1samp(ar_group, popmean=0, nan_policy="omit")

        # (2) Robust SE (回歸 mean=0, robust covariance)
        if len(ar_group.dropna()) > 1:
            X = np.ones(len(ar_group))
            model = sm.OLS(ar_group, X).fit(cov_type="HC1")
            aar_robust_t = model.tvalues.iloc[0]
            aar_robust_p = model.pvalues.iloc[0]
        else:
            aar_robust_t, aar_robust_p = np.nan, np.nan

        # (3) Bootstrap
        aar_obs, aar_boot_p = bootstrap_test(ar_group, B=B)

        # --- CAAR ---
        car_list = []
        for coin, group in ar_df.groupby("coin_id"):
            car_val = group[(group["rel_day"] >= -10) & (group["rel_day"] <= day)]["AR"].sum()
            car_list.append(car_val)
        car_series = pd.Series(car_list)

        # (1) 傳統 t-test
        caar_t_stat, caar_p_value = stats.ttest_1samp(car_series, popmean=0, nan_policy="omit")

        # (2) Robust SE
        if len(car_series.dropna()) > 1:
            X = np.ones(len(car_series))
            model = sm.OLS(car_series, X).fit(cov_type="HC1")
            caar_robust_t = model.tvalues.iloc[0]
            caar_robust_p = model.pvalues.iloc[0]
        else:
            caar_robust_t, caar_robust_p = np.nan, np.nan

        # (3) Bootstrap
        caar_obs, caar_boot_p = bootstrap_test(car_series, B=B)

        results.append({
            "rel_day": day,
            "AAR": aar_df.loc[aar_df["rel_day"] == day, "AAR"].values[0],
            "CAAR": aar_df.loc[aar_df["rel_day"] == day, "CAAR"].values[0],
            # AAR 結果
            "AAR_t_stat": aar_t_stat, "AAR_p_value": aar_p_value,
            "AAR_robust_t": aar_robust_t, "AAR_robust_p": aar_robust_p,
            "AAR_bootstrap": aar_obs, "AAR_boot_p": aar_boot_p,
            # CAAR 結果
            "CAAR_t_stat": caar_t_stat, "CAAR_p_value": caar_p_value,
            "CAAR_robust_t": caar_robust_t, "CAAR_robust_p": caar_robust_p,
            "CAAR_bootstrap": caar_obs, "CAAR_boot_p": caar_boot_p
        })

    results_df = pd.DataFrame(results).sort_values("rel_day")
    results_df.to_excel(AAR_CAAR_XLSX, index=False)
    print(f"\n三種檢定的 AAR / CAAR 結果已儲存至 {AAR_CAAR_XLSX}")

else:
    print("\nNo valid AR data generated.")
