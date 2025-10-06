import pandas as pd
import numpy as np
import statsmodels.api as sm
from scipy import stats
from datetime import datetime, timedelta

# === 基本設定 ===
INPUT_DIR = r"C:\Users\Administrator\Desktop\論文\crypto_data2"
INPUT_XLSX = f"{INPUT_DIR}/historical_data_2025-07-17_2.xlsx"
OUTPUT_DIR = INPUT_DIR
EVENT_DATE = pd.to_datetime("2025-07-17")

# === 讀取數據 ===
df = pd.read_excel(INPUT_XLSX)
df["date"] = pd.to_datetime(df["date"])
df["date"] = df["date"] - timedelta(days=1)  # 平移日期
df = df.sort_values(["coin_id", "date"])
df["rel_day"] = (df["date"] - EVENT_DATE).dt.days

# 排除穩定幣分類
df = df[df["classification"] != "5. 穩定幣"]
print(f"\n排除 '5. 穩定幣' 分類後，剩餘 {df['coin_id'].nunique()} 個幣種")

# 提取 BTC 作為市場指數
btc_df = df[df["coin_id"] == "bitcoin"][["date", "log_return"]].rename(columns={"log_return": "market_return"})
if btc_df.empty:
    print("\n錯誤：數據中未找到 BTC！")
    exit()

# === 市場模型 ===
def run_market_model(coin_df, btc_df, est_start=-150, est_end=-11, evt_start=-10, evt_end=10):
    tmp = coin_df[["date", "log_return", "rel_day"]].rename(columns={"log_return": "ret"})
    tmp = tmp.merge(btc_df, on="date", how="inner")

    est_df = tmp[(tmp["rel_day"] >= est_start) & (tmp["rel_day"] <= est_end)].dropna()
    if len(est_df) < 30:
        return None, "Insufficient estimation data"

    X = sm.add_constant(est_df["market_return"])
    Y = est_df["ret"]
    try:
        model = sm.OLS(Y, X).fit()
        alpha, beta = model.params["const"], model.params["market_return"]
    except Exception as e:
        return None, f"Regression failed: {str(e)}"

    evt_df = tmp[(tmp["rel_day"] >= evt_start) & (tmp["rel_day"] <= evt_end)].copy()
    if evt_df.empty:
        return None, "Empty event data"

    evt_df["expected_ret"] = alpha + beta * evt_df["market_return"]
    evt_df["AR"] = evt_df["ret"] - evt_df["expected_ret"]
    return evt_df, None

# === Winsorize AR (事件期截尾) ===
def winsorize_ar(ar_df, col="AR", lower=0.01, upper=0.99):
    q_low = ar_df[col].quantile(lower)
    q_high = ar_df[col].quantile(upper)
    ar_df = ar_df.copy()
    ar_df.loc[ar_df[col] < q_low, col] = q_low
    ar_df.loc[ar_df[col] > q_high, col] = q_high
    return ar_df

# === 計算 AAR / CAAR (t-test) ===
def compute_ttest(ar_df):
    aar_df = (
        ar_df[(ar_df["rel_day"] >= -10) & (ar_df["rel_day"] <= 10)]
        .groupby("rel_day")["AR"].mean()
        .reset_index()
        .rename(columns={"AR":"AAR"})
    )
    aar_df["CAAR"] = aar_df["AAR"].cumsum()

    results = []
    for day in aar_df["rel_day"]:
        ar_group = ar_df[ar_df["rel_day"]==day]["AR"]
        aar_t, aar_p = stats.ttest_1samp(ar_group, 0, nan_policy="omit")

        car_list = []
        for coin, g in ar_df.groupby("coin_id"):
            car_val = g[(g["rel_day"]>=-10)&(g["rel_day"]<=day)]["AR"].sum()
            car_list.append(car_val)
        caar_t, caar_p = stats.ttest_1samp(car_list, 0, nan_policy="omit")

        results.append({
            "rel_day": day,
            "AAR": aar_df.loc[aar_df["rel_day"]==day,"AAR"].values[0],
            "CAAR": aar_df.loc[aar_df["rel_day"]==day,"CAAR"].values[0],
            "AAR_t": aar_t, "AAR_p": aar_p,
            "CAAR_t": caar_t, "CAAR_p": caar_p
        })
    return pd.DataFrame(results)

# === 計算 AAR / CAAR (Robust SE) ===
def compute_robust(ar_df):
    aar_df = (
        ar_df[(ar_df["rel_day"] >= -10) & (ar_df["rel_day"] <= 10)]
        .groupby("rel_day")["AR"].mean()
        .reset_index()
        .rename(columns={"AR":"AAR"})
    )
    aar_df["CAAR"] = aar_df["AAR"].cumsum()

    results = []
    for day in aar_df["rel_day"]:
        # AAR robust
        ar_group = ar_df[ar_df["rel_day"]==day]["AR"].dropna()
        if len(ar_group)>1:
            X = np.ones(len(ar_group))
            model = sm.OLS(ar_group, X).fit(cov_type="HC1")
            aar_t, aar_p = model.tvalues.iloc[0], model.pvalues.iloc[0]
        else:
            aar_t, aar_p = np.nan, np.nan

        # CAAR robust
        car_list = []
        for coin, g in ar_df.groupby("coin_id"):
            car_val = g[(g["rel_day"]>=-10)&(g["rel_day"]<=day)]["AR"].sum()
            car_list.append(car_val)
        car_series = pd.Series(car_list).dropna()
        if len(car_series)>1:
            X = np.ones(len(car_series))
            model = sm.OLS(car_series, X).fit(cov_type="HC1")
            caar_t, caar_p = model.tvalues.iloc[0], model.pvalues.iloc[0]
        else:
            caar_t, caar_p = np.nan, np.nan

        results.append({
            "rel_day": day,
            "AAR": aar_df.loc[aar_df["rel_day"]==day,"AAR"].values[0],
            "CAAR": aar_df.loc[aar_df["rel_day"]==day,"CAAR"].values[0],
            "AAR_t": aar_t, "AAR_p": aar_p,
            "CAAR_t": caar_t, "CAAR_p": caar_p
        })
    return pd.DataFrame(results)

# === 主程式 ===
ar_all, excluded_coins = [], []
for coin_id in df["coin_id"].unique():
    if coin_id == "bitcoin": continue
    coin_df = df[df["coin_id"] == coin_id]
    evt_df, reason = run_market_model(coin_df, btc_df)
    if evt_df is None:
        excluded_coins.append({"coin_id": coin_id, "reason": reason})
        continue
    evt_df["coin_id"] = coin_id
    ar_all.append(evt_df)

if ar_all:
    ar_df = pd.concat(ar_all, ignore_index=True)

    # 1. Normal
    res_normal = compute_ttest(ar_df)
    res_normal.to_excel(f"{OUTPUT_DIR}/AAR_CAAR_normal.xlsx", index=False)

    # 2. Winsorize
    ar_df_winsor = winsorize_ar(ar_df)
    res_winsor = compute_ttest(ar_df_winsor)
    res_winsor.to_excel(f"{OUTPUT_DIR}/AAR_CAAR_winsor.xlsx", index=False)

    # 3. Robust
    res_robust = compute_robust(ar_df)
    res_robust.to_excel(f"{OUTPUT_DIR}/AAR_CAAR_robust.xlsx", index=False)

    print("\n三種方法結果已輸出：")
    print(" - AAR_CAAR_normal.xlsx")
    print(" - AAR_CAAR_winsor.xlsx")
    print(" - AAR_CAAR_robust.xlsx")
else:
    print("\n無有效 AR 數據。")

if excluded_coins:
    print("\n被排除的幣種：")
    print(pd.DataFrame(excluded_coins))
