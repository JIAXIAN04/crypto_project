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

# 指定開始分類
START_CLASS = '2. Layer2/3 擴展應用層代幣'

# === 讀取數據 ===
df = pd.read_excel(INPUT_XLSX)
df["date"] = pd.to_datetime(df["date"])

# 日期平移：CoinGecko 的 7/18 00:00:00 視為 7/17 23:59:59 收盤價
df["date"] = df["date"] - timedelta(days=1)

# 確保按 coin_id 和 date 排序
df = df.sort_values(["coin_id", "date"])

# 加上相對天數
df["rel_day"] = (df["date"] - EVENT_DATE).dt.days

# 提取 BTC 作為市場指數（假設 BTC 在數據中，且不屬於任何分類或可單獨處理）
btc_df = df[df["coin_id"] == "bitcoin"][["date", "log_return"]].rename(columns={"log_return": "market_return"})
print("\nBTC 市場指數 (最近 5 天):")
print(btc_df.tail())

# === 顯示每個分類的幣種數量 ===
print("\n=== 各分類幣種數量 ===")
class_counts = df.groupby("classification")["coin_id"].nunique().reset_index()
class_counts.columns = ["classification", "coin_count"]
print(class_counts.to_string(index=False))

# === 事件研究法函數 ===
def run_market_model(coin_df, btc_df, estimation_start=-150, estimation_end=-11, event_start=-10, event_end=10):
    # 合併幣種數據與 BTC 市場報酬
    tmp = coin_df[["date", "log_return", "rel_day"]].rename(columns={"log_return": "ret"})
    tmp = tmp.merge(btc_df, on="date", how="inner")

    # 估計期 (-150 到 -11)
    est_df = tmp[(tmp["rel_day"] >= estimation_start) & (tmp["rel_day"] <= estimation_end)].dropna()
    if est_df.empty or len(est_df) < 30:  # 至少需要 30 天數據以確保回歸穩定
        return None, "Insufficient estimation data"

    # 市場模型回歸
    X = sm.add_constant(est_df["market_return"])
    Y = est_df["ret"]
    try:
        model = sm.OLS(Y, X).fit()
        alpha, beta = model.params["const"], model.params["market_return"]
    except Exception as e:
        return None, f"Regression failed: {str(e)}"

    # 事件期 (-10 到 +10)
    evt_df = tmp[(tmp["rel_day"] >= event_start) & (tmp["rel_day"] <= event_end)].copy()
    if evt_df.empty:
        return None, "Empty event data"

    # 計算預期報酬和異常報酬
    evt_df["expected_ret"] = alpha + beta * evt_df["market_return"]
    evt_df["AR"] = evt_df["ret"] - evt_df["expected_ret"]

    return evt_df, None

# === 計算 AAR 和 CAAR 函數 ===
def compute_aar_caar(ar_df, classification):
    if ar_df.empty:
        return None, None

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
        # AAR t 檢定
        ar_group = ar_df[ar_df["rel_day"] == day]["AR"]
        aar_t_stat, aar_p_value = stats.ttest_1samp(ar_group, popmean=0, nan_policy="omit")

        # CAR t 檢定
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

    # 事件日前後 3 天檢查
    check_window = (-3, 3)
    check_df = results_df[(results_df["rel_day"] >= check_window[0]) & (results_df["rel_day"] <= check_window[1])]

    return results_df, check_df

# === 主程式：按分類運行 ===
unique_classes = df["classification"].unique()
print(f"\n所有分類: {unique_classes}")

# 從指定分類開始
if START_CLASS not in unique_classes:
    print(f"\n錯誤: 分類 '{START_CLASS}' 不存在於數據中。")
else:
    # 只處理 START_CLASS
    classification = START_CLASS
    print(f"\n處理分類: {classification}")

    class_df = df[df["classification"] == classification]
    ar_all = []
    excluded_coins = []

    for coin_id in class_df["coin_id"].unique():
        if coin_id == "bitcoin":  # 跳過 BTC，如果它在這個分類中
            continue

        print(f"  Processing coin: {coin_id}")
        coin_df = class_df[class_df["coin_id"] == coin_id]
        evt_df, reason = run_market_model(coin_df, btc_df)

        if evt_df is None:
            coin_info = class_df[class_df["coin_id"] == coin_id][["coin_id", "symbol", "name"]].iloc[0]
            excluded_coins.append({
                "coin_id": coin_info["coin_id"],
                "symbol": coin_info["symbol"],
                "name": coin_info["name"],
                "reason": reason
            })
            continue

        evt_df["coin_id"] = coin_id
        ar_all.append(evt_df)

    # 合併 AR 並計算 AAR/CAAR
    if ar_all:
        ar_df = pd.concat(ar_all, ignore_index=True)
        results_df, check_df = compute_aar_caar(ar_df, classification)

        if results_df is not None:
            class_safe = classification.replace("/", "_").replace(" ", "_")  # 安全檔名
            AAR_CAAR_XLSX = f"{OUTPUT_DIR}/AAR_CAAR_results_2025-07-17_{class_safe}.xlsx"
            CHECK_WINDOW_XLSX = f"{OUTPUT_DIR}/AAR_CAAR_check_window_2025-07-17_{class_safe}.xlsx"

            results_df.to_excel(AAR_CAAR_XLSX, index=False)
            check_df.to_excel(CHECK_WINDOW_XLSX, index=False)

            print(f"\n{classification} 的 AAR 和 CAAR 結果已儲存至 {AAR_CAAR_XLSX}")
            print(f"{classification} 的 3 天檢查結果已儲存至 {CHECK_WINDOW_XLSX}")

            print("\n=== 事件日前後 3 天檢查 (AAR / CAAR / t-test) ===")
            print(check_df.to_string(index=False))
    else:
        print(f"\n{classification} 無有效 AR 數據。")

    # 列出排除的幣種
    if excluded_coins:
        print(f"\n{classification} 被排除的幣種:")
        df_excluded = pd.DataFrame(excluded_coins)
        print(df_excluded[["coin_id", "symbol", "name", "reason"]])
    else:
        print(f"\n{classification} 無幣種被排除。")