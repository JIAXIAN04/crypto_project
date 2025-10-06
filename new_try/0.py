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

# 指定保留的穩定幣
SELECTED_STABLECOINS = ["dai", "tether", "usd-coin", "ethena-usde", "usds"]

# === 讀取數據 ===
df = pd.read_excel(INPUT_XLSX)
df["date"] = pd.to_datetime(df["date"])

# 日期平移：CoinGecko 的 7/18 00:00:00 視為 7/17 23:59:59 收盤價
df["date"] = df["date"] - timedelta(days=1)

# 確保按 coin_id 和 date 排序
df = df.sort_values(["coin_id", "date"])

# 加上相對天數
df["rel_day"] = (df["date"] - EVENT_DATE).dt.days

# === 步驟1: 列出所有穩定幣 ===
stablecoin_class = "5. 穩定幣"
stablecoins_df = df[df["classification"] == stablecoin_class][["coin_id", "symbol", "name"]].drop_duplicates()

print("\n=== 數據中所有穩定幣（分類: '5. 穩定幣'） ===")
if not stablecoins_df.empty:
    print(stablecoins_df.to_string(index=False))
    print(f"\n總穩定幣數量: {len(stablecoins_df)}")
else:
    print("無穩定幣數據。")

# 檢查指定穩定幣是否存在
print("\n=== 檢查指定穩定幣是否存在 ===")
available_stablecoins = []
missing_stablecoins = []
for coin_id in SELECTED_STABLECOINS:
    if coin_id in stablecoins_df["coin_id"].values:
        coin_info = stablecoins_df[stablecoins_df["coin_id"] == coin_id][["coin_id", "symbol", "name"]].iloc[0]
        available_stablecoins.append({
            "coin_id": coin_info["coin_id"],
            "symbol": coin_info["symbol"],
            "name": coin_info["name"]
        })
    else:
        missing_stablecoins.append(coin_id)

if available_stablecoins:
    print("\n找到的指定穩定幣:")
    print(pd.DataFrame(available_stablecoins).to_string(index=False))
else:
    print("\n未找到任何指定穩定幣！")

if missing_stablecoins:
    print("\n缺少的指定穩定幣:")
    for coin_id in missing_stablecoins:
        print(f"  {coin_id}")

# === 步驟2: 過濾指定穩定幣 ===
stable_df = df[(df["classification"] == stablecoin_class) & (df["coin_id"].isin(SELECTED_STABLECOINS))]
if stable_df.empty:
    print("\n錯誤：無符合指定穩定幣的數據！")
    exit()

print(f"\n過濾後保留 {stable_df['coin_id'].nunique()} 個穩定幣進行分析")

# === 步驟3: 計算 AV, CAV, AAV, CAAV ===
# BTC 作為市場指數（使用 log_volume）
btc_df = df[df["coin_id"] == "bitcoin"][["date", "log_volume"]].rename(columns={"log_volume": "market_volume"})

if btc_df.empty:
    print("\n錯誤：數據中未找到 BTC！")
    exit()
print("\nBTC 市場交易量 (最近 5 天):")
print(btc_df.tail())

# 事件研究法函數（針對交易量）
def run_volume_model(coin_df, btc_df, estimation_start=-150, estimation_end=-11, event_start=-10, event_end=10):
    # 合併幣種數據與 BTC 市場交易量
    tmp = coin_df[["date", "log_volume", "rel_day"]].rename(columns={"log_volume": "vol"})
    tmp = tmp.merge(btc_df, on="date", how="inner")

    # 估計期 (-150 到 -11)
    est_df = tmp[(tmp["rel_day"] >= estimation_start) & (tmp["rel_day"] <= estimation_end)].dropna()
    if est_df.empty or len(est_df) < 30:
        return None, "Insufficient estimation data"

    # 模型回歸：vol = alpha + beta * market_volume + epsilon
    X = sm.add_constant(est_df["market_volume"])
    Y = est_df["vol"]
    try:
        model = sm.OLS(Y, X).fit()
        alpha, beta = model.params["const"], model.params["market_volume"]
    except Exception as e:
        return None, f"Regression failed: {str(e)}"

    # 事件期 (-10 到 +10)
    evt_df = tmp[(tmp["rel_day"] >= event_start) & (tmp["rel_day"] <= event_end)].copy()
    if evt_df.empty:
        return None, "Empty event data"

    # 計算預期交易量和異常交易量 (AV)
    evt_df["expected_vol"] = alpha + beta * evt_df["market_volume"]
    evt_df["AV"] = evt_df["vol"] - evt_df["expected_vol"]

    return evt_df, None

# 主程式：針對指定穩定幣運行
av_all = []
excluded_coins = []

for coin_id in stable_df["coin_id"].unique():
    print(f"\n處理穩定幣: {coin_id}")
    coin_df = stable_df[stable_df["coin_id"] == coin_id]
    evt_df, reason = run_volume_model(coin_df, btc_df)

    if evt_df is None:
        coin_info = stable_df[stable_df["coin_id"] == coin_id][["coin_id", "symbol", "name"]].iloc[0]
        excluded_coins.append({
            "coin_id": coin_info["coin_id"],
            "symbol": coin_info["symbol"],
            "name": coin_info["name"],
            "reason": reason
        })
        continue

    evt_df["coin_id"] = coin_id
    av_all.append(evt_df)

# 合併所有 AV
if av_all:
    av_df = pd.concat(av_all, ignore_index=True)

    # 計算 AAV 和 CAAV
    aav_df = (
        av_df[(av_df["rel_day"] >= -10) & (av_df["rel_day"] <= 10)]
        .groupby("rel_day")["AV"]
        .mean()
        .reset_index()
        .rename(columns={"AV": "AAV"})
    )
    aav_df["CAAV"] = aav_df["AAV"].cumsum()

    # 對 AAV 和 CAAV 進行 t 檢定
    results = []
    for day in aav_df["rel_day"]:
        # AAV t 檢定
        av_group = av_df[av_df["rel_day"] == day]["AV"]
        aav_t_stat, aav_p_value = stats.ttest_1samp(av_group, popmean=0, nan_policy="omit")

        # CAV t 檢定 (每個幣種的 CAV 到該天)
        cav_list = []
        for coin, group in av_df.groupby("coin_id"):
            cav_val = group[(group["rel_day"] >= -10) & (group["rel_day"] <= day)]["AV"].sum()
            cav_list.append(cav_val)
        caav_t_stat, caav_p_value = stats.ttest_1samp(cav_list, popmean=0, nan_policy="omit")

        results.append({
            "rel_day": day,
            "AAV": aav_df.loc[aav_df["rel_day"] == day, "AAV"].values[0],
            "CAAV": aav_df.loc[aav_df["rel_day"] == day, "CAAV"].values[0],
            "AAV_t_stat": aav_t_stat,
            "AAV_p_value": aav_p_value,
            "CAAV_t_stat": caav_t_stat,
            "CAAV_p_value": caav_p_value
        })

    # 轉為 DataFrame 並儲存
    results_df = pd.DataFrame(results).sort_values("rel_day")
    AAV_CAAV_XLSX = f"{OUTPUT_DIR}/AAV_CAAV_results_2025-07-17_selected_stablecoins.xlsx"
    results_df.to_excel(AAV_CAAV_XLSX, index=False)
    print(f"\n指定穩定幣的 AAV 和 CAAV 結果已儲存至 {AAV_CAAV_XLSX}")

    # 事件日前後 3 天檢查
    check_window = (-3, 3)
    check_df = results_df[(results_df["rel_day"] >= check_window[0]) & (results_df["rel_day"] <= check_window[1])]
    CHECK_WINDOW_XLSX = f"{OUTPUT_DIR}/AAV_CAAV_check_window_2025-07-17_selected_stablecoins.xlsx"
    check_df.to_excel(CHECK_WINDOW_XLSX, index=False)
    print(f"指定穩定幣的 3 天檢查結果已儲存至 {CHECK_WINDOW_XLSX}")

    print("\n=== 事件日前後 3 天檢查 (AAV / CAAV / t-test) ===")
    print(check_df.to_string(index=False))
else:
    print("\n指定穩定幣無有效 AV 數據。")

# 列出排除的穩定幣
if excluded_coins:
    print("\n被排除的穩定幣:")
    df_excluded = pd.DataFrame(excluded_coins)
    print(df_excluded[["coin_id", "symbol", "name", "reason"]])
else:
    print("\n無穩定幣被排除。")