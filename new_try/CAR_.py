import pandas as pd
import numpy as np
import statsmodels.api as sm
from datetime import datetime, timedelta

# === 基本設定 ===
INPUT_DIR = r"C:\Users\Administrator\Desktop\論文\crypto_data2"
INPUT_XLSX = f"{INPUT_DIR}/historical_data_2025-07-17_2.xlsx"
OUTPUT_DIR = INPUT_DIR
EVENT_DATE = pd.to_datetime("2025-07-17")
CAR_WINDOW = (-3, 3)  # CAR 窗口：t=-3 到 t=3
STABLECOIN_CLASS = "5. 穩定幣"

# === 讀取數據 ===
df = pd.read_excel(INPUT_XLSX)

# 日期平移
df["date"] = pd.to_datetime(df["date"]) - timedelta(days=1)  # 平移日期

# 確保按 coin_id 和 date 排序
df = df.sort_values(["coin_id", "date"])

# 加上相對天數
df["rel_day"] = (df["date"] - EVENT_DATE).dt.days

# 排除穩定幣
non_stable_df = df[df["classification"] != STABLECOIN_CLASS].copy()
if non_stable_df.empty:
    print(f"錯誤：排除 '{STABLECOIN_CLASS}' 後無數據！")
    raise SystemExit
print(f"\n排除穩定幣後包含 {non_stable_df['coin_id'].nunique()} 個幣種")

# === 提取 BTC 作為市場指數 ===
btc_df = df[df["coin_id"] == "bitcoin"][["date", "log_return"]].rename(columns={"log_return": "mkt_return"})
if btc_df.empty:
    print("錯誤：數據中未找到 BTC！")
    raise SystemExit
print("\nBTC 市場指數 (最近 5 天):")
print(btc_df.tail())

# === 事件研究法函數（計算 CAR） ===
def run_market_model(coin_df, btc_df, estimation_start=-150, estimation_end=-11, event_start=0, event_end=5):
    tmp = coin_df[["date", "log_return", "rel_day", "symbol", "name", "coin_id"]].rename(columns={"log_return": "ret"})
    tmp = tmp.merge(btc_df, on="date", how="inner")

    # 估計期
    est_df = tmp[(tmp["rel_day"] >= estimation_start) & (tmp["rel_day"] <= estimation_end)].dropna()
    if est_df.empty or len(est_df) < 30:
        return None, None, "Insufficient estimation data"

    # 市場模型回歸
    X = sm.add_constant(est_df["mkt_return"])
    Y = est_df["ret"]
    try:
        model = sm.OLS(Y, X).fit()
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
print(f"\n排除穩定幣後包含 {non_stable_df['coin_id'].nunique()} 個幣種")

car_rows, excluded = [], []

for cid, g in non_stable_df.groupby("coin_id"):
    if cid == "bitcoin":  # 跳過 BTC
        continue
    evt_df, car, reason = run_market_model(g, btc_df, event_start=CAR_WINDOW[0], event_end=CAR_WINDOW[1])
    if evt_df is None:
        coin_info = g[["coin_id", "symbol", "name"]].iloc[0]
        excluded.append({"coin_id": coin_info["coin_id"], "symbol": coin_info["symbol"], "name": coin_info["name"], "reason": reason})
        continue

    # 控制變數：log 市值（t=0，若缺用 t=-1）
    ev = g[g["rel_day"] == 0]
    if ev.empty:
        ev = g[g["rel_day"] == -1]
    log_mcap = ev["log_market_cap"].iloc[0] if ("log_market_cap" in g and not ev.empty) else np.nan

    # 動能：t=-10 ~ -4 的累積 log_return（統一與前版）
    mom_win = g[(g["rel_day"] >= -10) & (g["rel_day"] <= -4)]
    momentum = mom_win["log_return"].sum() if not mom_win.empty else np.nan

    # ILLIQ（事件前 30 天平均）
    illiq_win = g[(g["rel_day"] >= -33) & (g["rel_day"] <= -4)].copy()
    illiq_win["ILLIQ"] = illiq_win["log_return"].abs() / (illiq_win["volume"] / 1000)
    illiq = illiq_win["ILLIQ"].mean() if not illiq_win.empty else np.nan

    # Volatility：事件前 30 天日報酬標準差（不含事件期，非 rolling）
    pre30 = g[(g["rel_day"] >= -10) & (g["rel_day"] <= -4)]
    vol30 = pre30["log_return"].std(ddof=1) if pre30["log_return"].count() >= 5 else np.nan

    car_rows.append({
        "coin_id": cid,
        "symbol": g["symbol"].iloc[0],
        "name": g["name"].iloc[0],
        "classification": g["classification"].iloc[0],
        "CAR": car,
        "log_market_cap": log_mcap,
        "momentum": momentum,
        "ILLIQ": illiq,
        "vol30": vol30
    })
# 轉為 DataFrame
car_df = pd.DataFrame(car_rows)
# 轉為 DataFrame
if car_df.empty:
    print("\n無有效 CAR 數據，無法進行回歸！")
    exit()

# 儲存 CAR 和控制變數
CAR_XLSX = f"{OUTPUT_DIR}/CAR_control_variables_2025-07-17_all_coins_no_stablecoin.xlsx"
car_df.to_excel(CAR_XLSX, index=False)
print(f"\nCAR 和控制變數已儲存至 {CAR_XLSX}")
print("\n=== CAR 和控制變數概覽 ===")
print(car_df[["coin_id", "symbol", "CAR", "log_market_cap", "momentum", "ILLIQ", "vol30"]].to_string(index=False))

# === 橫斷面回歸 ===
#def winsorize_series(series, lower=0.01, upper=0.99):
    #q_low = series.quantile(lower)
   # q_high = series.quantile(upper)
    #series = series.copy()
    #series[series < q_low] = q_low
    #series[series > q_high] = q_high
    #return series

# 清理數據
reg_df = car_df.dropna(subset=["CAR", "log_market_cap", "momentum", "ILLIQ", "vol30"]).copy()

# 對 CAR 截尾 (上下 1%)
#reg_df["CAR"] = winsorize_series(reg_df["CAR"], lower=0.01, upper=0.99)
if len(reg_df) < 5:
    print("\n有效樣本 < 5，無法進行穩健回歸。")
else:
    X = sm.add_constant(reg_df[["log_market_cap", "momentum", "ILLIQ", "vol30"]])
    y = reg_df["CAR"]
    model = sm.OLS(y, X).fit(cov_type="HC0")  # White robust SE
    print("\n=== 橫斷面回歸結果（All no stablecoin；HC0；未截尾） ===")
    print(model.summary())

    # 輸出
    CLASS_TAG = "all_no_stablecoin_BTCmkt"
    car_out = f"{OUTPUT_DIR}/CAR_{CLASS_TAG}_{EVENT_DATE.date()}_{CAR_WINDOW[0]}_{CAR_WINDOW[1]}.xlsx"
    reg_out = f"{OUTPUT_DIR}/crosssec_{CLASS_TAG}_HC0_{EVENT_DATE.date()}_{CAR_WINDOW[0]}_{CAR_WINDOW[1]}.xlsx"
    car_df.to_excel(car_out, index=False)

    results = {
        "variable": ["const", "log_market_cap", "momentum", "ILLIQ", "vol30"],
        "coefficient": model.params,
        "std_err_HC0": model.bse,
        "t_or_z": model.tvalues,
        "p_value": model.pvalues,
        "r_squared": [model.rsquared] + [np.nan] * 4,
        "r_squared_adj": [model.rsquared_adj] + [np.nan] * 4,
        "N": [int(model.nobs)] + [np.nan] * 4
    }
    pd.DataFrame(results).to_excel(reg_out, index=False)
    print(f"\n檔案已輸出：\n- CAR：{car_out}\n- 迴歸：{reg_out}")

    # 參考：同時做「只含常數」檢定，以對照平均 CAR 是否顯著
    only_c = sm.OLS(y, np.ones((len(y),1))).fit(cov_type="HC0")
    print("\n=== 僅常數回歸（檢驗平均 CAR 是否顯著） ===")
    print(only_c.summary())

# 列出排除的幣種
if excluded:
    print("\n被排除的幣種:")
    df_excluded = pd.DataFrame(excluded)
    print(df_excluded[["coin_id", "symbol", "name", "reason"]])
else:
    print("\n無幣種被排除。")