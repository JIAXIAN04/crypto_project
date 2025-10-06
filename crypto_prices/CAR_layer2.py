import pandas as pd
import numpy as np
import statsmodels.api as sm
from datetime import timedelta
from scipy import stats

# === 基本設定 ===
INPUT_DIR = r"C:\Users\Administrator\Desktop\論文\crypto_data2"
INPUT_XLSX = f"{INPUT_DIR}/historical_data_2025-07-17_2.xlsx"
CMC20_XLSX = r"C:\Users\Administrator\Desktop\論文\cryptodata\CMC20.xlsx"
OUTPUT_DIR = r"C:\Users\Administrator\Desktop\論文\crypto_data_cmc"
EVENT_DATE = pd.to_datetime("2025-07-17")
CAR_WINDOW = (-3, 3)  # t=0 到 t=5
START_CLASS = '2. Layer2/3 擴展應用層代幣'

# === 讀取個幣資料 ===
df = pd.read_excel(INPUT_XLSX)
df["date"] = pd.to_datetime(df["date"]) - timedelta(days=1)  # 收盤對齊
df = df.sort_values(["coin_id", "date"])
df["rel_day"] = (df["date"] - EVENT_DATE).dt.days

# === 讀取 CMC20 指數 (Rm) ===
cmc = pd.read_excel(CMC20_XLSX)
cmc["date"] = pd.to_datetime(cmc["date"]) - pd.Timedelta(days=1)
cmc = cmc.rename(columns={"log_return": "mkt_return"})[["date", "mkt_return"]]

# === 僅保留指定分類（嚴格匹配；若你的分類字串偶有空白/大小寫問題，可改用 contains 版本） ===
class_df = df[df["classification"] == START_CLASS].copy()
# class_df = df[df["classification"].str.contains("Layer2/3", na=False)].copy()  # ← 若需要寬鬆匹配，改用這行
if class_df.empty:
    print(f"錯誤：分類 {START_CLASS} 找不到資料！")
    raise SystemExit

print(f"\n分類 {START_CLASS} 幣種數：{class_df['coin_id'].nunique()}")

# === 事件研究：用 CMC20 市場模型拿到每個幣的 CAR ===
def run_market_model(coin_df, cmc_df, estimation_start=-150, estimation_end=-11,
                     event_start=-3, event_end=3):
    tmp = coin_df[["date", "log_return", "rel_day", "symbol", "name", "coin_id"]].rename(columns={"log_return": "ret"})
    tmp = tmp.merge(cmc_df, on="date", how="inner")

    est_df = tmp[(tmp["rel_day"] >= estimation_start) & (tmp["rel_day"] <= estimation_end)].dropna()
    if len(est_df) < 30:
        return None, None, "Insufficient estimation data"

    X = sm.add_constant(est_df["mkt_return"])
    Y = est_df["ret"]
    try:
        model = sm.OLS(Y, X).fit()
        alpha, beta = model.params["const"], model.params["mkt_return"]
    except Exception as e:
        return None, None, f"Regression failed: {str(e)}"

    evt_df = tmp[(tmp["rel_day"] >= event_start) & (tmp["rel_day"] <= event_end)].copy()
    if evt_df.empty:
        return None, None, "Empty event data"

    evt_df["expected_ret"] = alpha + beta * evt_df["mkt_return"]
    evt_df["AR"] = evt_df["ret"] - evt_df["expected_ret"]
    car = evt_df["AR"].sum()
    return evt_df, car, None

# === 逐幣計算 CAR 與控制變數（只在指定分類內） ===
car_rows, excluded = [], []
for cid, g in class_df.groupby("coin_id"):
    evt_df, car, reason = run_market_model(g, cmc, event_start=CAR_WINDOW[0], event_end=CAR_WINDOW[1])
    if evt_df is None:
        coin_info = g[["coin_id", "symbol", "name"]].iloc[0]
        excluded.append({"coin_id": coin_info["coin_id"], "symbol": coin_info["symbol"], "name": coin_info["name"], "reason": reason})
        continue

    # 控制變數：log 市值（t=0，若缺用 t=-1）
    ev = g[g["rel_day"] == 0]
    if ev.empty:
        ev = g[g["rel_day"] == -1]
    log_mcap = ev["log_market_cap"].iloc[0] if ("log_market_cap" in g and not ev.empty) else np.nan

    # 動能：t=-21 ~ -1 的累積 log_return（可按你的定義調整）
    mom_win = g[(g["rel_day"] >= -10) & (g["rel_day"] <= -4)]
    momentum = mom_win["log_return"].sum() if not mom_win.empty else np.nan

    # ILLIQ（事件前 30 天平均）
    illiq_win = g[(g["rel_day"] >= -33) & (g["rel_day"] <= -4)].copy()
    illiq_win["ILLIQ"] = illiq_win["log_return"].abs() / (illiq_win["volume"] / 1000)
    illiq = illiq_win["ILLIQ"].mean() if not illiq_win.empty else np.nan

    # === Volatility：事件前 30 天日報酬標準差（不含事件期，非 rolling）
    pre30 = g[(g["rel_day"] >= -10) & (g["rel_day"] <= -4)]
    vol30 = pre30["log_return"].std(ddof=1) if pre30["log_return"].count() >= 5 else np.nan

    car_rows.append({
        "coin_id": cid,
        "symbol": g["symbol"].iloc[0],
        "name": g["name"].iloc[0],
        "classification": START_CLASS,
        "CAR": car,
        "log_market_cap": log_mcap,
        "momentum": momentum,
        "ILLIQ": illiq,
        "vol30": vol30
    })

car_df = pd.DataFrame(car_rows)
print(f"\n可用樣本（{START_CLASS}）數量：{len(car_df)}")
if excluded:
    print("\n被排除的幣（估計期不足/回歸失敗/事件窗無資料）：")
    print(pd.DataFrame(excluded)[["coin_id","symbol","name","reason"]].to_string(index=False))

if car_df.empty:
    print("\n無有效 CAR 資料，無法回歸。")
    raise SystemExit

# === 橫斷面回歸：CAR ~ log_market_cap + momentum + ILLIQ + log_volume（White/HC0 標準誤；不截尾） ===

reg_df = car_df.dropna(subset=["CAR", "log_market_cap", "momentum", "ILLIQ", "vol30"]).copy()
if len(reg_df) < 5:
    print("\n有效樣本 < 5，無法進行穩健回歸。")
else:
    X = sm.add_constant(reg_df[["log_market_cap", "momentum", "ILLIQ", "vol30"]])
    y = reg_df["CAR"]
    model = sm.OLS(y, X).fit(cov_type="HC0")  # White robust SE
    print("\n=== 橫斷面回歸結果（Layer2/3；HC0；未截尾） ===")
    print(model.summary())

    # 輸出
    CLASS_TAG = "Layer23"
    car_out = f"{OUTPUT_DIR}/CAR_{CLASS_TAG}_only_{EVENT_DATE.date()}_{CAR_WINDOW[0]}_{CAR_WINDOW[1]}.xlsx"
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
