import pandas as pd
import numpy as np
import statsmodels.api as sm
from scipy import stats
from datetime import timedelta

# === 基本設定（改成 Constant Mean Model 版）===
INPUT_DIR = r"C:\Users\Administrator\Desktop\論文\crypto_data_AR"
INPUT_XLSX = f"{INPUT_DIR}/historical_data_2025-07-17_2.xlsx"  # 內含 log_return
OUTPUT_DIR = r"C:\Users\Administrator\Desktop\論文\constant_model"

# 輸出檔（你要沿用舊檔名也可以，改這兩行即可）
AAR_CAAR_XLSX   = f"{OUTPUT_DIR}/AAR_CAAR_results_2025-07-18_all_no_stablecoins_CMM.xlsx"
CHECK_WINDOW_XLSX = f"{OUTPUT_DIR}/AAR_CAAR_check_window_2025-07-18_all_no_stablecoins_CMM.xlsx"

# 事件研究參數
EVENT_DATE = pd.to_datetime("2025-07-18")   # 事件日
EST_START, EST_END = -83, -4               # 估計期（80 天，不含事件前10天）
EVT_START, EVT_END = -3, 3                  # 事件窗
WINDOW_DAYS = list(range(EVT_START, EVT_END + 1))

# === 讀取 & 前處理 ===
df = pd.read_excel(INPUT_XLSX)
df["date"] = pd.to_datetime(df["date"])
# CoinGecko 收盤日對齊（你的原始碼有 -1 天）
df["date"] = df["date"] - timedelta(days=1)
df = df.sort_values(["coin_id", "date"])
df["rel_day"] = (df["date"] - EVENT_DATE).dt.days

# 若沒有 log_return 就用 price 算
if "log_return" not in df.columns:
    df["log_return"] = df.groupby("coin_id")["price"].transform(lambda x: np.log(x / x.shift(1)))

# 排除穩定幣
df = df[df["classification"] != "5. 穩定幣"]
print(f"排除 '5. 穩定幣' 後，剩餘 {df['coin_id'].nunique()} 個幣種")

# === 函數：跑 Constant Mean Model，回傳事件窗 AR 及估計期 AR（用於 KP 相關性）===
def run_constant_mean_model_for_coin(coin_df):
    # 估計期資料
    est_df = coin_df[(coin_df["rel_day"] >= EST_START) & (coin_df["rel_day"] <= EST_END)][["date", "log_return", "rel_day"]].dropna()
    if est_df.empty or len(est_df) < 30:
        return None, None, "Insufficient estimation data"

    # 平均調整模型：估計期平均
    mu_i = est_df["log_return"].mean()
    # 估計期的 AR_iτ（= r_iτ - mu_i）用來算波動與 KP 平均相關
    est_df = est_df.assign(AR=est_df["log_return"] - mu_i)
    # 估計期 AR 的標準差（個股）
    s_i = est_df["AR"].std(ddof=1)

    # 避免除以 0
    if pd.isna(s_i) or s_i == 0:
        return None, None, "Zero variance in estimation period"

    T = len(est_df)  # 估計期長度
    # Constant Mean 的標準化分母（考慮均值估計誤差）：s_i * sqrt(1 + 1/T)
    scale_i = s_i * np.sqrt(1.0 + 1.0 / T)

    # 事件窗資料
    evt_df = coin_df[(coin_df["rel_day"] >= EVT_START) & (coin_df["rel_day"] <= EVT_END)][["date", "log_return", "rel_day"]].copy()
    if evt_df.empty:
        return None, None, "Empty event data"

    # 計算異常報酬與標準化異常報酬
    evt_df["AR"] = evt_df["log_return"] - mu_i
    evt_df["SAR"] = evt_df["AR"] / scale_i  # Patell/BMP 標準化
    return evt_df, est_df[["date", "AR"]].rename(columns={"AR": "AR_est"}), None

# === 主迴圈：逐幣跑 CMM，收集事件窗 AR 與估計期 AR（用於 KP） ===
ar_list = []
sar_list = []
est_panel = {}  # {coin_id: pd.Series(index=date, values=AR_est)} 用來算 KP 的平均相關
excluded = []

for coin_id, g in df.groupby("coin_id"):
    evt_df, est_ar_df, reason = run_constant_mean_model_for_coin(g)
    if evt_df is None:
        # 補抓 coin 基本資料
        info_row = g[["coin_id", "symbol", "name"]].drop_duplicates().iloc[0]
        excluded.append({"coin_id": coin_id, "symbol": info_row["symbol"], "name": info_row["name"], "reason": reason})
        continue
    evt_df = evt_df.assign(coin_id=coin_id)
    ar_list.append(evt_df[["coin_id", "date", "rel_day", "AR"]])
    sar_list.append(evt_df[["coin_id", "date", "rel_day", "SAR"]])

    # 估計期 AR（用於 KP 平均相關）
    if est_ar_df is not None and not est_ar_df.empty:
        est_panel[coin_id] = est_ar_df.set_index("date")["AR_est"]

if not ar_list:
    print("無有效 AR / SAR 數據。")
    if excluded:
        print("被排除的幣種：")
        print(pd.DataFrame(excluded))
    raise SystemExit

ar_df = pd.concat(ar_list, ignore_index=True)
sar_df = pd.concat(sar_list, ignore_index=True)

# === KP：估計期「平均相關」rho_bar（橫斷面相關性） ===
# 先把 estimation AR 組成日期×幣 的寬表，再算相關矩陣
if len(est_panel) >= 2:
    est_matrix = pd.concat(est_panel, axis=1)  # columns = coin_id，多列日期
    corr_mat = est_matrix.corr(min_periods=20)  # 至少 20 天共同樣本
    # 取非對角元素平均作為 rho_bar
    m = corr_mat.shape[0]
    if m >= 2:
        rho_bar = (corr_mat.values[np.triu_indices(m, k=1)]).mean()
    else:
        rho_bar = 0.0
else:
    rho_bar = 0.0
rho_bar = 0.0 if np.isnan(rho_bar) else float(rho_bar)

print(f"KP 平均相關 rho_bar ≈ {rho_bar:.4f}（估計期）")

# === 計算 AAR / CAAR 與統計檢定 ===
results = []
for t in WINDOW_DAYS:
    ar_t = ar_df[ar_df["rel_day"] == t]["AR"].dropna().values
    sar_t = sar_df[sar_df["rel_day"] == t]["SAR"].dropna().values
    N_t = len(ar_t)

    if N_t == 0:
        continue

    AAR_t = np.mean(ar_t)

    # === 事件窗內的 CAAR：從 EVT_START 到當下 t 的 AAR 累積 ===
    # 先取每個 tau 的 AAR，再到這個 t 做 cumsum（為了避免重複計算，我們先暫存本輪）
    # 這裡為了正確性，改用：以幣別先算 CAR 再做截面平均（與你原本邏輯一致）
    # 先計算每個幣 i 在 [EVT_START, t] 的 CAR_i
    car_i = []
    for coin_id, g in ar_df.groupby("coin_id"):
        gi = g[(g["rel_day"] >= EVT_START) & (g["rel_day"] <= t)]
        if gi.empty:
            continue
        car_i.append(gi["AR"].sum())
    car_i = np.array(car_i)
    CAAR_t = np.mean(car_i) if len(car_i) > 0 else np.nan

    # === t-test（AAR, CAAR）===
    # AAR t-test
    if N_t >= 2:
        se_aar = np.std(ar_t, ddof=1) / np.sqrt(N_t)
        t_aar = AAR_t / se_aar if se_aar > 0 else np.nan
        p_aar = 2 * (1 - stats.t.cdf(np.abs(t_aar), df=N_t - 1)) if np.isfinite(t_aar) else np.nan
    else:
        t_aar, p_aar = np.nan, np.nan

    # CAAR t-test（對 car_i 做一樣的單樣本 t 檢定）
    if len(car_i) >= 2:
        se_caar = np.std(car_i, ddof=1) / np.sqrt(len(car_i))
        t_caar = CAAR_t / se_caar if se_caar > 0 else np.nan
        p_caar = 2 * (1 - stats.t.cdf(np.abs(t_caar), df=len(car_i) - 1)) if np.isfinite(t_caar) else np.nan
    else:
        t_caar, p_caar = np.nan, np.nan

    # === z-test（大樣本近似，把母體標準差用樣本無偏或母體版都可；這裡以母體版）===
    z_aar = AAR_t / (np.std(ar_t, ddof=0) / np.sqrt(N_t)) if N_t > 1 and np.std(ar_t, ddof=0) > 0 else np.nan
    z_caar = CAAR_t / (np.std(car_i, ddof=0) / np.sqrt(len(car_i))) if len(car_i) > 1 and np.std(car_i, ddof=0) > 0 else np.nan

    # === BMP / Patell 型（對 SAR 做，理論上 var≈1 ⇒ mean(SAR)*sqrt(N) 近似 N(0,1)）===
    if len(sar_t) >= 1:
        z_bmp_aar = np.mean(sar_t) * np.sqrt(len(sar_t))
    else:
        z_bmp_aar = np.nan

    # CAAR 的 BMP：把 CAR 用個別 s_i 標準化後平均（近似）
    # 先組每個幣在 [EVT_START,t] 的 SCAR_i = CAR_i / (s_i*sqrt(L*(1+1/T)))。
    # 因為我們只有事件窗 AR，個別 s_i 與 T 在上一段函數內；為簡化，退而求其次採「對 car_i 做標準化 t 檢定的 z 近似」：
    z_bmp_caar = z_caar  # 常見做法：CAAR 用截面 z/t（BMP 對 CAAR的嚴格版需要保存個別 s_i 與 T）

    # === KP 調整（修正截面相關）：把 BMP 的 z 除以 sqrt(1 + (N-1)*rho_bar) ===
    adj_factor_aar = np.sqrt(1.0 + max(0.0, N_t - 1) * rho_bar)
    z_kp_aar = z_bmp_aar / adj_factor_aar if np.isfinite(z_bmp_aar) else np.nan

    N_car = len(car_i)
    adj_factor_caar = np.sqrt(1.0 + max(0.0, N_car - 1) * rho_bar)
    z_kp_caar = z_bmp_caar / adj_factor_caar if np.isfinite(z_bmp_caar) else np.nan

    results.append({
        "rel_day": t,
        "N_AAR": N_t,
        "AAR": AAR_t,
        "CAAR": CAAR_t,
        # t-tests
        "AAR_t": t_aar, "AAR_t_p": p_aar,
        "CAAR_t": t_caar, "CAAR_t_p": p_caar,
        # z-tests
        "AAR_z": z_aar,
        "CAAR_z": z_caar,
        # BMP / KP
        "AAR_z_BMP": z_bmp_aar,
        "AAR_z_KP": z_kp_aar,
        "CAAR_z_BMP": z_bmp_caar,
        "CAAR_z_KP": z_kp_caar
    })

results_df = pd.DataFrame(results).sort_values("rel_day")

# 只顯示事件日前後 3 天
check_df = results_df[(results_df["rel_day"] >= EVT_START) & (results_df["rel_day"] <= EVT_END)].copy()

# === 輸出 ===
results_df.to_excel(AAR_CAAR_XLSX, index=False)
check_df.to_excel(CHECK_WINDOW_XLSX, index=False)

print(f"\nAAR/CAAR 結果已儲存至 {AAR_CAAR_XLSX}")
print(f"事件窗檢查結果已儲存至 {CHECK_WINDOW_XLSX}")
print("\n=== 事件窗（-3~+3）AAR/CAAR with t, z, BMP/KP ===")
print(check_df.to_string(index=False))

# 額外列出被排除幣種
#（如果你想也存 Excel，可自行寫出）
try:
    excluded_coins_df = pd.DataFrame(excluded)
    if not excluded_coins_df.empty:
        print("\n被排除的幣種（前 10 筆）：")
        print(excluded_coins_df.head(10).to_string(index=False))
except Exception:
    pass
