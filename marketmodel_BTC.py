import pandas as pd
import numpy as np

# 讀取你剛剛輸出的 Excel
df = pd.read_excel(r"C:\Users\Administrator\Desktop\論文\cryptodata\eventday_top100final_timeseries.xlsx")

# 日期往前平移一天，確保符合收盤價定義
df["date"] = pd.to_datetime(df["date"]) - pd.Timedelta(days=1)
# 確保按日期排序
df = df.sort_values(["coin_id", "date"])

# 計算對數報酬
df["log_return"] = df.groupby("coin_id")["price"].transform(lambda x: np.log(x / x.shift(1)))

# 檢查 BTC (市場指數)
btc_df = df[df["coin_id"]=="bitcoin"][["date","log_return"]].rename(columns={"log_return":"mkt_return"})
print(btc_df.tail())

# 假設 EVENT_DATE = 2025-07-17
EVENT_DATE = pd.to_datetime("2025-07-17")

# 加上相對天數
df["rel_day"] = (pd.to_datetime(df["date"]) - EVENT_DATE).dt.days

estimation_df = df[(df["rel_day"] >= -150) & (df["rel_day"] <= -11)]
event_df      = df[(df["rel_day"] >= -10)  & (df["rel_day"] <= 30)]

import statsmodels.api as sm
import pandas as pd

# 假設 df 已經有: date, coin_id, log_return, rel_day
# 先抽 BTC (市場指數)
btc = df[df["coin_id"] == "bitcoin"][["date", "log_return"]].rename(columns={"log_return": "mkt_return"})

# === 4) 儲存所有幣的 AR 資料 ===
ar_all = []
results = []  # 存各幣結果

for cid in df["coin_id"].unique():
    if cid == "bitcoin":  # BTC 當市場，不跑
        continue

    tmp = df[df["coin_id"] == cid][["date", "log_return", "rel_day"]].rename(columns={"log_return": "ret"})
    tmp = tmp.merge(btc, on="date", how="inner")

    # 估計期 (-150 ~ -11)
    est_df = tmp[(tmp["rel_day"] >= -150) & (tmp["rel_day"] <= -11)].dropna()
    if est_df.empty:
        continue

    X = sm.add_constant(est_df["mkt_return"])
    Y = est_df["ret"]
    try:
        model = sm.OLS(Y, X).fit()
        alpha, beta = model.params["const"], model.params["mkt_return"]
    except:
        continue

    # 事件期 (-10 ~ +30)
    event_df = tmp[(tmp["rel_day"] >= -10) & (tmp["rel_day"] <= 30)].copy()
    event_df["expected_ret"] = alpha + beta * event_df["mkt_return"]
    event_df["AR"] = event_df["ret"] - event_df["expected_ret"]
    event_df["coin"] = cid
    ar_all.append(event_df)

# 合併所有 AR
ar_df = pd.concat(ar_all, ignore_index=True)

from scipy import stats
# 轉成 DataFrame
daily_t_results = []
for day, group in ar_df.groupby("rel_day"):
    t_stat, p_value = stats.ttest_1samp(group["AR"], popmean=0, nan_policy="omit")
    daily_t_results.append({
        "rel_day": day,
        "mean_AR": group["AR"].mean(),
        "t_stat": t_stat,
        "p_value": p_value
    })

daily_t_df = pd.DataFrame(daily_t_results).sort_values("rel_day")
daily_t_df.to_excel(r"C:\Users\Administrator\Desktop\論文\cryptodata\Daily_AR_ttest3.xlsx", index=False)

# windows = [(-1, 1), (-3, 3), (-5, 5), (0, 1), (0, 3), (0, 5), (-10, 10), (-10, 30)]
window = (-10, 10)
caar_results = []

 # for (a, b) in windows:
 #    car_list = []
 #    for coin, group in ar_df.groupby("coin"):
 #        car_val = group[(group["rel_day"] >= a) & (group["rel_day"] <= b)]["AR"].sum()
 #        car_list.append(car_val)
 #
 #    t_stat, p_value = stats.ttest_1samp(car_list, popmean=0, nan_policy="omit")
 #    car_results.append({
 #        "window": f"[{a},{b}]",
 #        "mean_CAR": np.mean(car_list),
 #        "t_stat": t_stat,
 #        "p_value": p_value
 #    })
# 先算每天的 AAR
aar_df = (
    ar_df[(ar_df["rel_day"] >= window[0]) & (ar_df["rel_day"] <= window[1])]
    .groupby("rel_day")["AR"]
    .mean()
    .reset_index()
    .rename(columns={"AR": "AAR"})
)
# 依序累積得到 CAAR
aar_df["CAAR"] = aar_df["AAR"].cumsum()

# 對每天的 CAR 做 t 檢定 (這裡的 CAR 是到當天的累積)
for day in aar_df["rel_day"]:
    car_list = []
    for coin, group in ar_df.groupby("coin"):
        car_val = group[(group["rel_day"] >= window[0]) & (group["rel_day"] <= day)]["AR"].sum()
        car_list.append(car_val)

    t_stat, p_value = stats.ttest_1samp(car_list, popmean=0, nan_policy="omit")
    caar_results.append({
        "rel_day": day,
        "AAR": aar_df.loc[aar_df["rel_day"]==day, "AAR"].values[0],
        "CAAR": aar_df.loc[aar_df["rel_day"]==day, "CAAR"].values[0],
        "t_stat": t_stat,
        "p_value": p_value
    })

# 轉成 DataFrame
caar_df = pd.DataFrame(caar_results)

# === 輸出 Excel，覆蓋到 CAR_ttest.xlsx ===
caar_df.to_excel(r"C:\Users\Administrator\Desktop\論文\cryptodata\CAR_ttest3.xlsx", index=False)

# car_df = pd.DataFrame(car_results)
# car_df.to_excel(r"C:\Users\Administrator\Desktop\論文\cryptodata\CAR_ttest.xlsx", index=False)

# === 事件日前後檢查輸出 ===
check_window = (-3, 3)
check_df = caar_df[(caar_df["rel_day"] >= check_window[0]) & (caar_df["rel_day"] <= check_window[1])]

print("\n=== 事件日前後 3 天檢查 (AAR / CAAR / t-test) ===")
print(check_df.to_string(index=False))