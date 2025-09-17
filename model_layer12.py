import pandas as pd
import numpy as np
import statsmodels.api as sm
from scipy import stats

# === 1) 讀取幣別資料 ===
df = pd.read_excel(r"C:\Users\Administrator\Desktop\論文\cryptodata\eventday_layer12_timeseries.xlsx")
# === 日期往前平移一天，確保是收盤日 ===
df["date"] = pd.to_datetime(df["date"])- pd.Timedelta(days=1)  # 保持 datetime

# 確保按日期排序
df = df.sort_values(["coin_id", "date"])

# 計算對數報酬
df["log_return"] = df.groupby("coin_id")["price"].transform(lambda x: np.log(x / x.shift(1)))

# === 2) 讀取 CMC100 指數 ===
cmc = pd.read_excel(r"C:\Users\Administrator\Desktop\論文\cryptodata\CMC100.xlsx")
cmc["date"] = pd.to_datetime(cmc["date"])- pd.Timedelta(days=1)  # 保持 datetime，日期往前平移一天，確保是收盤日
cmc = cmc.rename(columns={"log_return": "mkt_return"})[["date", "mkt_return"]]

# === 3) 設定事件日 ===
EVENT_DATE = pd.to_datetime("2025-07-17")  # 事件日保持 datetime
df["rel_day"] = (df["date"] - EVENT_DATE).dt.days

# === 4) 儲存所有幣的 AR 資料 ===
ar_all = []

for cid in df["coin_id"].unique():
    tmp = df[df["coin_id"] == cid][["date", "log_return", "rel_day"]].rename(columns={"log_return": "ret"})
    tmp = tmp.merge(cmc, on="date", how="inner")

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

# === 5) 每日 AR 的 t-test ===
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
daily_t_df.to_excel(r"C:\Users\Administrator\Desktop\論文\cryptodata\Daily_AR_ttest_cmclayer1.xlsx", index=False)

# === 6) CAAR 計算 ===
window = (-10, 10)
caar_results = []

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

# 對每天的 CAR 做 t 檢定
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

# === 輸出 Excel ===
caar_df.to_excel(r"C:\Users\Administrator\Desktop\論文\cryptodata\CAR_ttest_cmclayer1.xlsx", index=False)