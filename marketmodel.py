import pandas as pd
import numpy as np

# 讀取你剛剛輸出的 Excel
df = pd.read_excel(r"C:\Users\Administrator\Desktop\論文\cryptodata\eventday_top200_timeseries.xlsx")

# 確保按日期排序
df = df.sort_values(["coin_id", "date"])

# 計算對數報酬
df["log_return"] = df.groupby("coin_id")["price"].transform(lambda x: np.log(x / x.shift(1)))

# 檢查 BTC (市場指數)
btc_df = df[df["coin_id"]=="bitcoin"][["date","log_return"]].rename(columns={"log_return":"mkt_return"})
print(btc_df.tail())

print(df.head())