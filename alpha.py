import requests
import pandas as pd
from datetime import datetime, timedelta, timezone

def get_alpha_id(coin: str) -> str:
    """查 Token List，回傳指定 coin 的 alphaId (例如 'ALPHA_381')"""
    url = "https://www.binance.com/bapi/defi/v1/public/wallet-direct/buw/wallet/cex/alpha/all/token/list"
    res = requests.get(url).json()
    for t in res.get("data", []):
        if t.get("symbol", "").upper() == coin.upper():
            return t.get("alphaId")
    raise ValueError(f"{coin} not found in Alpha Token List")

def check_symbol(alpha_id: str, quote="USDT") -> str:
    """確認交易對是否有效，回傳正確 symbol"""
    symbol = f"{alpha_id}{quote}" if not alpha_id.startswith("ALPHA_") else f"{alpha_id}{quote}"
    url = "https://www.binance.com/bapi/defi/v1/public/alpha-trade/ticker"
    res = requests.get(url, params={"symbol": symbol}).json()
    if res.get("code") == "000000":
        return symbol
    else:
        raise ValueError(f"Invalid symbol {symbol}: {res}")

def fetch_alpha_klines(coin: str, date="2025-09-27", interval="5m", quote="USDT"):
    """抓 Alpha 幣的 K 線"""
    # 1. 找 alphaId
    alpha_id = get_alpha_id(coin)
    # 2. 確認交易對
    symbol = check_symbol(alpha_id, quote)
    print(f"Using symbol: {symbol}")

    # 3. 計算 UTC 毫秒時間範圍
    day = datetime.strptime(date, "%Y-%m-%d").replace(tzinfo=timezone.utc)
    start_ms = int(day.timestamp() * 1000)
    end_ms = int((day + timedelta(days=1)).timestamp() * 1000)

    # 4. 抓取 K 線
    url = "https://www.binance.com/bapi/defi/v1/public/alpha-trade/klines"
    params = {
        "symbol": symbol,
        "interval": interval,
        "startTime": start_ms,
        "endTime": end_ms,
        "limit": 1500
    }
    res = requests.get(url, params=params).json()

    if res.get("code") != "000000":
        raise RuntimeError(f"API error: {res}")

    cols = [
        "open_time","open","high","low","close","volume",
        "close_time","quote_volume","num_trades",
        "taker_base_vol","taker_quote_vol","ignore"
    ]
    df = pd.DataFrame(res["data"], columns=cols)

    # 5. 型別轉換
    df["open_time"] = pd.to_datetime(df["open_time"].astype(int), unit="ms")
    df["close_time"] = pd.to_datetime(df["close_time"].astype(int), unit="ms")
    for c in ["open","high","low","close","volume","quote_volume","taker_base_vol","taker_quote_vol"]:
        df[c] = df[c].astype(float)
    df["num_trades"] = df["num_trades"].astype(int)

    return df

# === 測試 ===
df = fetch_alpha_klines("RIVER", date="2025-09-27", interval="1m")
print(df.tail())
print("共抓到:", len(df), "根 K 線")
