import requests
import time
import pandas as pd
import os
from datetime import datetime, timezone

# === 基本設定 ===
API_KEY = "CG-38G8j12guYuvo8ksrALeu9Du"  # 你的 Pro API key
VS = "usd"
EVENT_DATE = datetime(2025, 7, 17, tzinfo=timezone.utc)  # 過去日期測試

OUTPUT_DIR = r"C:\Users\Administrator\Desktop\論文\crypto_data2"
os.makedirs(OUTPUT_DIR, exist_ok=True)
OUT_XLSX = os.path.join(OUTPUT_DIR, "coins_list_event_day_2025-07-17_top300.xlsx")  # 更新檔案名


# === 基本工具 ===
def cg_get(url, params=None):
    p = params.copy() if params else {}
    p["x_cg_demo_api_key"] = API_KEY
    try:
        r = requests.get(url, params=p, timeout=30)
        r.raise_for_status()
    except requests.exceptions.RequestException as e:
        print(f"API Error: {e}, URL: {url}, Params: {p}")
        return []  # 返回空列表，跳過該幣
    return r.json()


# 步驟1: 抓取事件日市值前 300 幣 (每頁 250，page 1 和 2)
def get_top_300_coins_on_date(event_dt):
    date_str = event_dt.strftime('%Y-%m-%d')
    all_coins = []
    for page in range(1, 3):  # page 1 和 page 2，總共 500 幣
        url = f"https://api.coingecko.com/api/v3/coins/markets"
        params = {
            "vs_currency": VS,
            "order": "market_cap_desc",
            "per_page": 150,  # 每頁 250 幣
            "page": page,
            "sparkline": False,
            "date": date_str
        }
        data = cg_get(url, params)
        coins = [{"id": coin["id"], "symbol": coin["symbol"], "name": coin["name"], "market_cap": coin["market_cap"]}
                 for coin in data]
        all_coins.extend(coins)
        print(f"Fetched {len(coins)} coins for page {page}")
        time.sleep(5)  # 每頁間隔 5 秒
    print(f"Total coins fetched: {len(all_coins)}")
    return all_coins


# 步驟2: 獲取單一幣的 categories
def get_coin_categories(coin_id):
    url = f"https://api.coingecko.com/api/v3/coins/{coin_id}"
    params = {"localization": "false", "market_data": "false"}
    data = cg_get(url, params)
    categories = data.get("categories", [])
    return categories


# 步驟3: 分類邏輯（包含未分類到 Excel）
def classify_and_filter(categories):
    exclude_categories = ["Tokenized BTC", "Tokenized Assets","Wrapped-Tokens", "Liquid Restaking Tokens", " Yield-Bearing Stablecoin","Liquid Staking Tokens", "Tokenized BTC", "Bridged WETH","Bridged Stablecoin", "Bridged USDT", "Bridged USDC","Bridged WBTC", "Crypto-Backed Tokens", "Midas Liquid Yield Tokens", "LP Tokens", "Binance-Peg Tokens","Wallets"]  # 排除衍生幣

    if any(ex_cat in categories for ex_cat in exclude_categories):
        print(f"Excluded coin with categories: {categories}")  # Debug 排除幣
        return None

    # 優先級 5: 穩定幣
    if any(cat in categories for cat in ["Stablecoins", "USD Stablecoin", "Fiat-backed Stablecoin", "Crypto-Backed Stablecoin", "Yield-Bearing Stablecoin"]):
        print(f"Matched 5. 穩定幣 for categories: {categories}")  # Debug
        return "5. 穩定幣"

    # 優先級 3: 中心化交易所平台幣 (檢查無 Layer0/1 條件)
    has_cex = any(cat in categories for cat in ["Centralized Exchange (CEX) Token", "Bitcoin Fork"])
    has_layer01 = any(cat in categories for cat in ["Layer 1 (L1)", "Smart Contract Platform", "BNB Chain Ecosystem"])
    if has_cex and not has_layer01:
        print(f"Matched 3. 中心化交易所平台幣 for categories: {categories}")  # Debug
        return "3. 中心化交易所平台幣"

    # 優先級 1: Layer0/1 區塊鏈基礎層
    if any(cat in categories for cat in ["Layer 1 (L1)", "Layer 0 (L0)", "GMCI Layer 1 Index", "Polkadot Ecosystem", "Cosmos Ecosystem", "DePIN", "Cross-chain Communication"]):
        print(f"Matched 1. Layer0/1 區塊鏈基礎層 for categories: {categories}")  # Debug
        return "1. Layer0/1 區塊鏈基礎層"

    # 優先級 2: Layer2/3 擴展應用層代幣
    if any(cat in categories for cat in ["Layer 2 (L2)", "Decentralized Finance (DeFi)", "Decentralized Exchange (DEX)","Rollups-as-a-Service (RaaS)", "Automated Market Maker", "NFT", "Play-to-Earn", "Metaverse", "Oracle", "Yield Farming", "Lending/Borrowing", "Derivatives", "Perpetuals", "Artificial Intelligence (AI)", "Gaming (GameFi)", "Gambling (GambleFi)", "SocialFi", "Zero Knowledge (ZK)", "Decentralized Science (DeSci)", "Decentralized Identifier (DID)", "Payment Solutions", "ZkSync Ecosystem"
]):
        print(f"Matched 2. Layer2/3 擴展應用層代幣 for categories: {categories}")  # Debug
        return "2. Layer2/3 擴展應用層代幣"

    # 優先級 4: 迷因幣 (優先於 Layer0/1)
    if any(cat in categories for cat in ["Meme", "Dog-Themed","Frog-Themed", "AI Meme"]):
        print(f"Matched 4. 迷因幣 for categories: {categories}")  # Debug
        return "4. 迷因幣"

    # 優先級 6: 未分類
    print(f"Unclassified categories: {categories}")  # Debug 未分類幣
    return "6. 未分類"


# 主程式
top_coins = get_top_300_coins_on_date(EVENT_DATE)
results = []

for i, coin in enumerate(top_coins):
    print(f"Processing coin {i + 1}/{len(top_coins)}: {coin['symbol']}")  # 進度顯示
    categories = get_coin_categories(coin["id"])
    print(f"  Categories for {coin['symbol']}: {categories}")  # Debug categories
    classification = classify_and_filter(categories)

    if classification:  # 包含 "未分類" 情況
        print(f"  Classified as: {classification}")
        results.append({
            "ID": coin["id"],
            "Symbol": coin["symbol"],
            "Name": coin["name"],
            "Market Cap": coin["market_cap"],
            "Categories": ", ".join(categories),
            "Classification": classification
        })

    time.sleep(10)  # 增加到 7 秒間隔，降低超時風險

# 輸出 Excel
df = pd.DataFrame(results)
df.to_excel(OUT_XLSX, index=False)
print(f"輸出到 {OUT_XLSX}")
print(f"最終樣本: {len(df)} 種幣")
print("\n分類分佈:")
print(df["Classification"].value_counts())