import requests, os, math, time
import pandas as pd
from datetime import datetime, timezone

# === 基本設定 ===
API_KEY = "CG-38G8j12guYuvo8ksrALeu9Du"  # 換成你的 CoinGecko demo key
VS = "usd"
EVENT_DATE = datetime(2025, 7, 17)   # 事件日

OUTPUT_DIR = r"C:\Users\Administrator\Desktop\論文\cryptodata"
os.makedirs(OUTPUT_DIR, exist_ok=True)
OUT_XLSX = os.path.join(OUTPUT_DIR, "eventday_top152_marketcap.xlsx")

# === 事件日 Top 幣 IDs ===
coin_ids = [
"aave","aerodrome-finance","algorand","apecoin","apenft","aptos","arbitrum",
"arweave","avalanche-2","based-brett","beldex","binancecoin","bitcoin",
"bitcoin-cash","bitcoin-cash-sv","bitget-token","bittensor","bittorrent",
"blockstack","bonk","build-on","cardano","celestia","chainlink","conflux-token",
"cosmos","crypto-com-chain","curve-dao-token","dai","decentraland","dogecoin",
"dogwifcoin","dydx-chain","eigenlayer","ethena","ethena-usde","ether-fi",
"ethereum","ethereum-classic","ethereum-name-service","falcon-finance",
"fartcoin","fasttoken","fetch-ai","filecoin","first-digital-usd","flare-networks",
"floki","flow","four","gala","gatechain-token","global-dollar",
"hedera-hashgraph","helium","hyperliquid","immutable-x","injective-protocol",
"instadapp","internet-computer","iota","jasmycoin","jito-governance-token",
"jupiter-exchange-solana","jupiter-perpetuals-liquidity-provider-token","kaia",
"kaspa","keeta","kucoin-shares","leo-token","lido-dao","litecoin",
"mantle","memecore","monero","morpho","near","neo","newton-project","nexo",
"official-trump","okb","ondo-finance","ondo-us-dollar-yield","optimism",
"pancakeswap-token","pax-gold","paypal-usd","pendle","pepe","pi-network",
"polkadot","polygon-ecosystem-token","pudgy-penguins","pump-fun","pyth-network",
"quant-network","raydium","render-token","reserve-rights-token","ripple",
"saros-finance","sei-network","shiba-inu","sky","solana","sonic-3","spx6900",
"starknet","stellar","story-2","sui","susds","syrup","syrupusdc","telcoin",
"tether","tether-gold","tezos","the-graph","the-open-network","the-sandbox",
"theta-token","tron","true-usd","uniswap","usd-coin","usd1-wlfi","usdd","usds",
"usdtb","usdx-money-usdx","usual-usd","vaulta","vechain","virtual-protocol",
"vision-3","walrus-2","whitebit","worldcoin-wld","world-liberty-financial","zcash"
]

# === 基本工具 ===
def cg_get(url, params=None):
    p = params.copy() if params else {}
    p["x_cg_demo_api_key"] = API_KEY
    r = requests.get(url, params=p, timeout=30)
    r.raise_for_status()
    return r.json()

def fetch_eventday(coin_id, vs, event_dt):
    """只抓事件日 (市值)"""
    url = f"https://api.coingecko.com/api/v3/coins/{coin_id}/market_chart/range"
    start_unix = int(event_dt.replace(tzinfo=timezone.utc).timestamp())
    end_unix   = int((event_dt.replace(tzinfo=timezone.utc) + pd.Timedelta(days=1)).timestamp())
    js = cg_get(url, {"vs_currency": vs, "from": start_unix, "to": end_unix})

    arr = js.get("market_caps", [])
    if not arr:
        return None
    df = pd.DataFrame(arr, columns=["ts", "market_cap"])
    df["date"] = pd.to_datetime(df["ts"], unit="ms", utc=True).dt.date
    df = df.groupby("date", as_index=False).last()
    df.insert(0, "coin_id", coin_id)
    return df[["date","coin_id","market_cap"]]

# === 主流程 ===
rows = []
for i, cid in enumerate(coin_ids, 1):
    print(f"[{i}/{len(coin_ids)}] 抓 {cid} …")
    try:
        df = fetch_eventday(cid, VS, EVENT_DATE)
        if df is not None:
            rows.append(df)
    except Exception as e:
        print(f"  [跳過] {cid}: {e}")
    time.sleep(5)  # 防止被限流

data = pd.concat(rows, ignore_index=True).sort_values("coin_id")

# 輸出
import openpyxl
with pd.ExcelWriter(OUT_XLSX, engine="openpyxl") as writer:
    data.to_excel(writer, index=False, sheet_name="eventday")

print("完成，已輸出：", OUT_XLSX)