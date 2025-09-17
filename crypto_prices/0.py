import requests

API_KEY = "e34b7dd5-b2e3-42c7-864b-b9d84f491f20"

url = "https://pro-api.coinmarketcap.com/v1/indices/list"
headers = {"Accepts": "application/json", "X-CMC_PRO_API_KEY": API_KEY}

resp = requests.get(url, headers=headers)
print(resp.json())
