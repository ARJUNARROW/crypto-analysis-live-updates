import requests
import pandas as pd
from openpyxl import Workbook

def fetch_crypto_data():
    url = "https://api.coingecko.com/api/v3/coins/markets"
    params = {
        "vs_currency": "usd",
        "order": "market_cap_desc",
        "per_page": 50,
        "page": 1,
        "sparkline": False
    }
    response = requests.get(url, params=params)
    if response.status_code == 200:
        return response.json()
    else:
        print("Error fetching data:", response.status_code)
        return []

def create_dataframe(data):
    df = pd.DataFrame(data, columns=[
        "name", "symbol", "current_price", "market_cap", 
        "total_volume", "price_change_percentage_24h"
    ])
    return df


def write_to_excel(df, filename="crypto_data.xlsx"):
    with pd.ExcelWriter(filename, engine="openpyxl", mode="w") as writer:
        df.to_excel(writer, index=False, sheet_name="Top 50 Cryptos")
    print(f"Data written to {filename}")


import schedule
import time

def update_data():
    print("Fetching live data...")
    data = fetch_crypto_data()
    if data:
        df = create_dataframe(data)
        write_to_excel(df)
    else:
        print("Failed to fetch data.")

# Schedule updates every 5 minutes
schedule.every(5).minutes.do(update_data)

# Initial run
update_data()

print("Live updating started. Press Ctrl+C to stop.")
while True:
    schedule.run_pending()
    time.sleep(1)

