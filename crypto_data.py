import requests
import pandas as pd

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

data = fetch_crypto_data()
for crypto in data[:5]:  # Print top 5 for testing
    print(crypto["name"], crypto["current_price"])


def analyze_crypto_data(data):
    df = pd.DataFrame(data, columns=[
        "name", "symbol", "current_price", "market_cap", 
        "total_volume", "price_change_percentage_24h"
    ])

    # Top 5 by market capitalization
    top_5 = df.nlargest(5, "market_cap")[["name", "market_cap"]]

    # Average price of top 50 cryptocurrencies
    average_price = df["current_price"].mean()

    # Highest and lowest 24-hour percentage price change
    highest_change = df.nlargest(1, "price_change_percentage_24h")
    lowest_change = df.nsmallest(1, "price_change_percentage_24h")

    print("\nTop 5 Cryptocurrencies by Market Cap:")
    print(top_5)

    print("\nAverage Price of Top 50 Cryptocurrencies: $", round(average_price, 2))

    print("\nHighest 24-hour Change:")
    print(highest_change[["name", "price_change_percentage_24h"]])

    print("\nLowest 24-hour Change:")
    print(lowest_change[["name", "price_change_percentage_24h"]])

    return df

data = fetch_crypto_data()
df = analyze_crypto_data(data)


from openpyxl import Workbook
import schedule
import time

def update_excel(df):
    with pd.ExcelWriter("crypto_data.xlsx", engine="openpyxl", mode="w") as writer:
        df.to_excel(writer, index=False, sheet_name="Top 50 Cryptos")
    print("Excel updated.")

def job():
    data = fetch_crypto_data()
    df = analyze_crypto_data(data)
    update_excel(df)

# Schedule updates every 5 minutes
schedule.every(5).minutes.do(job)

# Initial Run
job()

print("Starting live updates. Press Ctrl+C to stop.")
while True:
    schedule.run_pending()
    time.sleep(1)

