import requests
import pandas as pd
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from time import sleep

# Constants
API_URL = "https://api.coingecko.com/api/v3/coins/markets"
PARAMS = {
    "vs_currency": "usd",
    "order": "market_cap_desc",
    "per_page": 50,
    "page": 1,
    "sparkline": False
}
UPDATE_INTERVAL = 300  # seconds (5 minutes)
EXCEL_FILE = "crypto_data.xlsx"

# Fetch live cryptocurrency data
def fetch_crypto_data():
    response = requests.get(API_URL, params=PARAMS)
    response.raise_for_status()
    return response.json()

# Analyze data
def analyze_data(data):
    df = pd.DataFrame(data)
    df = df[["name", "symbol", "current_price", "market_cap", "total_volume", "price_change_percentage_24h"]]
    
    top_5_by_market_cap = df.nlargest(5, "market_cap")[["name", "market_cap"]]
    average_price = df["current_price"].mean()
    highest_change = df.nlargest(1, "price_change_percentage_24h")[["name", "price_change_percentage_24h"]]
    lowest_change = df.nsmallest(1, "price_change_percentage_24h")[["name", "price_change_percentage_24h"]]
    
    return df, top_5_by_market_cap, average_price, highest_change, lowest_change

# Update Excel file
def update_excel(df, analysis):
    top_5, avg_price, high_change, low_change = analysis
    
    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl", mode="a", if_sheet_exists="overlay") as writer:
        df.to_excel(writer, sheet_name="Live Data", index=False)
        
        # Add analysis results to a separate sheet
        analysis_sheet = writer.book.create_sheet("Analysis")
        analysis_sheet.append(["Metric", "Value"])
        analysis_sheet.append(["Top 5 by Market Cap", top_5.to_string(index=False)])
        analysis_sheet.append(["Average Price", avg_price])
        analysis_sheet.append(["Highest 24h Change", high_change.to_string(index=False)])
        analysis_sheet.append(["Lowest 24h Change", low_change.to_string(index=False)])

# Main loop
def main():
    while True:
        try:
            # Fetch and process data
            crypto_data = fetch_crypto_data()
            df, top_5, avg_price, high_change, low_change = analyze_data(crypto_data)

            # Update Excel
            update_excel(df, (top_5, avg_price, high_change, low_change))

            print("Data updated successfully in Excel.")
        except Exception as e:
            print(f"An error occurred: {e}")

        sleep(UPDATE_INTERVAL)

if __name__ == "__main__":
    main()
