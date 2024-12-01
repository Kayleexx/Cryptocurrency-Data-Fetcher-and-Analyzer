import requests
import pandas as pd
import time
from datetime import datetime
import logging
import os

# Configure logging
logging.basicConfig(filename='crypto_data.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# CoinMarketCap API (sandbox) key
CMC_SANDBOX_API_KEY = 'f5c16d6f-c1d4-4dc3-a222-19f16679aaa7'

# Define the expected columns for the DataFrame
expected_columns = [
    "Name", "Symbol", "Current Price (USD)", "Market Cap", 
    "24h Trading Volume", "24h Price Change (%)"
]

def fetch_top_cryptocurrencies():
    """
    Fetch top cryptocurrencies from CoinMarketCap Sandbox API.
    """
    url = 'https://sandbox-api.coinmarketcap.com/v1/cryptocurrency/listings/latest'
    parameters = {
        'start': '1',
        'limit': '50',  # Fetch top 50 cryptocurrencies
        'convert': 'USD'
    }
    headers = {
        'Accepts': 'application/json',
        'X-CMC_PRO_API_KEY': CMC_SANDBOX_API_KEY,
    }

    try:
        response = requests.get(url, headers=headers, params=parameters)
        response.raise_for_status()  # Raise exception for HTTP errors

        data = response.json()
        crypto_data = []

        # Loop through each coin's data and ensure all records have the same fields
        for coin in data.get("data", []):
            record = {
                "Name": coin.get("name", None),
                "Symbol": coin.get("symbol", None),
                "Current Price (USD)": coin.get("quote", {}).get("USD", {}).get("price", None),
                "Market Cap": coin.get("quote", {}).get("USD", {}).get("market_cap", None),
                "24h Trading Volume": coin.get("quote", {}).get("USD", {}).get("volume_24h", None),
                "24h Price Change (%)": coin.get("quote", {}).get("USD", {}).get("percent_change_24h", None),
            }
            crypto_data.append(record)

        # Create DataFrame ensuring all records have the same fields
        crypto_df = pd.DataFrame(crypto_data, columns=expected_columns)
        return crypto_df

    except requests.exceptions.RequestException as e:
        logging.error(f"API request error: {e}")
        print(f"Error: {e}")
        return None
    except Exception as e:
        logging.error(f"General error in fetching data: {e}")
        print(f"Error: {e}")
        return None

def perform_data_analysis(df):
    """
    Perform basic analysis on cryptocurrency data.
    """
    if df is None or df.empty:
        logging.warning("DataFrame is empty or None. Skipping analysis.")
        return {}

    try:
        analysis = {
            'Top 5 Cryptocurrencies by Market Cap': df.nlargest(5, 'Market Cap')['Name'].tolist(),
            'Average Price': df['Current Price (USD)'].mean(),
            'Highest 24h Price Change': df.loc[df['24h Price Change (%)'].idxmax(), 'Name'],
            'Lowest 24h Price Change': df.loc[df['24h Price Change (%)'].idxmin(), 'Name'],
            'Highest 24h Price Change Value': df['24h Price Change (%)'].max(),
            'Lowest 24h Price Change Value': df['24h Price Change (%)'].min()
        }

        # Ensure the analysis values are not pandas objects, but plain values
        analysis = {k: v.item() if hasattr(v, 'item') else v for k, v in analysis.items()}
        return analysis
    except Exception as e:
        logging.error(f"Error performing data analysis: {e}")
        print(f"Error: {e}")
        return {}

def update_csv_file():
    """
    Update CSV file with live cryptocurrency data.
    """
    try:
        file_path = 'crypto_data.csv'

        while True:
            crypto_df = fetch_top_cryptocurrencies()

            if crypto_df is not None and len(crypto_df) > 0:
                # Write the dataframe to CSV
                crypto_df.to_csv(file_path, index=False, mode='w', header=True)

                analysis = perform_data_analysis(crypto_df)
                if analysis: 
                    with open(file_path, 'a') as f:
                        f.write("\nAnalysis Results:\n")
                        for key, value in analysis.items():
                            f.write(f"{key}: {value}\n")

                    # Log last updated time
                    with open(file_path, 'a') as f:
                        f.write(f"Last Updated: {datetime.now()}\n")

                logging.info("Data successfully updated in CSV.")
            else:
                logging.warning("No data retrieved from the API.")

            time.sleep(300)  # Sleep for 5 minutes

    except Exception as e:
        logging.error(f"Error updating CSV: {e}")
        print(f"Error: {e}")

if __name__ == '__main__':
    update_csv_file()
