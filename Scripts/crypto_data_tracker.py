import requests
import pandas as pd
import time
from datetime import datetime
import logging
import os
import win32com.client as win32
import pythoncom

# Configure logging
logging.basicConfig(filename='crypto_data.log', level=logging.INFO, 
                    format='%(asctime)s - %(levelname)s - %(message)s')

<<<<<<< HEAD
# CoinGecko API (free public API)
COINGECKO_API_URL = 'https://api.coingecko.com/api/v3/coins/markets'
=======
# CoinMarketCap API (sandbox) key
CMC_SANDBOX_API_KEY = 'API_KEY'

# Define the expected columns for the DataFrame
expected_columns = [
    "Name", "Symbol", "Current Price (USD)", "Market Cap", 
    "24h Trading Volume", "24h Price Change (%)"
]
>>>>>>> 8c8f0a2cd8b8e5208af83f638b4ea01dbd9e275d

def fetch_top_cryptocurrencies():
    """
    Fetch top cryptocurrencies from CoinGecko API.
    """
    try:
        # Parameters for fetching top 50 cryptocurrencies by market cap
        params = {
            'vs_currency': 'usd',
            'order': 'market_cap_desc',
            'per_page': 50,
            'page': 1,
            'sparkline': 'false'
        }

        # Send request to CoinGecko API
        response = requests.get(COINGECKO_API_URL, params=params)
        response.raise_for_status()

        # Extract relevant data
        crypto_data = []
        for coin in response.json():
            record = {
                "Name": coin['name'],
                "Symbol": coin['symbol'].upper(),
                "Current Price (USD)": round(coin['current_price'], 2),
                "Market Cap": coin['market_cap'],
                "24h Trading Volume": coin['total_volume'],
                "24h Price Change (%)": round(coin['price_change_percentage_24h'], 2)
            }
            crypto_data.append(record)

        # Create DataFrame
        crypto_df = pd.DataFrame(crypto_data)
        return crypto_df

    except requests.exceptions.RequestException as e:
        logging.error(f"API request error: {e}")
        print(f"Error fetching data: {e}")
        return None

def perform_data_analysis(df):
    """
    Perform comprehensive analysis on cryptocurrency data.
    """
    if df is None or df.empty:
        logging.warning("DataFrame is empty or None. Skipping analysis.")
        return {}

    try:
        # Prepare analysis dictionary
        analysis = {
            'Total Cryptocurrencies': len(df),
            'Top 5 Cryptocurrencies by Market Cap': ', '.join(df.nlargest(5, 'Market Cap')['Name'].tolist()),
            'Average Price (USD)': round(df['Current Price (USD)'].mean(), 2),
            'Median Price (USD)': round(df['Current Price (USD)'].median(), 2),
            'Highest 24h Price Change (%)': round(df['24h Price Change (%)'].max(), 2),
            'Lowest 24h Price Change (%)': round(df['24h Price Change (%)'].min(), 2),
            'Coin with Highest 24h Price Change': df.loc[df['24h Price Change (%)'].idxmax(), 'Name'],
            'Coin with Lowest 24h Price Change': df.loc[df['24h Price Change (%)'].idxmin(), 'Name'],
            'Total Market Cap (USD)': df['Market Cap'].sum(),
        }

        return analysis

    except Exception as e:
        logging.error(f"Error performing data analysis: {e}")
        print(f"Error in analysis: {e}")
        return {}

def update_excel_sheet():
    """
    Continuously update WPS Office Excel sheet with cryptocurrency data.
    """
    pythoncom.CoInitialize()  # Important for multi-threaded COM usage
    
    try:
        # Determine the correct file path
        script_dir = os.path.dirname(os.path.abspath(__file__))
        excel_file_path = os.path.join(script_dir, 'Cryptocurrency_Tracker.xlsx')
        
        # Open or create Excel workbook
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        excel.Visible = True
        excel.DisplayAlerts = False  # Suppress any warning dialogs
        
        # Create or open workbook
        try:
            wb = excel.Workbooks.Open(excel_file_path)
        except:
            wb = excel.Workbooks.Add()
            wb.SaveAs(excel_file_path)
        
        # Prepare sheets
        sheet_names = [sheet.Name for sheet in wb.Sheets]
        
        # Cryptocurrency Data Sheet
        ws_data = wb.Sheets('Cryptocurrency Data') if 'Cryptocurrency Data' in sheet_names \
                  else wb.Sheets.Add(After=wb.Sheets(wb.Sheets.Count))
        ws_data.Name = 'Cryptocurrency Data'
        
        # Market Analysis Sheet
        ws_analysis = wb.Sheets('Market Analysis') if 'Market Analysis' in sheet_names \
                      else wb.Sheets.Add(After=wb.Sheets(wb.Sheets.Count))
        ws_analysis.Name = 'Market Analysis'

        while True:
            # Fetch live data
            crypto_df = fetch_top_cryptocurrencies()
            
            if crypto_df is not None and not crypto_df.empty:
                # Cryptocurrency Data Sheet
                ws_data.Cells.Clear()
                
                # Write headers
                headers = crypto_df.columns.tolist()
                for col, header in enumerate(headers, start=1):
                    ws_data.Cells(1, col).Value = header
                
                # Write data
                for row in range(len(crypto_df)):
                    for col, header in enumerate(headers, start=1):
                        ws_data.Cells(row+2, col).Value = crypto_df.iloc[row][header]
                
                # Perform analysis
                analysis = perform_data_analysis(crypto_df)
                
                # Market Analysis Sheet
                ws_analysis.Cells.Clear()
                ws_analysis.Cells(1, 1).Value = 'Analysis Metric'
                ws_analysis.Cells(1, 2).Value = 'Value'
                
                for row, (metric, value) in enumerate(analysis.items(), start=2):
                    ws_analysis.Cells(row, 1).Value = metric
                    ws_analysis.Cells(row, 2).Value = str(value)
                
                # Auto-fit columns
                ws_data.Columns.AutoFit()
                ws_analysis.Columns.AutoFit()
                
                # Save workbook
                wb.Save()
                
                logging.info("Cryptocurrency data updated successfully.")
                print("Data updated at:", datetime.now())
            
            # Wait for 5 minutes before next update
            time.sleep(300)

    except Exception as e:
        logging.error(f"Excel update error: {e}")
        print(f"Error updating Excel: {e}")
        print(f"Error details: {repr(e)}")
    finally:
        try:
            wb.Close(SaveChanges=True)
            excel.Quit()
        except:
            pass
        pythoncom.CoUninitialize()

if __name__ == '__main__':
    update_excel_sheet()