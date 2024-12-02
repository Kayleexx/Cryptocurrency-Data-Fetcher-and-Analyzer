# Cryptocurrency Data Fetcher and Analyzer

This Python project fetches live cryptocurrency data from the CoinMarketCap Sandbox API, performs basic analysis, and stores the data in a CSV file. The project also logs the results and analysis into a separate log file for tracking.

## Features

- Fetches top 50 cryptocurrencies by market cap from the CoinMarketCap API.
- Analyzes key metrics such as:
  - Top 5 cryptocurrencies by market cap.
  - Average price of all fetched cryptocurrencies.
  - Highest and lowest 24-hour price change.
- Saves the cryptocurrency data into a CSV file.
- Appends analysis results to the CSV file.
- Logs all activities, errors, and updates in a log file (`crypto_data.log`).
- Updates the data every 5 minutes.

## Prerequisites

- Python 3.6+
- `requests` library for making API requests.
- `pandas` library for handling data and creating DataFrames.

You can install the required libraries by running the following:

```bash
pip install requests pandas
```

## API Key

This project uses the CoinMarketCap Sandbox API. You need a valid API key to fetch cryptocurrency data.

1. Go to the [CoinMarketCap API](https://coinmarketcap.com/api/) website.
2. Create an account or log in to get your API key.
3. Replace the `CMC_SANDBOX_API_KEY` variable with your actual API key in the script.

## How to Run

1. Clone the repository to your local machine:

```bash
git clone https://github.com/Kayleexx/Cryptocurrency-Data-Fetcher-and-Analyzer.git
cd crypto-data-fetcher
```


2. Run the script:

```bash
python crypto_data_fetcher.py
```

### Data will be updating every 5 minutes and following is the output:
![image](https://github.com/user-attachments/assets/acc278f8-8c55-4316-a04c-388ceb76966a)



## Log File

The log file (`crypto_data.log`) contains information about the script's execution, including:

- Successful data fetches
- Errors during API requests
- Data analysis results
- Time of last update


