import yfinance as yf
import matplotlib.pyplot as plt

# Define tickers
stocks = ['TSLA', 'AMZN', 'META', 'MSFT']
cryptos = ['BTC-USD', 'ETH-USD', 'XRP-USD', 'SOL-USD', 'ADA-USD']
sp500 = '^GSPC'
assets = stocks + [sp500] + cryptos

# Fetch data
start_date = '2022-01-01'
end_date = '2024-12-01'
data = yf.download(assets, start=start_date, end=end_date)['Adj Close']

# Debug: Print raw data
print("Raw data:")
print(data)

# Handle missing data by forward-filling and back-filling
data = data.ffill().bfill()

# Debug: Print cleaned data
print("\nCleaned data (after filling missing values):")
print(data)

# Normalize data for growth comparison
normalized_data = data / data.iloc[0] * 100

# Debug: Print normalized data
print("\nNormalized data:")
print(normalized_data)

# Plot configuration
plt.figure(figsize=(14, 10))
crypto_colors = ['lightblue', 'lightgreen', 'lightcoral', 'lightpink', 'lightyellow']
stock_colors = ['navy', 'darkred', 'purple', 'orange']
sp500_color = 'black'

# Plot stocks
for i, stock in enumerate(stocks):
    if stock in normalized_data.columns:
        plt.plot(normalized_data[stock], label=stock, color=stock_colors[i])

# Plot S&P 500
if sp500 in normalized_data.columns:
    plt.plot(normalized_data[sp500], label='S&P 500', color=sp500_color, linestyle='--')

# Plot cryptocurrencies
for i, crypto in enumerate(cryptos):
    if crypto in normalized_data.columns:
        plt.plot(normalized_data[crypto], label=crypto, color=crypto_colors[i])

# Add chart details
plt.title("Stock and Cryptocurrency Growth vs S&P 500 (2022-2024)", fontsize=14)
plt.xlabel("Date", fontsize=12)
plt.ylabel("Normalized Growth (%)", fontsize=12)
plt.legend(loc='upper left', fontsize=10)
plt.grid(True)
plt.tight_layout()

# Show plot
plt.show()
