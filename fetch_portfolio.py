import yfinance as yf
import pandas as pd


yf.set_tz_cache_location("C:/temp/yfinance_cache")


tickers = {
    "TotalEnergies": "TTE.PA",
    "LVMH":          "MC.PA",
    "BNP Paribas":   "BNP.PA",
    "Airbus":        "AIR.PA",
    "Sanofi":        "SAN.PA"
}


raw = yf.download(list(tickers.values()), period="2y", auto_adjust=True)["Close"]
raw.columns = list(tickers.keys())


raw = raw.dropna(axis=1, how="all")   
raw = raw.dropna(axis=0, how="any")  

print(f" Stocks loaded: {list(raw.columns)}")
print(f" Date range: {raw.index[0].date()} → {raw.index[-1].date()}")
print(f" Rows: {len(raw)}")

if raw.empty:
    raise ValueError(" No data loaded at all — check your internet connection")

#  Daily returns
returns = raw.pct_change().dropna()

# Cumulative returns
cumulative = (1 + returns).cumprod() - 1

# Volatility (annualized, 30-day rolling)
volatility = returns.rolling(30).std() * (252 ** 0.5)

#  Sharpe Ratio (annualized, risk-free rate ~3.5% for EUR)
risk_free = 0.035 / 252
sharpe_daily = (returns - risk_free).mean() / returns.std()
sharpe_annual = sharpe_daily * (252 ** 0.5)

# Summary table (one row per stock)
summary = pd.DataFrame({
    "Total Return (%)":   (cumulative.iloc[-1] * 100).round(2),
    "Annualized Vol (%)": (volatility.iloc[-1] * 100).round(2),
    "Sharpe Ratio":       sharpe_annual.round(2),
    "Last Price (€)":     raw.iloc[-1].round(2)
})
summary.index.name = "Stock"

# Add ticker column
ticker_map = {v: k for k, v in tickers.items()}
summary.insert(0, "Ticker", [tickers[s] for s in summary.index])

#  Export to Excel
with pd.ExcelWriter("portfolio_data.xlsx", engine="openpyxl") as writer:
    raw.reset_index().to_excel(writer, sheet_name="Prices", index=False)
    returns.reset_index().to_excel(writer, sheet_name="Daily Returns", index=False)
    cumulative.reset_index().to_excel(writer, sheet_name="Cumulative Returns", index=False)
    volatility.reset_index().to_excel(writer, sheet_name="Volatility", index=False)
    summary.reset_index().to_excel(writer, sheet_name="Summary", index=False)

print("\n portfolio_data.xlsx generated successfully!")
print(summary)