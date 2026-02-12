#!/usr/bin/env python3
"""
Data Fetcher — yfinance reference data + historical prices.

Provides price history, reference data, and exchange rates for the
portfolio tracker workbook. Follows the same yfinance patterns as
networth_newsletter/networth.py.
"""

import yfinance as yf
import numpy as np
from typing import Dict, List, Optional, Tuple
from datetime import datetime, timedelta
from dataclasses import dataclass, field
import warnings

warnings.filterwarnings("ignore")


@dataclass
class ReferenceData:
    """Reference data for a single ticker."""
    ticker: str
    company_name: str = ""
    sector: str = "Unknown"
    industry: str = "Unknown"
    country: str = "Unknown"
    currency: str = "USD"
    beta: float = 1.0
    market_cap: float = 0.0
    asset_class: str = "Equity"


@dataclass
class PriceHistory:
    """Price history for multiple tickers aligned by date."""
    dates: List[datetime] = field(default_factory=list)
    prices: Dict[str, List[float]] = field(default_factory=dict)
    exchange_rates: List[float] = field(default_factory=list)  # USDCAD


def fetch_reference_data(tickers: List[str]) -> Dict[str, ReferenceData]:
    """
    Fetch reference data (sector, industry, country, beta, market cap) for tickers.

    Args:
        tickers: List of ticker symbols

    Returns:
        Dict mapping ticker to ReferenceData
    """
    result = {}

    # Classify special tickers
    crypto_tickers = {"BTC-USD", "ETH-USD", "SOL-USD", "DOGE-USD"}
    commodity_tickers = {"GC=F", "SI=F", "CL=F", "HG=F"}
    etf_tickers = {"SPY", "QQQ", "SOXX", "XLE", "XLF", "GLD", "SLV",
                   "COPX", "CPER", "TLT", "IEF"}

    for ticker in tickers:
        ref = ReferenceData(ticker=ticker)

        # Classify asset class
        if ticker in crypto_tickers:
            ref.asset_class = "Crypto"
            ref.sector = "Cryptocurrency"
            ref.country = "Global"
            ref.currency = "USD"
        elif ticker in commodity_tickers:
            ref.asset_class = "Commodity"
            ref.sector = "Commodities"
            ref.country = "Global"
            ref.currency = "USD"
        elif ticker in etf_tickers:
            ref.asset_class = "ETF"
        elif ticker.endswith(".TO") or ticker.endswith(".V"):
            ref.currency = "CAD"

        try:
            info = yf.Ticker(ticker).info
            ref.company_name = info.get("longName", info.get("shortName", ticker))
            if ref.sector == "Unknown":
                ref.sector = info.get("sector", "Unknown")
            ref.industry = info.get("industry", "Unknown")
            if ref.country == "Unknown":
                ref.country = info.get("country", "Unknown")
            ref.beta = info.get("beta", 1.0) or 1.0
            ref.market_cap = info.get("marketCap", 0) or 0
            if ref.asset_class == "ETF":
                ref.sector = info.get("category", "ETF")
                ref.industry = "Exchange Traded Fund"
        except Exception as e:
            print(f"Warning: Could not fetch reference data for {ticker}: {e}")

        result[ticker] = ref

    return result


def fetch_price_history(tickers: List[str], period: str = "1y") -> PriceHistory:
    """
    Fetch aligned historical prices for all tickers.

    Uses yfinance download for efficiency. Aligns all tickers to common dates.

    Args:
        tickers: List of ticker symbols
        period: Time period (e.g., "1y", "6mo", "2y")

    Returns:
        PriceHistory with aligned dates and prices
    """
    history = PriceHistory()

    if not tickers:
        return history

    # Include SPY as benchmark
    all_tickers = list(set(tickers + ["SPY"]))

    try:
        # Batch download for efficiency
        data = yf.download(all_tickers, period=period, auto_adjust=True, progress=False)

        if data.empty:
            print("Warning: No price data returned from yfinance")
            return history

        # Handle single-ticker vs multi-ticker DataFrame structure
        if len(all_tickers) == 1:
            close = data[["Close"]].copy()
            close.columns = [all_tickers[0]]
        else:
            close = data["Close"].copy()

        # Drop rows where ALL tickers are NaN
        close = close.dropna(how="all")

        # Forward-fill gaps (weekends, holidays for crypto vs equity)
        close = close.ffill()

        if close.empty:
            return history

        history.dates = [d.to_pydatetime() for d in close.index]

        for ticker in all_tickers:
            if ticker in close.columns:
                prices = close[ticker].tolist()
                # Replace remaining NaN with 0
                history.prices[ticker] = [
                    float(p) if not (isinstance(p, float) and np.isnan(p)) else 0.0
                    for p in prices
                ]

    except Exception as e:
        print(f"Warning: Batch download failed: {e}")
        # Fallback: fetch individually
        for ticker in all_tickers:
            try:
                stock = yf.Ticker(ticker)
                hist = stock.history(period=period)
                if not hist.empty:
                    if not history.dates:
                        history.dates = [d.to_pydatetime() for d in hist.index]
                    history.prices[ticker] = hist["Close"].tolist()
            except Exception as e2:
                print(f"Warning: Could not fetch {ticker}: {e2}")

    return history


def fetch_exchange_rate_history(period: str = "1y") -> Tuple[List[datetime], List[float]]:
    """
    Fetch USDCAD exchange rate history.

    Args:
        period: Time period

    Returns:
        Tuple of (dates, rates)
    """
    try:
        fx = yf.Ticker("USDCAD=X")
        hist = fx.history(period=period)
        if not hist.empty:
            dates = [d.to_pydatetime() for d in hist.index]
            rates = hist["Close"].tolist()
            return dates, rates
    except Exception as e:
        print(f"Warning: Could not fetch USDCAD: {e}")

    return [], []


def fetch_current_price(ticker: str) -> Optional[float]:
    """Fetch the latest price for a single ticker."""
    try:
        stock = yf.Ticker(ticker)
        hist = stock.history(period="1d")
        if not hist.empty:
            return float(hist["Close"].iloc[-1])
    except Exception as e:
        print(f"Warning: Could not fetch price for {ticker}: {e}")
    return None


def fetch_exchange_rate() -> float:
    """Fetch current USDCAD exchange rate. Fallback 1.35."""
    try:
        fx = yf.Ticker("USDCAD=X")
        hist = fx.history(period="1d")
        if not hist.empty:
            return float(hist["Close"].iloc[-1])
    except Exception:
        pass
    return 1.35
