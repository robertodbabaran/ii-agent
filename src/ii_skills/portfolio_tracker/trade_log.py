#!/usr/bin/env python3
"""
Trade Log — Synthesizes holdings from networth_config.py into trade records
and computes Daily_Units / Daily_Cash time series.

Initial holdings are treated as "BUY" trades at their cost basis on
the earliest available date. Real trades can be added via add_trade().
"""

from typing import Dict, List, Optional, Tuple
from datetime import datetime, timedelta
from dataclasses import dataclass, field
import copy


@dataclass
class Trade:
    """A single trade record."""
    date: datetime
    ticker: str
    action: str  # BUY, SELL, DIVIDEND
    quantity: float
    price: float
    currency: str = "CAD"
    fees: float = 0.0
    account: str = ""
    notes: str = ""

    @property
    def total_cost(self) -> float:
        """Total cost including fees."""
        if self.action == "BUY":
            return self.quantity * self.price + self.fees
        elif self.action == "SELL":
            return self.quantity * self.price - self.fees
        return 0.0


class TradeLog:
    """
    Manages trade records and computes daily portfolio state.

    Usage:
        log = TradeLog.from_holdings(config_holdings, cash_accounts)
        log.add_trade(Trade(...))
        daily_units = log.compute_daily_units(dates)
        daily_cash = log.compute_daily_cash(dates)
    """

    def __init__(self):
        self.trades: List[Trade] = []
        self.initial_cash: float = 0.0

    @classmethod
    def from_holdings(cls, holdings: List[Dict], cash_accounts: Optional[List[Dict]] = None,
                      base_date: Optional[datetime] = None) -> "TradeLog":
        """
        Synthesize trade log from networth_config.py holdings.

        Each holding becomes a BUY trade at cost basis on base_date.

        Args:
            holdings: List of holding dicts from MARKET_HOLDINGS config
                      Each has: name, ticker, quantity, cost_basis, currency, account, etc.
            cash_accounts: List of cash account dicts from CASH_ACCOUNTS config
            base_date: Date for initial trades (default: 1 year ago)

        Returns:
            TradeLog with synthesized trades
        """
        log = cls()

        if base_date is None:
            base_date = datetime.now() - timedelta(days=365)

        # Convert holdings to trades
        for h in holdings:
            ticker = h.get("ticker", "")
            if not ticker:
                continue

            quantity = float(h.get("quantity", 0))
            if quantity <= 0:
                continue

            # Use cost_basis per unit if available, else 0
            cost_basis = float(h.get("cost_basis", 0))
            if cost_basis > 0 and quantity > 0:
                price_per_unit = cost_basis / quantity
            else:
                price_per_unit = 0.0

            trade = Trade(
                date=base_date,
                ticker=ticker,
                action="BUY",
                quantity=quantity,
                price=price_per_unit,
                currency=h.get("currency", "CAD"),
                account=h.get("account", ""),
                notes="Initial holding from config",
            )
            log.trades.append(trade)

        # Sum cash accounts
        if cash_accounts:
            for acct in cash_accounts:
                balance = float(acct.get("balance", 0))
                log.initial_cash += balance

        # Sort trades by date
        log.trades.sort(key=lambda t: t.date)

        return log

    def add_trade(self, trade: Trade):
        """Add a trade and re-sort the log."""
        self.trades.append(trade)
        self.trades.sort(key=lambda t: t.date)

    def get_tickers(self) -> List[str]:
        """Get unique tickers from all trades."""
        tickers = []
        seen = set()
        for t in self.trades:
            if t.ticker not in seen:
                tickers.append(t.ticker)
                seen.add(t.ticker)
        return tickers

    def compute_daily_units(self, dates: List[datetime]) -> Dict[str, List[float]]:
        """
        Compute cumulative shares held for each ticker on each date.

        Args:
            dates: List of dates (from Market_Data)

        Returns:
            Dict mapping ticker to list of cumulative units per date
        """
        tickers = self.get_tickers()
        # Build cumulative units
        units = {ticker: [0.0] * len(dates) for ticker in tickers}

        # Group trades by ticker
        trades_by_ticker: Dict[str, List[Trade]] = {}
        for t in self.trades:
            trades_by_ticker.setdefault(t.ticker, []).append(t)

        for ticker in tickers:
            ticker_trades = trades_by_ticker.get(ticker, [])
            cumulative = 0.0
            trade_idx = 0

            for i, date in enumerate(dates):
                # Apply all trades on or before this date
                while trade_idx < len(ticker_trades) and ticker_trades[trade_idx].date <= date:
                    t = ticker_trades[trade_idx]
                    if t.action == "BUY":
                        cumulative += t.quantity
                    elif t.action == "SELL":
                        cumulative -= t.quantity
                    trade_idx += 1

                units[ticker][i] = cumulative

        return units

    def compute_daily_cash(self, dates: List[datetime]) -> List[float]:
        """
        Compute daily cash balance.

        Starting from initial_cash, adjusted by trade costs.

        Args:
            dates: List of dates

        Returns:
            List of cash balances per date
        """
        cash = [0.0] * len(dates)
        current_cash = self.initial_cash
        trade_idx = 0

        sorted_trades = sorted(self.trades, key=lambda t: t.date)

        for i, date in enumerate(dates):
            while trade_idx < len(sorted_trades) and sorted_trades[trade_idx].date <= date:
                t = sorted_trades[trade_idx]
                if t.action == "BUY":
                    current_cash -= t.total_cost
                elif t.action == "SELL":
                    current_cash += t.total_cost
                elif t.action == "DIVIDEND":
                    current_cash += t.quantity * t.price
                trade_idx += 1

            cash[i] = current_cash

        return cash

    def to_rows(self) -> List[List]:
        """Convert trades to rows for Excel output."""
        rows = []
        for t in self.trades:
            rows.append([
                t.date,
                t.ticker,
                t.action,
                t.quantity,
                t.price,
                t.currency,
                t.fees,
                t.account,
                t.notes,
            ])
        return rows
