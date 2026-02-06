#!/usr/bin/env python3
"""
Risk Engine — Extended risk metrics for the portfolio tracker.

Wraps portfolio_analytics.py patterns with additional metrics:
Sortino ratio, VaR 99%, CVaR 95%, individual betas, stress scenarios.
All computed in Python for the Risk_Metrics and Stress_Testing sheets.
"""

import numpy as np
from typing import Dict, List, Optional, Tuple
from dataclasses import dataclass, field


RISK_FREE_RATE = 0.045  # Annual risk-free rate
TRADING_DAYS = 252


@dataclass
class PortfolioRiskMetrics:
    """Comprehensive risk metrics for the portfolio."""
    # Volatility
    daily_volatility: float = 0.0
    annual_volatility: float = 0.0
    # Returns
    total_return: float = 0.0
    annualized_return: float = 0.0
    # Risk-adjusted
    sharpe_ratio: float = 0.0
    sortino_ratio: float = 0.0
    # VaR / CVaR
    var_95: float = 0.0
    var_99: float = 0.0
    cvar_95: float = 0.0
    # Drawdown
    max_drawdown: float = 0.0
    current_drawdown: float = 0.0
    # Beta
    portfolio_beta: float = 0.0


@dataclass
class StressScenario:
    """A single stress test scenario."""
    name: str
    market_shock: float  # e.g., -0.20 for -20%
    portfolio_impact: float = 0.0
    stressed_nav: float = 0.0
    position_impacts: Dict[str, float] = field(default_factory=dict)


def calculate_returns(prices: List[float]) -> np.ndarray:
    """Calculate daily returns from price series."""
    arr = np.array(prices)
    if len(arr) < 2:
        return np.array([])
    returns = np.diff(arr) / arr[:-1]
    # Replace inf/nan with 0
    returns = np.where(np.isfinite(returns), returns, 0.0)
    return returns


def calculate_portfolio_returns(
    daily_units: Dict[str, List[float]],
    prices: Dict[str, List[float]],
    daily_cash: List[float],
) -> Tuple[np.ndarray, np.ndarray]:
    """
    Calculate daily portfolio NAV and returns.

    Args:
        daily_units: Dict of ticker -> daily unit counts
        prices: Dict of ticker -> daily prices
        daily_cash: Daily cash balances

    Returns:
        Tuple of (nav_series, return_series) as numpy arrays
    """
    n_days = len(daily_cash)
    if n_days == 0:
        return np.array([]), np.array([])

    nav = np.zeros(n_days)

    for i in range(n_days):
        invested = 0.0
        for ticker, units in daily_units.items():
            if ticker in prices and i < len(units) and i < len(prices[ticker]):
                invested += units[i] * prices[ticker][i]
        nav[i] = invested + daily_cash[i]

    # Calculate returns
    returns = np.zeros(n_days)
    for i in range(1, n_days):
        if nav[i - 1] > 0:
            returns[i] = (nav[i] - nav[i - 1]) / nav[i - 1]

    return nav, returns


def calculate_risk_metrics(
    returns: np.ndarray,
    benchmark_returns: Optional[np.ndarray] = None,
    risk_free_rate: float = RISK_FREE_RATE,
) -> PortfolioRiskMetrics:
    """
    Calculate comprehensive risk metrics from daily returns.

    Args:
        returns: Daily portfolio returns (including leading zero)
        benchmark_returns: Daily benchmark returns (SPY)
        risk_free_rate: Annual risk-free rate

    Returns:
        PortfolioRiskMetrics
    """
    metrics = PortfolioRiskMetrics()

    # Strip leading zero if present
    if len(returns) > 1 and returns[0] == 0:
        returns = returns[1:]

    if len(returns) < 20:
        return metrics

    daily_rf = risk_free_rate / TRADING_DAYS

    # Volatility
    metrics.daily_volatility = float(np.std(returns, ddof=1))
    metrics.annual_volatility = metrics.daily_volatility * np.sqrt(TRADING_DAYS)

    # Returns
    cumulative = np.prod(1 + returns) - 1
    metrics.total_return = float(cumulative)
    n_years = len(returns) / TRADING_DAYS
    if n_years > 0 and (1 + cumulative) > 0:
        metrics.annualized_return = float((1 + cumulative) ** (1 / n_years) - 1)

    # Sharpe Ratio
    excess_return = metrics.annualized_return - risk_free_rate
    if metrics.annual_volatility > 0:
        metrics.sharpe_ratio = float(excess_return / metrics.annual_volatility)

    # Sortino Ratio (using downside deviation)
    downside = returns[returns < daily_rf]
    if len(downside) > 0:
        downside_dev = np.sqrt(np.mean((downside - daily_rf) ** 2)) * np.sqrt(TRADING_DAYS)
        if downside_dev > 0:
            metrics.sortino_ratio = float(excess_return / downside_dev)

    # VaR
    metrics.var_95 = float(np.percentile(returns, 5))
    metrics.var_99 = float(np.percentile(returns, 1))

    # CVaR (Expected Shortfall)
    tail = returns[returns <= metrics.var_95]
    if len(tail) > 0:
        metrics.cvar_95 = float(np.mean(tail))

    # Drawdown
    cumulative_returns = np.cumprod(1 + returns)
    running_max = np.maximum.accumulate(cumulative_returns)
    drawdowns = cumulative_returns / running_max - 1
    metrics.max_drawdown = float(np.min(drawdowns))
    metrics.current_drawdown = float(drawdowns[-1])

    # Beta vs benchmark
    if benchmark_returns is not None:
        bench = benchmark_returns
        if len(bench) > 1 and bench[0] == 0:
            bench = bench[1:]
        min_len = min(len(returns), len(bench))
        if min_len > 20:
            r = returns[:min_len]
            b = bench[:min_len]
            cov = np.cov(r, b)
            if cov[1, 1] > 0:
                metrics.portfolio_beta = float(cov[0, 1] / cov[1, 1])

    return metrics


def calculate_position_betas(
    ticker_returns: Dict[str, np.ndarray],
    benchmark_returns: np.ndarray,
) -> Dict[str, float]:
    """Calculate individual betas for each position."""
    betas = {}

    if len(benchmark_returns) < 20:
        return betas

    bench = benchmark_returns
    if len(bench) > 1 and bench[0] == 0:
        bench = bench[1:]

    for ticker, returns in ticker_returns.items():
        r = returns
        if len(r) > 1 and r[0] == 0:
            r = r[1:]
        min_len = min(len(r), len(bench))
        if min_len > 20:
            cov = np.cov(r[:min_len], bench[:min_len])
            if cov[1, 1] > 0:
                betas[ticker] = float(cov[0, 1] / cov[1, 1])
            else:
                betas[ticker] = 0.0
        else:
            betas[ticker] = 0.0

    return betas


def run_stress_tests(
    nav: float,
    position_values: Dict[str, float],
    position_betas: Dict[str, float],
    scenarios: Optional[List[Dict]] = None,
) -> List[StressScenario]:
    """
    Run stress test scenarios.

    Default scenarios: -20%, -10%, -5%, +10% market moves.
    Uses individual betas for position-level impact.

    Args:
        nav: Current portfolio NAV
        position_values: Dict of ticker -> current market value
        position_betas: Dict of ticker -> beta
        scenarios: Optional custom scenarios

    Returns:
        List of StressScenario results
    """
    if scenarios is None:
        scenarios = [
            {"name": "Severe Bear (-20%)", "shock": -0.20},
            {"name": "Moderate Bear (-10%)", "shock": -0.10},
            {"name": "Mild Correction (-5%)", "shock": -0.05},
            {"name": "Bull Rally (+10%)", "shock": 0.10},
        ]

    results = []

    for s in scenarios:
        scenario = StressScenario(
            name=s["name"],
            market_shock=s["shock"],
        )

        total_impact = 0.0
        for ticker, value in position_values.items():
            beta = position_betas.get(ticker, 1.0)
            position_impact = value * s["shock"] * beta
            scenario.position_impacts[ticker] = position_impact
            total_impact += position_impact

        scenario.portfolio_impact = total_impact
        scenario.stressed_nav = nav + total_impact
        results.append(scenario)

    return results
