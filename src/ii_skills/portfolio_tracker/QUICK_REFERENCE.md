# Portfolio Tracker — Quick Reference

## Commands

| I want to... | Action | Example |
|--------------|--------|---------|
| Generate full workbook | `generate_workbook` | `skill.execute("generate_workbook")` |
| Refresh everything | `full_sync` | `skill.execute("full_sync")` |
| Update today's prices | `quick_update` | `skill.execute("quick_update")` |
| Record a trade | `add_trade` | `skill.execute("add_trade", ticker="AAPL", action="BUY", quantity=50, price=180)` |
| Add investment thesis | `update_theses` | `skill.execute("update_theses", ticker="AAPL", thesis="AI monetization leader")` |
| See positions | `list_positions` | `skill.execute("list_positions")` |
| See risk metrics | `get_risk_metrics` | `skill.execute("get_risk_metrics")` |
| Change benchmark | `configure` | `skill.execute("configure", benchmark="QQQ")` |

## Sheet → Data Mapping

| Sheet | Source | Type |
|-------|--------|------|
| Trade_Log | `networth_config.py` holdings | Data |
| Market_Data | yfinance batch download | Data |
| Reference_Data | yfinance Ticker.info | Data |
| Daily_Units | Computed from Trade_Log | Data |
| Daily_Cash | Computed from Trade_Log + CASH_ACCOUNTS | Data |
| Positions | Formulas → Daily_Units, Market_Data, Reference_Data | Formula |
| Equity_Curve | Formulas → Daily_Units, Market_Data, Daily_Cash | Formula |
| Sector_Exposure | SUMIFS → Positions | Formula |
| Geographic_Exposure | SUMPRODUCT → Positions, Reference_Data | Formula |
| Investment_Theses | INDEX/MATCH → Positions | Formula |
| Correlation | CORREL → Market_Data returns | Formula |
| Risk_Metrics | STDEV, PERCENTILE, AVERAGEIFS → Equity_Curve | Formula |
| Stress_Testing | Python-computed beta-adjusted scenarios | Computed |
| Dashboard | Refs → Positions, Equity_Curve, Risk_Metrics + Charts | Formula |

## Risk Metrics Available

| Metric | Formula |
|--------|---------|
| Annual Volatility | `STDEV.S(returns)*SQRT(252)` |
| Total Return | `last_NAV/first_NAV - 1` |
| Sharpe Ratio | `(return - rf) / volatility` |
| Sortino Ratio | `(return - rf) / downside_dev` |
| VaR 95% | `PERCENTILE.INC(returns, 0.05)` |
| VaR 99% | `PERCENTILE.INC(returns, 0.01)` |
| CVaR 95% | `AVERAGEIFS(returns, returns, "<="&VaR95)` |
| Max Drawdown | `MIN(drawdown_series)` |
| Portfolio Beta | Python-computed vs SPY |
