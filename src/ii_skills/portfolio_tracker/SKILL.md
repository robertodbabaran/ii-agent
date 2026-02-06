# Portfolio Tracker Skill

Institutional-grade portfolio tracking workbook with Bloomberg Terminal styling, formula-driven analytics, and a 3-mode update system.

## Overview

Generates a comprehensive 14-sheet Excel workbook from your `networth_config.py` holdings. All analytics sheets use **Excel formulas** (not hardcoded values) so the workbook auto-recalculates when prices change.

## Actions (9)

| Action | Description |
|--------|-------------|
| `generate_workbook` | Full workbook generation (all 14 sheets) |
| `full_sync` | Refresh trades + prices + reference data + regenerate |
| `quick_update` | Append latest prices to Market_Data in existing workbook |
| `add_trade` | Manual trade entry → recompute units/cash → regenerate |
| `update_theses` | Update Investment Theses text for a ticker |
| `export_risk_report` | Standalone risk metrics as dict |
| `list_positions` | Current positions summary |
| `get_risk_metrics` | Computed risk metrics |
| `configure` | Set benchmark, risk-free rate, target allocation |

## Workbook Architecture (14 Sheets)

### Data Lake (Python-populated)

| Sheet | Content |
|-------|---------|
| **Trade_Log** | Date, Ticker, Action, Qty, Price, Currency, Fees, Account |
| **Market_Data** | Rows=dates, Cols=tickers + USDCAD + SPY (1yr history) |
| **Reference_Data** | Ticker, Company, Sector, Industry, Country, Beta, MarketCap |
| **Daily_Units** | Cumulative shares held per ticker per date |
| **Daily_Cash** | Daily cash balance |

### Analytics (Excel Formulas)

| Sheet | Key Formulas |
|-------|-------------|
| **Positions** | Shares=`Daily_Units!last`, Price=`Market_Data!last`, Value=`shares*price`, Weight=`value/total` |
| **Equity_Curve** | Invested=`SUM(units*prices)`, NAV=`invested+cash`, Return=`NAV/prev-1`, Drawdown=`NAV/MAX-1` |
| **Sector_Exposure** | `SUMIFS(MV, Sector, "name")` + `COUNTIFS` |
| **Geographic_Exposure** | `SUMPRODUCT((country="X")*MV)` |
| **Investment_Theses** | Weight via `INDEX/MATCH`, text fields for thesis/catalyst/edge |

### Risk (Excel Formulas + Python computed)

| Sheet | Key Formulas |
|-------|-------------|
| **Correlation** | Helper return columns + `CORREL()` matrix with 3-color conditional formatting |
| **Risk_Metrics** | `STDEV*SQRT(252)`, `PERCENTILE.INC`, `AVERAGEIFS` for CVaR, Sharpe, Sortino |
| **Stress_Testing** | 4 scenarios (-20%, -10%, -5%, +10%) with beta-adjusted position impacts |
| **Dashboard** | KPI cells referencing other sheets + NAV line chart + sector pie chart |

## Update Modes

### 1. Full Sync (`full_sync`)
Refreshes all data from yfinance and regenerates the entire workbook. Use for weekly/monthly refresh.

### 2. Quick Update (`quick_update`)
Opens existing workbook, appends one row to Market_Data with latest prices. All formula sheets auto-recalculate. Use for daily price updates.

### 3. Add Trade (`add_trade`)
Adds a trade to Trade_Log, recomputes Daily_Units/Daily_Cash, and regenerates. Use when you buy/sell a position.

## Bloomberg Terminal Styling

- **Font**: Segoe UI, 10pt headers / 9pt data
- **Colors**: Navy headers (#1C2541), dark background, green/red for +/-
- **Zoom**: 85%, gridlines hidden, freeze panes at B2
- **Number formats**: `#,##0.00` currency, `0.0%` percentages

## Data Sources

- **Holdings**: `networth_newsletter/config.py` (restored from vault by runner .bat)
- **Prices**: yfinance (batch download, 1yr history)
- **Reference**: yfinance Ticker.info (sector, industry, beta, market cap)
- **FX**: USDCAD=X via yfinance

## Configuration

```python
skill.execute("configure",
    benchmark="SPY",
    risk_free_rate=0.045,
    target_allocation={"Equities": 0.50, "Commodities": 0.15, "Crypto": 0.10}
)
```

## File Structure

```
portfolio_tracker/
├── __init__.py              # PortfolioTrackerSkill (9 actions)
├── workbook_generator.py    # 14 sheet methods + generate orchestrator
├── data_fetcher.py          # yfinance data fetching
├── risk_engine.py           # Risk metrics computation
├── trade_log.py             # Trade synthesis + daily units/cash
├── updater.py               # 3-mode updater
├── formula_builder.py       # Extended FormulaBuilder
├── styling.py               # Bloomberg Terminal constants
├── run_portfolio_tracker.bat # Windows runner
└── output/                  # Generated workbooks
```
