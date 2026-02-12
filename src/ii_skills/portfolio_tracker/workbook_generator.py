#!/usr/bin/env python3
"""
Portfolio Workbook Generator — 14 sheet methods + Bloomberg styling.

Phases 2-4 combined:
  Data Lake (5 sheets):   Trade_Log, Market_Data, Reference_Data, Daily_Units, Daily_Cash
  Analytics (5 sheets):   Positions, Equity_Curve, Sector_Exposure, Geographic_Exposure, Investment_Theses
  Risk (4 sheets):        Correlation, Risk_Metrics, Stress_Testing, Dashboard

All analytics/risk sheets use Excel FORMULAS referencing the data lake sheets
so the workbook auto-recalculates when prices change.
"""

import os
from datetime import datetime
from typing import Dict, List, Optional, Tuple

import numpy as np
from openpyxl import Workbook
from openpyxl.chart import LineChart, PieChart, Reference as ChartRef
from openpyxl.chart.series import SeriesLabel
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.utils import get_column_letter

from .formula_builder import FormulaBuilder as FB
from .styling import (
    apply_sheet_defaults, style_header_row, style_data_row,
    apply_bloomberg_styling,
    FMT_CURRENCY, FMT_PCT, FMT_PCT_2, FMT_DATE, FMT_NUMBER, FMT_RATIO,
    HEADER_FONT, DATA_FONT, TITLE_FONT, KPI_VALUE_FONT, KPI_LABEL_FONT,
    HEADER_FILL, DARK_BG_FILL, GREEN, RED,
)
from .data_fetcher import (
    fetch_reference_data, fetch_price_history,
    fetch_exchange_rate_history, ReferenceData, PriceHistory,
)
from .trade_log import TradeLog, Trade
from .risk_engine import (
    calculate_returns, calculate_portfolio_returns,
    calculate_risk_metrics, calculate_position_betas, run_stress_tests,
    PortfolioRiskMetrics,
)


def _load_holdings():
    """Load holdings from networth_config.py via the networth_newsletter config."""
    try:
        import importlib.util
        # Try the networth_newsletter config (restored by runner .bat)
        config_path = os.path.join(
            os.path.dirname(__file__), "..", "networth_newsletter", "config.py"
        )
        config_path = os.path.normpath(config_path)
        if os.path.exists(config_path):
            spec = importlib.util.spec_from_file_location("nw_config", config_path)
            mod = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(mod)
            holdings_dict = getattr(mod, "MARKET_HOLDINGS", {})
            cash_dict = getattr(mod, "CASH_ACCOUNTS", {})
            # Convert dict-of-dicts to list-of-dicts, adding name key
            holdings = []
            for name, cfg in holdings_dict.items():
                h = dict(cfg)
                h["name"] = name
                holdings.append(h)
            cash_list = []
            for name, cfg in cash_dict.items():
                c = dict(cfg)
                c["name"] = name
                cash_list.append(c)
            return holdings, cash_list
    except Exception as e:
        print(f"Warning: Could not load holdings config: {e}")
    return [], []


class PortfolioWorkbookGenerator:
    """
    Generates a 14-sheet portfolio tracker Excel workbook.

    Usage:
        gen = PortfolioWorkbookGenerator(output_dir="output/")
        path = gen.generate()
    """

    def __init__(
        self,
        output_dir: str = "output",
        benchmark: str = "SPY",
        risk_free_rate: float = 0.045,
        target_allocation: Optional[Dict[str, float]] = None,
    ):
        self.output_dir = output_dir
        self.benchmark = benchmark
        self.risk_free_rate = risk_free_rate
        self.target_allocation = target_allocation or {}
        self.wb: Optional[Workbook] = None

        # cell_map tracks row/col positions across sheets for formula references
        self.cell_map: Dict = {}
        self.sheet_names: List[str] = []

        # Cached data (populated during generate)
        self._trade_log: Optional[TradeLog] = None
        self._price_history: Optional[PriceHistory] = None
        self._ref_data: Dict[str, ReferenceData] = {}
        self._daily_units: Dict[str, List[float]] = {}
        self._daily_cash: List[float] = []
        self._tickers: List[str] = []

    # ================================================================
    # PUBLIC API
    # ================================================================

    def generate(self, output_name: Optional[str] = None, period: str = "1y") -> str:
        """Generate the full workbook. Returns output file path."""
        self.wb = Workbook()
        # Remove default sheet
        self.wb.remove(self.wb.active)

        # 1. Load data
        self._load_all_data(period)

        # 2. Data Lake sheets (Phase 2)
        self.add_trade_log()
        self.add_market_data()
        self.add_reference_data()
        self.add_daily_units()
        self.add_daily_cash()

        # 3. Analytics sheets (Phase 3)
        self.add_positions_sheet()
        self.add_equity_curve()
        self.add_sector_exposure()
        self.add_geographic_exposure()
        self.add_investment_theses()

        # 4. Risk sheets (Phase 4)
        self.add_correlation_matrix()
        self.add_risk_metrics()
        self.add_stress_testing()
        self.add_dashboard()

        # 5. Apply Bloomberg styling
        apply_bloomberg_styling(self.wb)

        # Save
        if output_name is None:
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_name = f"Portfolio_Tracker_{ts}.xlsx"
        os.makedirs(self.output_dir, exist_ok=True)
        output_path = os.path.join(self.output_dir, output_name)
        self.wb.save(output_path)
        print(f"Workbook saved to: {output_path}")
        return output_path

    def compute_risk_data(self) -> Tuple[PortfolioRiskMetrics, float, Dict]:
        """Load data and compute risk metrics (for export_risk_report)."""
        self._load_all_data("1y")
        dates = self._price_history.dates if self._price_history else []
        nav_arr, ret_arr = calculate_portfolio_returns(
            self._daily_units, self._price_history.prices, self._daily_cash,
        )
        bench_prices = self._price_history.prices.get(self.benchmark, [])
        bench_ret = calculate_returns(bench_prices) if bench_prices else np.array([])
        metrics = calculate_risk_metrics(ret_arr, bench_ret, self.risk_free_rate)
        nav = float(nav_arr[-1]) if len(nav_arr) > 0 else 0.0
        # Build positions dict
        positions = {}
        for t in self._tickers:
            if t == self.benchmark:
                continue
            units = self._daily_units.get(t, [])
            prices = self._price_history.prices.get(t, [])
            qty = units[-1] if units else 0
            price = prices[-1] if prices else 0
            positions[t] = {"quantity": qty, "price": price, "value": qty * price}
        return metrics, nav, positions

    def get_positions_summary(self) -> Dict:
        """Load data and return positions summary."""
        _, _, positions = self.compute_risk_data()
        return positions

    # ================================================================
    # DATA LOADING
    # ================================================================

    def _load_all_data(self, period: str):
        """Load holdings, prices, reference data, and compute daily series."""
        holdings, cash_accounts = _load_holdings()
        self._trade_log = TradeLog.from_holdings(holdings, cash_accounts)
        self._tickers = self._trade_log.get_tickers()

        # Fetch market data
        all_tickers = list(set(self._tickers + [self.benchmark]))
        self._price_history = fetch_price_history(all_tickers, period)
        self._ref_data = fetch_reference_data(self._tickers)

        # Compute daily units and cash
        dates = self._price_history.dates
        self._daily_units = self._trade_log.compute_daily_units(dates)
        self._daily_cash = self._trade_log.compute_daily_cash(dates)

    # ================================================================
    # PHASE 2 — DATA LAKE SHEETS
    # ================================================================

    def add_trade_log(self):
        """Sheet 1: Trade_Log — raw trade records."""
        ws = self.wb.create_sheet("Trade_Log")
        self.sheet_names.append("Trade_Log")

        headers = ["Date", "Ticker", "Action", "Quantity", "Price",
                    "Currency", "Fees", "Account", "Notes"]
        for col, h in enumerate(headers, 1):
            ws.cell(row=1, column=col, value=h)

        rows = self._trade_log.to_rows()
        for r_idx, row in enumerate(rows, 2):
            for c_idx, val in enumerate(row, 1):
                cell = ws.cell(row=r_idx, column=c_idx, value=val)
                if c_idx == 1:
                    cell.number_format = FMT_DATE
                elif c_idx in (4, 5, 7):
                    cell.number_format = FMT_CURRENCY

        last_row = 1 + len(rows)
        self.cell_map["Trade_Log"] = {
            "headers": headers,
            "data_start_row": 2,
            "last_row": last_row,
            "col_map": {h: i + 1 for i, h in enumerate(headers)},
        }

    def add_market_data(self):
        """Sheet 2: Market_Data — historical prices (rows=dates, cols=tickers+USDCAD+SPY)."""
        ws = self.wb.create_sheet("Market_Data")
        self.sheet_names.append("Market_Data")

        # Build ordered columns: Date, then each ticker, then USDCAD, SPY
        price_tickers = [t for t in self._tickers if t != self.benchmark]
        cols = ["Date"] + price_tickers + ["USDCAD", self.benchmark]

        for col, h in enumerate(cols, 1):
            ws.cell(row=1, column=col, value=h)

        dates = self._price_history.dates
        prices = self._price_history.prices

        # Fetch USDCAD separately if not in prices
        usdcad_prices = prices.get("USDCAD=X", [1.35] * len(dates))
        if len(usdcad_prices) < len(dates):
            usdcad_prices = usdcad_prices + [1.35] * (len(dates) - len(usdcad_prices))

        for r_idx, date in enumerate(dates):
            row = r_idx + 2
            ws.cell(row=row, column=1, value=date).number_format = FMT_DATE

            for c_idx, ticker in enumerate(price_tickers, 2):
                tp = prices.get(ticker, [])
                val = tp[r_idx] if r_idx < len(tp) else 0
                ws.cell(row=row, column=c_idx, value=val).number_format = FMT_CURRENCY

            # USDCAD column
            usdcad_col = len(price_tickers) + 2
            usdcad_val = usdcad_prices[r_idx] if r_idx < len(usdcad_prices) else 1.35
            ws.cell(row=row, column=usdcad_col, value=usdcad_val).number_format = FMT_CURRENCY

            # Benchmark column
            bench_col = len(price_tickers) + 3
            bp = prices.get(self.benchmark, [])
            bench_val = bp[r_idx] if r_idx < len(bp) else 0
            ws.cell(row=row, column=bench_col, value=bench_val).number_format = FMT_CURRENCY

        last_row = 1 + len(dates)
        ticker_cols = {}
        for c_idx, ticker in enumerate(price_tickers, 2):
            ticker_cols[ticker] = c_idx
        ticker_cols["USDCAD"] = len(price_tickers) + 2
        ticker_cols[self.benchmark] = len(price_tickers) + 3

        self.cell_map["Market_Data"] = {
            "data_start_row": 2,
            "last_row": last_row,
            "date_col": 1,
            "ticker_cols": ticker_cols,
            "all_cols": cols,
            "n_tickers": len(price_tickers),
        }

    def add_reference_data(self):
        """Sheet 3: Reference_Data — ticker metadata."""
        ws = self.wb.create_sheet("Reference_Data")
        self.sheet_names.append("Reference_Data")

        headers = ["Ticker", "Company", "Sector", "Industry", "Country",
                    "Currency", "Beta", "Market Cap", "Asset Class"]
        for col, h in enumerate(headers, 1):
            ws.cell(row=1, column=col, value=h)

        tickers = [t for t in self._tickers if t != self.benchmark]
        for r_idx, ticker in enumerate(tickers, 2):
            ref = self._ref_data.get(ticker, ReferenceData(ticker=ticker))
            ws.cell(row=r_idx, column=1, value=ticker)
            ws.cell(row=r_idx, column=2, value=ref.company_name)
            ws.cell(row=r_idx, column=3, value=ref.sector)
            ws.cell(row=r_idx, column=4, value=ref.industry)
            ws.cell(row=r_idx, column=5, value=ref.country)
            ws.cell(row=r_idx, column=6, value=ref.currency)
            ws.cell(row=r_idx, column=7, value=ref.beta).number_format = '0.00'
            ws.cell(row=r_idx, column=8, value=ref.market_cap).number_format = FMT_NUMBER
            ws.cell(row=r_idx, column=9, value=ref.asset_class)

        last_row = 1 + len(tickers)
        self.cell_map["Reference_Data"] = {
            "data_start_row": 2,
            "last_row": last_row,
            "col_map": {h: i + 1 for i, h in enumerate(headers)},
            "ticker_rows": {t: i + 2 for i, t in enumerate(tickers)},
        }

    def add_daily_units(self):
        """Sheet 4: Daily_Units — cumulative shares held per date per ticker."""
        ws = self.wb.create_sheet("Daily_Units")
        self.sheet_names.append("Daily_Units")

        tickers = [t for t in self._tickers if t != self.benchmark]
        cols = ["Date"] + tickers
        for col, h in enumerate(cols, 1):
            ws.cell(row=1, column=col, value=h)

        dates = self._price_history.dates
        for r_idx, date in enumerate(dates):
            row = r_idx + 2
            ws.cell(row=row, column=1, value=date).number_format = FMT_DATE
            for c_idx, ticker in enumerate(tickers, 2):
                units = self._daily_units.get(ticker, [])
                val = units[r_idx] if r_idx < len(units) else 0
                ws.cell(row=row, column=c_idx, value=val).number_format = '#,##0.0000'

        last_row = 1 + len(dates)
        ticker_cols = {t: i + 2 for i, t in enumerate(tickers)}
        self.cell_map["Daily_Units"] = {
            "data_start_row": 2,
            "last_row": last_row,
            "date_col": 1,
            "ticker_cols": ticker_cols,
        }

    def add_daily_cash(self):
        """Sheet 5: Daily_Cash — daily cash balance."""
        ws = self.wb.create_sheet("Daily_Cash")
        self.sheet_names.append("Daily_Cash")

        headers = ["Date", "Cash Balance"]
        for col, h in enumerate(headers, 1):
            ws.cell(row=1, column=col, value=h)

        dates = self._price_history.dates
        for r_idx, date in enumerate(dates):
            row = r_idx + 2
            ws.cell(row=row, column=1, value=date).number_format = FMT_DATE
            val = self._daily_cash[r_idx] if r_idx < len(self._daily_cash) else 0
            ws.cell(row=row, column=2, value=val).number_format = FMT_CURRENCY

        last_row = 1 + len(dates)
        self.cell_map["Daily_Cash"] = {
            "data_start_row": 2,
            "last_row": last_row,
            "date_col": 1,
            "cash_col": 2,
        }

    # ================================================================
    # PHASE 3 — ANALYTICS SHEETS (ALL FORMULAS)
    # ================================================================

    def add_positions_sheet(self):
        """Sheet 6: Positions — current holdings with formula-driven values.

        Columns: Ticker, Company, Shares, Price, Market Value, Weight,
                 Cost Basis, Gain/Loss, Gain/Loss %, Sector, Currency
        All numeric columns are formulas referencing data lake sheets.
        """
        ws = self.wb.create_sheet("Positions")
        self.sheet_names.append("Positions")

        headers = ["Ticker", "Company", "Shares", "Price", "Market Value",
                    "Weight", "Cost Basis", "Gain/Loss", "Gain/Loss %",
                    "Sector", "Currency"]
        for col, h in enumerate(headers, 1):
            ws.cell(row=1, column=col, value=h)

        tickers = [t for t in self._tickers if t != self.benchmark]
        md = self.cell_map["Market_Data"]
        du = self.cell_map["Daily_Units"]
        ref = self.cell_map["Reference_Data"]
        last_data_row = md["last_row"]

        # We need to know where the total market value will be for Weight calc
        total_mv_row = len(tickers) + 3  # row after last ticker + 1 blank
        mv_col = 5  # Market Value is column E

        for i, ticker in enumerate(tickers):
            row = i + 2
            ref_row = ref["ticker_rows"].get(ticker, 2)

            # A: Ticker (text)
            ws.cell(row=row, column=1, value=ticker)

            # B: Company (formula → Reference_Data)
            ws.cell(row=row, column=2,
                    value=f"={FB.sheet_ref('Reference_Data', ref_row, 2)}")

            # C: Shares = Daily_Units last row for this ticker
            du_col = du["ticker_cols"].get(ticker)
            if du_col:
                ws.cell(row=row, column=3,
                        value=f"={FB.sheet_ref('Daily_Units', last_data_row, du_col)}")
            else:
                ws.cell(row=row, column=3, value=0)

            # D: Price = Market_Data last row for this ticker
            md_col = md["ticker_cols"].get(ticker)
            if md_col:
                ws.cell(row=row, column=4,
                        value=f"={FB.sheet_ref('Market_Data', last_data_row, md_col)}")
            else:
                ws.cell(row=row, column=4, value=0)

            # E: Market Value = Shares * Price
            shares_ref = FB.ref(row, 3)
            price_ref = FB.ref(row, 4)
            ws.cell(row=row, column=5, value=f"={shares_ref}*{price_ref}")

            # F: Weight = Market Value / Total MV (use IFERROR)
            mv_ref = FB.ref(row, 5)
            total_ref = FB.ref(total_mv_row, 5, abs_row=True)
            ws.cell(row=row, column=6,
                    value=FB.iferror(f"{mv_ref}/{total_ref}", '"0"'))

            # G: Cost Basis — hardcoded from trade log (no formula source)
            trades = [t for t in self._trade_log.trades if t.ticker == ticker]
            cost_basis = sum(t.quantity * t.price for t in trades if t.action == "BUY")
            ws.cell(row=row, column=7, value=cost_basis)

            # H: Gain/Loss = Market Value - Cost Basis
            ws.cell(row=row, column=8, value=f"={FB.ref(row, 5)}-{FB.ref(row, 7)}")

            # I: Gain/Loss % = Gain/Loss / Cost Basis
            ws.cell(row=row, column=9,
                    value=FB.iferror(f"{FB.ref(row, 8)}/{FB.ref(row, 7)}", "0"))

            # J: Sector (formula → Reference_Data)
            ws.cell(row=row, column=10,
                    value=f"={FB.sheet_ref('Reference_Data', ref_row, 3)}")

            # K: Currency (formula → Reference_Data)
            ws.cell(row=row, column=11,
                    value=f"={FB.sheet_ref('Reference_Data', ref_row, 6)}")

            # Number formats
            ws.cell(row=row, column=3).number_format = '#,##0.00'
            ws.cell(row=row, column=4).number_format = FMT_CURRENCY
            ws.cell(row=row, column=5).number_format = FMT_CURRENCY
            ws.cell(row=row, column=6).number_format = FMT_PCT_2
            ws.cell(row=row, column=7).number_format = FMT_CURRENCY
            ws.cell(row=row, column=8).number_format = FMT_CURRENCY
            ws.cell(row=row, column=9).number_format = FMT_PCT_2

        # Total row
        last_ticker_row = len(tickers) + 1
        ws.cell(row=total_mv_row, column=1, value="TOTAL")
        ws.cell(row=total_mv_row, column=1).font = HEADER_FONT
        # Total Market Value
        ws.cell(row=total_mv_row, column=5,
                value=FB.sum_range(2, last_ticker_row, 5))
        ws.cell(row=total_mv_row, column=5).number_format = FMT_CURRENCY
        # Total Weight (should = 100%)
        ws.cell(row=total_mv_row, column=6,
                value=FB.sum_range(2, last_ticker_row, 6))
        ws.cell(row=total_mv_row, column=6).number_format = FMT_PCT_2
        # Total Cost Basis
        ws.cell(row=total_mv_row, column=7,
                value=FB.sum_range(2, last_ticker_row, 7))
        ws.cell(row=total_mv_row, column=7).number_format = FMT_CURRENCY
        # Total Gain/Loss
        ws.cell(row=total_mv_row, column=8,
                value=FB.sum_range(2, last_ticker_row, 8))
        ws.cell(row=total_mv_row, column=8).number_format = FMT_CURRENCY

        self.cell_map["Positions"] = {
            "data_start_row": 2,
            "last_row": last_ticker_row,
            "total_row": total_mv_row,
            "col_map": {h: i + 1 for i, h in enumerate(headers)},
            "ticker_rows": {t: i + 2 for i, t in enumerate(tickers)},
            "n_tickers": len(tickers),
        }

    def add_equity_curve(self):
        """Sheet 7: Equity_Curve — NAV, daily returns, drawdown (all formulas).

        Columns: Date, Invested Value (SUMPRODUCT), Cash, NAV, Daily Return, Drawdown
        """
        ws = self.wb.create_sheet("Equity_Curve")
        self.sheet_names.append("Equity_Curve")

        headers = ["Date", "Invested Value", "Cash", "NAV",
                    "Daily Return", "Cumulative Max NAV", "Drawdown"]
        for col, h in enumerate(headers, 1):
            ws.cell(row=1, column=col, value=h)

        md = self.cell_map["Market_Data"]
        du = self.cell_map["Daily_Units"]
        dc = self.cell_map["Daily_Cash"]
        n_dates = md["last_row"] - 1  # number of data rows

        tickers = [t for t in self._tickers if t != self.benchmark]
        # We need SUMPRODUCT across ticker columns in Daily_Units * Market_Data
        # Daily_Units ticker cols and Market_Data ticker cols should align

        for r_idx in range(n_dates):
            row = r_idx + 2
            data_row = r_idx + 2  # same row in data lake sheets

            # A: Date
            ws.cell(row=row, column=1,
                    value=f"={FB.sheet_ref('Market_Data', data_row, 1)}")
            ws.cell(row=row, column=1).number_format = FMT_DATE

            # B: Invested Value = SUMPRODUCT(Daily_Units row, Market_Data row)
            if tickers:
                du_start_col = 2
                du_end_col = 1 + len(tickers)
                # Build SUMPRODUCT across matching columns
                du_range = FB.sheet_range("Daily_Units", data_row, du_start_col, data_row, du_end_col)
                # Market_Data ticker columns may not be contiguous with Daily_Units
                # Build a SUM of individual products instead
                products = []
                for ticker in tickers:
                    du_c = du["ticker_cols"].get(ticker)
                    md_c = md["ticker_cols"].get(ticker)
                    if du_c and md_c:
                        du_ref = FB.sheet_ref("Daily_Units", data_row, du_c)
                        md_ref = FB.sheet_ref("Market_Data", data_row, md_c)
                        products.append(f"{du_ref}*{md_ref}")
                if products:
                    ws.cell(row=row, column=2, value="=" + "+".join(products))
                else:
                    ws.cell(row=row, column=2, value=0)
            else:
                ws.cell(row=row, column=2, value=0)
            ws.cell(row=row, column=2).number_format = FMT_CURRENCY

            # C: Cash
            ws.cell(row=row, column=3,
                    value=f"={FB.sheet_ref('Daily_Cash', data_row, 2)}")
            ws.cell(row=row, column=3).number_format = FMT_CURRENCY

            # D: NAV = Invested + Cash
            ws.cell(row=row, column=4,
                    value=f"={FB.ref(row, 2)}+{FB.ref(row, 3)}")
            ws.cell(row=row, column=4).number_format = FMT_CURRENCY

            # E: Daily Return = (NAV / prev NAV) - 1
            if r_idx == 0:
                ws.cell(row=row, column=5, value=0)
            else:
                prev_nav = FB.ref(row - 1, 4)
                nav_ref = FB.ref(row, 4)
                ws.cell(row=row, column=5,
                        value=FB.iferror(f"{nav_ref}/{prev_nav}-1", "0"))
            ws.cell(row=row, column=5).number_format = FMT_PCT_2

            # F: Cumulative Max NAV = MAX(D$2:D{row})
            nav_range = f"D$2:D{row}"
            ws.cell(row=row, column=6, value=f"=MAX({nav_range})")
            ws.cell(row=row, column=6).number_format = FMT_CURRENCY

            # G: Drawdown = NAV / Max NAV - 1
            ws.cell(row=row, column=7,
                    value=FB.iferror(f"{FB.ref(row, 4)}/{FB.ref(row, 6)}-1", "0"))
            ws.cell(row=row, column=7).number_format = FMT_PCT_2

        last_row = 1 + n_dates
        self.cell_map["Equity_Curve"] = {
            "data_start_row": 2,
            "last_row": last_row,
            "col_map": {h: i + 1 for i, h in enumerate(headers)},
        }

    def add_sector_exposure(self):
        """Sheet 8: Sector_Exposure — SUMIFS-based sector aggregation."""
        ws = self.wb.create_sheet("Sector_Exposure")
        self.sheet_names.append("Sector_Exposure")

        headers = ["Sector", "Market Value", "Weight", "# Positions"]
        for col, h in enumerate(headers, 1):
            ws.cell(row=1, column=col, value=h)

        # Get unique sectors
        sectors = sorted(set(
            ref.sector for ref in self._ref_data.values()
            if ref.sector and ref.sector != "Unknown"
        ))
        if not sectors:
            sectors = ["Unknown"]

        pos = self.cell_map["Positions"]
        n_tickers = pos["n_tickers"]
        mv_range = FB.range_ref(2, 5, n_tickers + 1, 5)  # E2:E{last}
        sector_range = FB.range_ref(2, 10, n_tickers + 1, 10)  # J2:J{last}
        pos_mv_range = f"'Positions'!{mv_range}"
        pos_sector_range = f"'Positions'!{sector_range}"
        total_mv_ref = FB.sheet_ref("Positions", pos["total_row"], 5)

        for i, sector in enumerate(sectors):
            row = i + 2
            ws.cell(row=row, column=1, value=sector)

            # Market Value = SUMIFS
            ws.cell(row=row, column=2,
                    value=FB.sumifs(pos_mv_range, pos_sector_range, f'"{sector}"'))
            ws.cell(row=row, column=2).number_format = FMT_CURRENCY

            # Weight = MV / Total MV
            ws.cell(row=row, column=3,
                    value=FB.iferror(f"{FB.ref(row, 2)}/{total_mv_ref}", "0"))
            ws.cell(row=row, column=3).number_format = FMT_PCT_2

            # # Positions = COUNTIFS
            ws.cell(row=row, column=4,
                    value=FB.countifs(pos_sector_range, f'"{sector}"'))

        # Total row
        total_row = len(sectors) + 2
        ws.cell(row=total_row, column=1, value="TOTAL")
        ws.cell(row=total_row, column=1).font = HEADER_FONT
        ws.cell(row=total_row, column=2,
                value=FB.sum_range(2, total_row - 1, 2))
        ws.cell(row=total_row, column=2).number_format = FMT_CURRENCY
        ws.cell(row=total_row, column=3,
                value=FB.sum_range(2, total_row - 1, 3))
        ws.cell(row=total_row, column=3).number_format = FMT_PCT_2

        self.cell_map["Sector_Exposure"] = {
            "data_start_row": 2,
            "last_row": total_row,
            "sectors": sectors,
        }

    def add_geographic_exposure(self):
        """Sheet 9: Geographic_Exposure — SUMIFS by country."""
        ws = self.wb.create_sheet("Geographic_Exposure")
        self.sheet_names.append("Geographic_Exposure")

        headers = ["Country", "Market Value", "Weight", "# Positions"]
        for col, h in enumerate(headers, 1):
            ws.cell(row=1, column=col, value=h)

        countries = sorted(set(
            ref.country for ref in self._ref_data.values()
            if ref.country and ref.country != "Unknown"
        ))
        if not countries:
            countries = ["Unknown"]

        pos = self.cell_map["Positions"]
        n_tickers = pos["n_tickers"]

        # We need a Country column on Positions — but we used Reference_Data
        # Use INDEX/MATCH to look up country from Reference_Data
        # For SUMIFS, we reference Reference_Data country column directly
        ref_cm = self.cell_map["Reference_Data"]
        ref_mv_approach = True  # use Positions MV with Reference_Data country

        # Build ranges: Positions MV and Reference_Data country (aligned by ticker)
        pos_mv_range = f"'Positions'!{FB.range_ref(2, 5, n_tickers + 1, 5)}"
        ref_country_range = f"'Reference_Data'!{FB.range_ref(2, 5, ref_cm['last_row'], 5)}"
        pos_ticker_range = f"'Positions'!{FB.range_ref(2, 1, n_tickers + 1, 1)}"
        total_mv_ref = FB.sheet_ref("Positions", pos["total_row"], 5)

        # Since Positions and Reference_Data have same ticker order, we can use
        # SUMIFS on Positions MV with Reference_Data country
        # But SUMIFS requires same-size ranges from same sheet — workaround: use Positions col 10 (sector)
        # Better: Positions has no country column. Use SUMPRODUCT instead.
        for i, country in enumerate(countries):
            row = i + 2
            ws.cell(row=row, column=1, value=country)

            # Market Value: SUMPRODUCT approach
            # =SUMPRODUCT((Reference_Data!E2:E{n}="{country}")*(Positions!E2:E{n}))
            ref_cty = f"'Reference_Data'!{FB.range_ref(2, 5, n_tickers + 1, 5)}"
            pos_mv = f"'Positions'!{FB.range_ref(2, 5, n_tickers + 1, 5)}"
            ws.cell(row=row, column=2,
                    value=f'=SUMPRODUCT(({ref_cty}="{country}")*{pos_mv})')
            ws.cell(row=row, column=2).number_format = FMT_CURRENCY

            # Weight
            ws.cell(row=row, column=3,
                    value=FB.iferror(f"{FB.ref(row, 2)}/{total_mv_ref}", "0"))
            ws.cell(row=row, column=3).number_format = FMT_PCT_2

            # Count
            ws.cell(row=row, column=4,
                    value=FB.countifs(ref_cty, f'"{country}"'))

        total_row = len(countries) + 2
        ws.cell(row=total_row, column=1, value="TOTAL")
        ws.cell(row=total_row, column=1).font = HEADER_FONT
        ws.cell(row=total_row, column=2,
                value=FB.sum_range(2, total_row - 1, 2))
        ws.cell(row=total_row, column=2).number_format = FMT_CURRENCY
        ws.cell(row=total_row, column=3,
                value=FB.sum_range(2, total_row - 1, 3))
        ws.cell(row=total_row, column=3).number_format = FMT_PCT_2

        self.cell_map["Geographic_Exposure"] = {
            "data_start_row": 2,
            "last_row": total_row,
            "countries": countries,
        }

    def add_investment_theses(self):
        """Sheet 10: Investment_Theses — ticker, weight (formula), text fields."""
        ws = self.wb.create_sheet("Investment_Theses")
        self.sheet_names.append("Investment_Theses")

        headers = ["Ticker", "Weight", "Thesis", "Catalyst", "Edge"]
        for col, h in enumerate(headers, 1):
            ws.cell(row=1, column=col, value=h)

        tickers = [t for t in self._tickers if t != self.benchmark]
        pos = self.cell_map["Positions"]

        for i, ticker in enumerate(tickers):
            row = i + 2
            ws.cell(row=row, column=1, value=ticker)

            # Weight = lookup from Positions via INDEX/MATCH
            pos_weight_range = f"'Positions'!{FB.range_ref(2, 6, pos['last_row'], 6)}"
            pos_ticker_range = f"'Positions'!{FB.range_ref(2, 1, pos['last_row'], 1)}"
            ws.cell(row=row, column=2,
                    value=FB.iferror(
                        FB.index_match(pos_weight_range, pos_ticker_range, f'"{ticker}"').lstrip("="),
                        "0"
                    ))
            ws.cell(row=row, column=2).number_format = FMT_PCT_2

            # Thesis, Catalyst, Edge — blank text fields for user input
            ws.cell(row=row, column=3, value="")
            ws.cell(row=row, column=4, value="")
            ws.cell(row=row, column=5, value="")

        # Set column widths for text
        ws.column_dimensions["C"].width = 50
        ws.column_dimensions["D"].width = 35
        ws.column_dimensions["E"].width = 35

        self.cell_map["Investment_Theses"] = {
            "data_start_row": 2,
            "last_row": 1 + len(tickers),
            "ticker_rows": {t: i + 2 for i, t in enumerate(tickers)},
        }

    # ================================================================
    # PHASE 4 — RISK + CORRELATION + STRESS + DASHBOARD
    # ================================================================

    def add_correlation_matrix(self):
        """Sheet 11: Correlation — daily-return helper columns + CORREL formulas."""
        ws = self.wb.create_sheet("Correlation")
        self.sheet_names.append("Correlation")

        tickers = [t for t in self._tickers if t != self.benchmark]
        md = self.cell_map["Market_Data"]
        n_dates = md["last_row"] - 1

        if not tickers or n_dates < 20:
            ws.cell(row=1, column=1, value="Insufficient data for correlation matrix")
            self.cell_map["Correlation"] = {}
            return

        # Part 1: Daily return helper columns (hidden section to the right)
        # We'll put the return columns starting at column 1
        # Row 1: header "Return: TICKER"
        # Row 2+: =(price_today - price_yesterday) / price_yesterday
        ret_start_col = len(tickers) + 3  # after correlation matrix
        for c_idx, ticker in enumerate(tickers):
            col = ret_start_col + c_idx
            ws.cell(row=1, column=col, value=f"Ret:{ticker}")
            md_col = md["ticker_cols"].get(ticker)
            if not md_col:
                continue
            for r_idx in range(n_dates):
                row = r_idx + 2
                data_row = r_idx + 2
                if r_idx == 0:
                    ws.cell(row=row, column=col, value=0)
                else:
                    today = FB.sheet_ref("Market_Data", data_row, md_col)
                    yesterday = FB.sheet_ref("Market_Data", data_row - 1, md_col)
                    ws.cell(row=row, column=col,
                            value=FB.iferror(f"{today}/{yesterday}-1", "0"))
                ws.cell(row=row, column=col).number_format = FMT_PCT_2

        # Part 2: Correlation matrix (tickers x tickers)
        # Row 1: blank, then ticker headers
        # Col 1: ticker labels
        for c_idx, ticker in enumerate(tickers):
            ws.cell(row=1, column=c_idx + 2, value=ticker)
        for r_idx, ticker in enumerate(tickers):
            ws.cell(row=r_idx + 2, column=1, value=ticker)

        # CORREL formulas
        for r_idx, t1 in enumerate(tickers):
            for c_idx, t2 in enumerate(tickers):
                row = r_idx + 2
                col = c_idx + 2
                if r_idx == c_idx:
                    ws.cell(row=row, column=col, value=1.0)
                else:
                    r1_col = ret_start_col + r_idx
                    r2_col = ret_start_col + c_idx
                    cl1 = get_column_letter(r1_col)
                    cl2 = get_column_letter(r2_col)
                    range1 = f"{cl1}3:{cl1}{n_dates + 1}"  # skip first return (0)
                    range2 = f"{cl2}3:{cl2}{n_dates + 1}"
                    ws.cell(row=row, column=col, value=FB.correl(range1, range2))
                ws.cell(row=row, column=col).number_format = '0.00'

        # Conditional formatting: 3-color scale on the matrix
        matrix_range = f"B2:{get_column_letter(len(tickers) + 1)}{len(tickers) + 1}"
        ws.conditional_formatting.add(matrix_range, ColorScaleRule(
            start_type="num", start_value=-1, start_color="B81D13",
            mid_type="num", mid_value=0, mid_color="FFFFFF",
            end_type="num", end_value=1, end_color="007A33",
        ))

        self.cell_map["Correlation"] = {
            "matrix_start_row": 2,
            "matrix_start_col": 2,
            "ret_start_col": ret_start_col,
            "n_tickers": len(tickers),
            "n_dates": n_dates,
        }

    def add_risk_metrics(self):
        """Sheet 12: Risk_Metrics — 9 key metrics as formulas referencing Equity_Curve."""
        ws = self.wb.create_sheet("Risk_Metrics")
        self.sheet_names.append("Risk_Metrics")

        ec = self.cell_map.get("Equity_Curve", {})
        data_start = ec.get("data_start_row", 2)
        last_row = ec.get("last_row", 2)

        # Return range on Equity_Curve col 5 (Daily Return), skip row 2 (first = 0)
        ret_range = f"'Equity_Curve'!E3:E{last_row}"
        nav_range = f"'Equity_Curve'!D{data_start}:D{last_row}"
        dd_range = f"'Equity_Curve'!G{data_start}:G{last_row}"

        headers = ["Metric", "Value"]
        ws.cell(row=1, column=1, value="Metric")
        ws.cell(row=1, column=2, value="Value")

        rf_daily = self.risk_free_rate / 252

        metrics = [
            ("Annual Volatility",
             f"=STDEV.S({ret_range})*SQRT(252)",
             FMT_PCT_2),
            ("Total Return",
             f"={FB.sheet_ref('Equity_Curve', last_row, 4)}/{FB.sheet_ref('Equity_Curve', data_start, 4)}-1",
             FMT_PCT_2),
            ("Sharpe Ratio",
             FB.iferror(
                 f"(({FB.sheet_ref('Equity_Curve', last_row, 4)}/{FB.sheet_ref('Equity_Curve', data_start, 4)}-1)-{self.risk_free_rate})/(STDEV.S({ret_range})*SQRT(252))",
                 "0"
             ),
             '0.00'),
            ("Sortino Ratio",
             # Approximate: use STDEV of negative returns only
             FB.iferror(
                 f"(({FB.sheet_ref('Equity_Curve', last_row, 4)}/{FB.sheet_ref('Equity_Curve', data_start, 4)}-1)-{self.risk_free_rate})/(STDEV.S(IF({ret_range}<0,{ret_range}))*SQRT(252))",
                 "0"
             ),
             '0.00'),
            ("VaR (95%)",
             f"=PERCENTILE.INC({ret_range},0.05)",
             FMT_PCT_2),
            ("VaR (99%)",
             f"=PERCENTILE.INC({ret_range},0.01)",
             FMT_PCT_2),
            ("CVaR (95%)",
             # AVERAGEIFS for returns <= VaR95
             f'=AVERAGEIFS({ret_range},{ret_range},"<="&PERCENTILE.INC({ret_range},0.05))',
             FMT_PCT_2),
            ("Max Drawdown",
             f"=MIN({dd_range})",
             FMT_PCT_2),
            ("Portfolio Beta",
             # Computed via Python and placed as value (CORREL needs benchmark returns)
             self._compute_beta_value(),
             '0.00'),
        ]

        for i, (name, formula, fmt) in enumerate(metrics):
            row = i + 2
            ws.cell(row=row, column=1, value=name)
            ws.cell(row=row, column=2, value=formula)
            ws.cell(row=row, column=2).number_format = fmt

        # Note: Sortino uses array formula — mark it
        # openpyxl doesn't natively support CSE array formulas in older format
        # The formula will work in modern Excel (365) with implicit arrays

        ws.column_dimensions["A"].width = 25
        ws.column_dimensions["B"].width = 18

        self.cell_map["Risk_Metrics"] = {
            "data_start_row": 2,
            "last_row": len(metrics) + 1,
            "metric_rows": {m[0]: i + 2 for i, m in enumerate(metrics)},
        }

    def _compute_beta_value(self) -> float:
        """Compute portfolio beta in Python (for the Risk_Metrics sheet)."""
        prices = self._price_history.prices if self._price_history else {}
        bench_prices = prices.get(self.benchmark, [])
        if not bench_prices:
            return 0.0
        bench_ret = calculate_returns(bench_prices)
        nav_arr, ret_arr = calculate_portfolio_returns(
            self._daily_units, prices, self._daily_cash
        )
        metrics = calculate_risk_metrics(ret_arr, bench_ret, self.risk_free_rate)
        return round(metrics.portfolio_beta, 4)

    def add_stress_testing(self):
        """Sheet 13: Stress_Testing — 4 scenarios with beta-adjusted impacts."""
        ws = self.wb.create_sheet("Stress_Testing")
        self.sheet_names.append("Stress_Testing")

        headers = ["Scenario", "Market Shock", "Portfolio Impact", "Stressed NAV"]
        for col, h in enumerate(headers, 1):
            ws.cell(row=1, column=col, value=h)

        # Compute current NAV and position values for stress testing
        prices = self._price_history.prices if self._price_history else {}
        nav_arr, ret_arr = calculate_portfolio_returns(
            self._daily_units, prices, self._daily_cash
        )
        current_nav = float(nav_arr[-1]) if len(nav_arr) > 0 else 0

        tickers = [t for t in self._tickers if t != self.benchmark]
        position_values = {}
        for t in tickers:
            units = self._daily_units.get(t, [])
            tp = prices.get(t, [])
            if units and tp:
                position_values[t] = units[-1] * tp[-1]

        bench_ret = calculate_returns(prices.get(self.benchmark, []))
        ticker_returns = {}
        for t in tickers:
            tp = prices.get(t, [])
            if tp:
                ticker_returns[t] = calculate_returns(tp)
        position_betas = calculate_position_betas(ticker_returns, bench_ret)

        stress_results = run_stress_tests(current_nav, position_values, position_betas)

        # Scenario summary rows
        for i, s in enumerate(stress_results):
            row = i + 2
            ws.cell(row=row, column=1, value=s.name)
            ws.cell(row=row, column=2, value=s.market_shock).number_format = FMT_PCT
            ws.cell(row=row, column=3, value=s.portfolio_impact).number_format = FMT_CURRENCY
            ws.cell(row=row, column=4, value=s.stressed_nav).number_format = FMT_CURRENCY

        # Position-level detail below
        detail_start = len(stress_results) + 3
        ws.cell(row=detail_start, column=1, value="Position Detail")
        ws.cell(row=detail_start, column=1).font = HEADER_FONT

        # Headers: Ticker, Beta, then one column per scenario
        ws.cell(row=detail_start + 1, column=1, value="Ticker")
        ws.cell(row=detail_start + 1, column=2, value="Beta")
        ws.cell(row=detail_start + 1, column=3, value="Market Value")
        for s_idx, s in enumerate(stress_results):
            ws.cell(row=detail_start + 1, column=4 + s_idx, value=s.name)

        for t_idx, ticker in enumerate(tickers):
            row = detail_start + 2 + t_idx
            ws.cell(row=row, column=1, value=ticker)
            ws.cell(row=row, column=2, value=position_betas.get(ticker, 1.0)).number_format = '0.00'
            ws.cell(row=row, column=3, value=position_values.get(ticker, 0)).number_format = FMT_CURRENCY
            for s_idx, s in enumerate(stress_results):
                impact = s.position_impacts.get(ticker, 0)
                ws.cell(row=row, column=4 + s_idx, value=impact).number_format = FMT_CURRENCY

        self.cell_map["Stress_Testing"] = {
            "scenario_rows": {s.name: i + 2 for i, s in enumerate(stress_results)},
            "detail_start": detail_start,
        }

    def add_dashboard(self):
        """Sheet 14: Dashboard — KPI summary cells + charts.

        KPIs reference Positions, Equity_Curve, Risk_Metrics.
        Charts: NAV trend line + allocation pie.
        """
        ws = self.wb.create_sheet("Dashboard")
        self.sheet_names.append("Dashboard")
        # Move Dashboard to be the first sheet
        self.wb.move_sheet("Dashboard", offset=-len(self.wb.sheetnames) + 1)

        pos = self.cell_map.get("Positions", {})
        ec = self.cell_map.get("Equity_Curve", {})
        rm = self.cell_map.get("Risk_Metrics", {})

        # ── Title ──
        ws.cell(row=1, column=1, value="PORTFOLIO DASHBOARD")
        ws.cell(row=1, column=1).font = TITLE_FONT
        ws.merge_cells("A1:H1")

        # ── KPI Row ──
        kpi_row = 3
        kpis = [
            ("Total NAV", f"={FB.sheet_ref('Equity_Curve', ec.get('last_row', 2), 4)}",
             FMT_CURRENCY),
            ("Total Return", f"={FB.sheet_ref('Risk_Metrics', rm.get('metric_rows', {}).get('Total Return', 3), 2)}",
             FMT_PCT_2),
            ("# Positions", f"={pos.get('n_tickers', 0)}",
             FMT_NUMBER),
            ("Sharpe Ratio", f"={FB.sheet_ref('Risk_Metrics', rm.get('metric_rows', {}).get('Sharpe Ratio', 4), 2)}",
             '0.00'),
            ("Max Drawdown", f"={FB.sheet_ref('Risk_Metrics', rm.get('metric_rows', {}).get('Max Drawdown', 9), 2)}",
             FMT_PCT_2),
            ("Ann. Volatility", f"={FB.sheet_ref('Risk_Metrics', rm.get('metric_rows', {}).get('Annual Volatility', 2), 2)}",
             FMT_PCT_2),
        ]

        for i, (label, formula, fmt) in enumerate(kpis):
            col = i * 2 + 1
            ws.cell(row=kpi_row, column=col, value=label)
            ws.cell(row=kpi_row, column=col).font = KPI_LABEL_FONT
            ws.cell(row=kpi_row + 1, column=col, value=formula)
            ws.cell(row=kpi_row + 1, column=col).font = KPI_VALUE_FONT
            ws.cell(row=kpi_row + 1, column=col).number_format = fmt

        # ── NAV Trend Chart ──
        chart_start_row = 6
        ws.cell(row=chart_start_row, column=1, value="NAV Trend")
        ws.cell(row=chart_start_row, column=1).font = HEADER_FONT

        ec_last = ec.get("last_row", 2)
        ec_start = ec.get("data_start_row", 2)

        if ec_last > ec_start + 1:
            nav_chart = LineChart()
            nav_chart.title = "Portfolio NAV"
            nav_chart.y_axis.title = "NAV ($)"
            nav_chart.x_axis.title = "Date"
            nav_chart.style = 2
            nav_chart.width = 32
            nav_chart.height = 15

            # NAV data (column D of Equity_Curve)
            nav_data = ChartRef(
                self.wb["Equity_Curve"],
                min_col=4, max_col=4,
                min_row=1, max_row=ec_last,
            )
            nav_dates = ChartRef(
                self.wb["Equity_Curve"],
                min_col=1, max_col=1,
                min_row=ec_start, max_row=ec_last,
            )
            nav_chart.add_data(nav_data, titles_from_data=True)
            nav_chart.set_categories(nav_dates)

            ws.add_chart(nav_chart, "A7")

        # ── Allocation Pie Chart ──
        pie_row = 23
        ws.cell(row=pie_row, column=1, value="Sector Allocation")
        ws.cell(row=pie_row, column=1).font = HEADER_FONT

        se = self.cell_map.get("Sector_Exposure", {})
        sectors = se.get("sectors", [])
        if sectors:
            pie_chart = PieChart()
            pie_chart.title = "Sector Allocation"
            pie_chart.style = 2
            pie_chart.width = 20
            pie_chart.height = 14

            se_ws = self.wb["Sector_Exposure"]
            pie_data = ChartRef(
                se_ws,
                min_col=2, max_col=2,
                min_row=1, max_row=len(sectors) + 1,
            )
            pie_cats = ChartRef(
                se_ws,
                min_col=1, max_col=1,
                min_row=2, max_row=len(sectors) + 1,
            )
            pie_chart.add_data(pie_data, titles_from_data=True)
            pie_chart.set_categories(pie_cats)

            ws.add_chart(pie_chart, "A24")

        # ── Top Holdings Table ──
        top_row = 23
        top_col = 7
        ws.cell(row=top_row, column=top_col, value="Top Holdings")
        ws.cell(row=top_row, column=top_col).font = HEADER_FONT

        top_headers = ["Ticker", "Value", "Weight"]
        for c, h in enumerate(top_headers):
            ws.cell(row=top_row + 1, column=top_col + c, value=h)
            ws.cell(row=top_row + 1, column=top_col + c).font = HEADER_FONT

        tickers = [t for t in self._tickers if t != self.benchmark]
        # Use LARGE/INDEX to get top 10 by value — simpler: just reference Positions rows
        for i, ticker in enumerate(tickers[:10]):
            row = top_row + 2 + i
            p_row = pos.get("ticker_rows", {}).get(ticker, 2)
            ws.cell(row=row, column=top_col, value=ticker)
            ws.cell(row=row, column=top_col + 1,
                    value=f"={FB.sheet_ref('Positions', p_row, 5)}")
            ws.cell(row=row, column=top_col + 1).number_format = FMT_CURRENCY
            ws.cell(row=row, column=top_col + 2,
                    value=f"={FB.sheet_ref('Positions', p_row, 6)}")
            ws.cell(row=row, column=top_col + 2).number_format = FMT_PCT_2

        self.cell_map["Dashboard"] = {"kpi_row": kpi_row}
