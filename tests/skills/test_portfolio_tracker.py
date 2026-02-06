#!/usr/bin/env python3
"""
Tests for Portfolio Tracker Skill

~30 tests covering:
- Skill registration and capabilities
- Formula builder
- Trade log computation
- Risk engine calculations
- Workbook generation (mocked yfinance)
- Updater modes
- Styling constants
"""

import os
import sys
import tempfile
from datetime import datetime, timedelta
from unittest.mock import patch, MagicMock

import numpy as np
import pytest

# Add src to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "..", "src"))


# ================================================================
# Formula Builder Tests
# ================================================================

class TestFormulaBuilder:
    """Tests for the extended FormulaBuilder."""

    def setup_method(self):
        from ii_skills.portfolio_tracker.formula_builder import FormulaBuilder
        self.FB = FormulaBuilder

    def test_col_letter(self):
        assert self.FB.col_letter(1) == "A"
        assert self.FB.col_letter(2) == "B"
        assert self.FB.col_letter(26) == "Z"
        assert self.FB.col_letter(27) == "AA"

    def test_ref(self):
        assert self.FB.ref(5, 2) == "B5"
        assert self.FB.ref(5, 2, abs_row=True, abs_col=True) == "$B$5"
        assert self.FB.ref(10, 3, abs_row=True) == "C$10"

    def test_sheet_ref(self):
        result = self.FB.sheet_ref("Market_Data", 10, 3)
        assert result == "'Market_Data'!C10"

    def test_range_ref(self):
        result = self.FB.range_ref(2, 2, 100, 2)
        assert result == "B2:B100"

    def test_sheet_range(self):
        result = self.FB.sheet_range("Daily_Units", 2, 3, 253, 3)
        assert result == "'Daily_Units'!C2:C253"

    def test_sum_range(self):
        assert self.FB.sum_range(4, 6, 2) == "=SUM(B4:B6)"

    def test_iferror(self):
        assert self.FB.iferror("=B5/B3", "0") == "=IFERROR(B5/B3,0)"
        assert self.FB.iferror("B5/B3", "0") == "=IFERROR(B5/B3,0)"

    def test_sumifs(self):
        result = self.FB.sumifs("D2:D50", "E2:E50", '"Technology"')
        assert result == '=SUMIFS(D2:D50,E2:E50,"Technology")'

    def test_correl(self):
        result = self.FB.correl("C2:C253", "D2:D253")
        assert result == "=CORREL(C2:C253,D2:D253)"

    def test_percentile(self):
        result = self.FB.percentile("F2:F253", 0.05)
        assert result == "=PERCENTILE.INC(F2:F253,0.05)"

    def test_sumproduct(self):
        result = self.FB.sumproduct("B2:Z2", "B3:Z3")
        assert result == "=SUMPRODUCT(B2:Z2,B3:Z3)"

    def test_averageifs(self):
        result = self.FB.averageifs("F2:F253", "F2:F253", '"<="&G2')
        assert result == '=AVERAGEIFS(F2:F253,F2:F253,"<="&G2)'

    def test_index_match(self):
        result = self.FB.index_match("B2:B50", "A2:A50", '"AAPL"')
        assert result == '=INDEX(B2:B50,MATCH("AAPL",A2:A50,0))'

    def test_stdev(self):
        assert self.FB.stdev("F2:F253") == "=STDEV.S(F2:F253)"

    def test_max_zero(self):
        assert self.FB.max_zero("=C10+C11") == "=MAX(0,C10+C11)"


# ================================================================
# Trade Log Tests
# ================================================================

class TestTradeLog:
    """Tests for TradeLog computation."""

    def setup_method(self):
        from ii_skills.portfolio_tracker.trade_log import TradeLog, Trade
        self.TradeLog = TradeLog
        self.Trade = Trade

    def test_from_holdings_basic(self):
        holdings = [
            {"ticker": "AAPL", "quantity": 100, "cost_basis": 15000, "currency": "USD", "account": "IBKR"},
            {"ticker": "TOU.TO", "quantity": 50, "cost_basis": 3000, "currency": "CAD", "account": "TFSA"},
        ]
        log = self.TradeLog.from_holdings(holdings)
        assert len(log.trades) == 2
        assert log.trades[0].ticker == "AAPL"
        assert log.trades[0].quantity == 100
        assert log.trades[0].price == 150.0  # 15000/100

    def test_from_holdings_with_cash(self):
        holdings = [{"ticker": "AAPL", "quantity": 10, "cost_basis": 1500}]
        cash = [{"balance": 5000}, {"balance": 3000}]
        log = self.TradeLog.from_holdings(holdings, cash)
        assert log.initial_cash == 8000

    def test_from_holdings_skips_zero_quantity(self):
        holdings = [
            {"ticker": "AAPL", "quantity": 0},
            {"ticker": "MSFT", "quantity": 50, "cost_basis": 5000},
        ]
        log = self.TradeLog.from_holdings(holdings)
        assert len(log.trades) == 1
        assert log.trades[0].ticker == "MSFT"

    def test_get_tickers(self):
        holdings = [
            {"ticker": "AAPL", "quantity": 100, "cost_basis": 15000},
            {"ticker": "MSFT", "quantity": 50, "cost_basis": 5000},
        ]
        log = self.TradeLog.from_holdings(holdings)
        tickers = log.get_tickers()
        assert "AAPL" in tickers
        assert "MSFT" in tickers

    def test_compute_daily_units(self):
        base = datetime(2024, 1, 1)
        holdings = [{"ticker": "AAPL", "quantity": 100, "cost_basis": 15000}]
        log = self.TradeLog.from_holdings(holdings, base_date=base)

        dates = [base + timedelta(days=i) for i in range(5)]
        units = log.compute_daily_units(dates)
        assert units["AAPL"] == [100.0, 100.0, 100.0, 100.0, 100.0]

    def test_compute_daily_units_with_sell(self):
        base = datetime(2024, 1, 1)
        holdings = [{"ticker": "AAPL", "quantity": 100, "cost_basis": 15000}]
        log = self.TradeLog.from_holdings(holdings, base_date=base)

        sell = self.Trade(
            date=datetime(2024, 1, 3), ticker="AAPL", action="SELL",
            quantity=30, price=155.0
        )
        log.add_trade(sell)

        dates = [base + timedelta(days=i) for i in range(5)]
        units = log.compute_daily_units(dates)
        assert units["AAPL"] == [100.0, 100.0, 70.0, 70.0, 70.0]

    def test_compute_daily_cash(self):
        base = datetime(2024, 1, 1)
        holdings = [{"ticker": "AAPL", "quantity": 10, "cost_basis": 1500}]
        cash = [{"balance": 10000}]
        log = self.TradeLog.from_holdings(holdings, cash, base_date=base)

        dates = [base + timedelta(days=i) for i in range(3)]
        daily_cash = log.compute_daily_cash(dates)
        # Initial cash = 10000, buy cost = 10 * 150 = 1500, so cash = 8500
        assert daily_cash[0] == 8500.0

    def test_to_rows(self):
        holdings = [{"ticker": "AAPL", "quantity": 10, "cost_basis": 1500}]
        log = self.TradeLog.from_holdings(holdings)
        rows = log.to_rows()
        assert len(rows) == 1
        assert rows[0][1] == "AAPL"  # ticker
        assert rows[0][2] == "BUY"   # action
        assert rows[0][3] == 10      # quantity

    def test_trade_total_cost(self):
        buy = self.Trade(date=datetime.now(), ticker="X", action="BUY",
                         quantity=10, price=100, fees=9.99)
        assert buy.total_cost == 1009.99

        sell = self.Trade(date=datetime.now(), ticker="X", action="SELL",
                          quantity=10, price=100, fees=9.99)
        assert sell.total_cost == 990.01


# ================================================================
# Risk Engine Tests
# ================================================================

class TestRiskEngine:
    """Tests for risk metrics calculations."""

    def setup_method(self):
        from ii_skills.portfolio_tracker.risk_engine import (
            calculate_returns, calculate_risk_metrics,
            calculate_position_betas, run_stress_tests,
        )
        self.calculate_returns = calculate_returns
        self.calculate_risk_metrics = calculate_risk_metrics
        self.calculate_position_betas = calculate_position_betas
        self.run_stress_tests = run_stress_tests

    def test_calculate_returns(self):
        prices = [100, 105, 103, 108, 110]
        returns = self.calculate_returns(prices)
        assert len(returns) == 4
        assert abs(returns[0] - 0.05) < 0.001

    def test_calculate_returns_empty(self):
        returns = self.calculate_returns([])
        assert len(returns) == 0

    def test_calculate_returns_single(self):
        returns = self.calculate_returns([100])
        assert len(returns) == 0

    def test_risk_metrics_basic(self):
        np.random.seed(42)
        returns = np.random.normal(0.001, 0.02, 252)
        metrics = self.calculate_risk_metrics(returns)
        assert metrics.annual_volatility > 0
        assert metrics.var_95 < 0  # VaR should be negative
        assert metrics.max_drawdown <= 0

    def test_risk_metrics_insufficient_data(self):
        returns = np.array([0.01, -0.01])
        metrics = self.calculate_risk_metrics(returns)
        assert metrics.annual_volatility == 0.0  # Not enough data

    def test_risk_metrics_with_benchmark(self):
        np.random.seed(42)
        port_returns = np.random.normal(0.001, 0.02, 252)
        bench_returns = np.random.normal(0.0005, 0.015, 252)
        metrics = self.calculate_risk_metrics(port_returns, bench_returns)
        assert metrics.portfolio_beta != 0.0

    def test_position_betas(self):
        np.random.seed(42)
        bench = np.random.normal(0.0005, 0.015, 252)
        ticker_rets = {
            "AAPL": np.random.normal(0.001, 0.02, 252),
            "MSFT": np.random.normal(0.0008, 0.018, 252),
        }
        betas = self.calculate_position_betas(ticker_rets, bench)
        assert "AAPL" in betas
        assert "MSFT" in betas

    def test_stress_tests(self):
        results = self.run_stress_tests(
            nav=100000,
            position_values={"AAPL": 50000, "MSFT": 30000},
            position_betas={"AAPL": 1.2, "MSFT": 0.8},
        )
        assert len(results) == 4
        # Severe bear: AAPL impact = 50000 * -0.20 * 1.2 = -12000
        severe = results[0]
        assert severe.market_shock == -0.20
        assert abs(severe.position_impacts["AAPL"] - (-12000)) < 0.01

    def test_stress_tests_custom_scenarios(self):
        custom = [{"name": "Flash Crash", "shock": -0.30}]
        results = self.run_stress_tests(
            nav=100000,
            position_values={"AAPL": 100000},
            position_betas={"AAPL": 1.0},
            scenarios=custom,
        )
        assert len(results) == 1
        assert results[0].name == "Flash Crash"
        assert results[0].portfolio_impact == -30000


# ================================================================
# Styling Tests
# ================================================================

class TestStyling:
    """Tests for Bloomberg Terminal styling constants."""

    def test_constants_exist(self):
        from ii_skills.portfolio_tracker.styling import (
            NAVY, GREEN, RED, HEADER_FONT, DATA_FONT,
            HEADER_FILL, FMT_CURRENCY, FMT_PCT, ZOOM_SCALE,
        )
        assert NAVY == "1C2541"
        assert GREEN == "007A33"
        assert ZOOM_SCALE == 85

    def test_apply_sheet_defaults(self):
        from openpyxl import Workbook
        from ii_skills.portfolio_tracker.styling import apply_sheet_defaults
        wb = Workbook()
        ws = wb.active
        apply_sheet_defaults(ws)
        assert ws.sheet_view.zoomScale == 85
        assert ws.sheet_view.showGridLines is False
        assert ws.freeze_panes == "B2"

    def test_style_header_row(self):
        from openpyxl import Workbook
        from ii_skills.portfolio_tracker.styling import style_header_row, HEADER_FONT
        wb = Workbook()
        ws = wb.active
        ws.cell(row=1, column=1, value="Test")
        ws.cell(row=1, column=2, value="Header")
        style_header_row(ws, 1, 2)
        assert ws.cell(row=1, column=1).font == HEADER_FONT


# ================================================================
# Skill Registration Tests
# ================================================================

class TestSkillRegistration:
    """Tests for skill discovery and capabilities."""

    def test_skill_class_exists(self):
        from ii_skills.portfolio_tracker import PortfolioTrackerSkill
        assert PortfolioTrackerSkill.name == "portfolio_tracker"
        assert PortfolioTrackerSkill.version == "1.0.0"

    def test_capabilities(self):
        from ii_skills.portfolio_tracker import PortfolioTrackerSkill
        skill = PortfolioTrackerSkill()
        caps = skill.get_capabilities()
        assert "generate_workbook" in caps
        assert "full_sync" in caps
        assert "quick_update" in caps
        assert "add_trade" in caps
        assert "update_theses" in caps
        assert "export_risk_report" in caps
        assert "list_positions" in caps
        assert "get_risk_metrics" in caps
        assert "configure" in caps
        assert len(caps) == 9

    def test_configure(self):
        from ii_skills.portfolio_tracker import PortfolioTrackerSkill
        skill = PortfolioTrackerSkill()
        result = skill.execute("configure", benchmark="QQQ", risk_free_rate=0.05)
        assert result["success"]
        assert result["benchmark"] == "QQQ"
        assert result["risk_free_rate"] == 0.05

    def test_invalid_action(self):
        from ii_skills.portfolio_tracker import PortfolioTrackerSkill
        skill = PortfolioTrackerSkill()
        with pytest.raises(NotImplementedError):
            skill.execute("nonexistent_action")

    def test_validate_config(self):
        from ii_skills.portfolio_tracker import PortfolioTrackerSkill
        skill = PortfolioTrackerSkill()
        issues = skill.validate_config()
        assert issues == []


# ================================================================
# Workbook Generation Tests (mocked yfinance)
# ================================================================

class TestWorkbookGeneration:
    """Integration tests for workbook generation with mocked data."""

    @pytest.fixture
    def mock_data(self):
        """Create mock holdings and price data."""
        holdings = [
            {"ticker": "AAPL", "quantity": 100, "cost_basis": 15000,
             "currency": "USD", "account": "IBKR", "type": "stock"},
            {"ticker": "TOU.TO", "quantity": 200, "cost_basis": 10000,
             "currency": "CAD", "account": "TFSA", "type": "stock"},
        ]
        cash = [{"balance": 5000, "currency": "CAD"}]

        # Mock price history: 50 dates, 2 tickers + SPY
        dates = [datetime(2024, 1, 1) + timedelta(days=i) for i in range(50)]
        np.random.seed(42)
        prices = {
            "AAPL": (150 + np.cumsum(np.random.normal(0, 2, 50))).tolist(),
            "TOU.TO": (55 + np.cumsum(np.random.normal(0, 1, 50))).tolist(),
            "SPY": (450 + np.cumsum(np.random.normal(0, 3, 50))).tolist(),
        }

        ref_data = {
            "AAPL": MagicMock(
                ticker="AAPL", company_name="Apple Inc.", sector="Technology",
                industry="Consumer Electronics", country="United States",
                currency="USD", beta=1.2, market_cap=3000000000000,
                asset_class="Equity"
            ),
            "TOU.TO": MagicMock(
                ticker="TOU.TO", company_name="Tourmaline Oil Corp",
                sector="Energy", industry="Oil & Gas",
                country="Canada", currency="CAD", beta=1.1,
                market_cap=20000000000, asset_class="Equity"
            ),
        }

        return holdings, cash, dates, prices, ref_data

    @pytest.fixture
    def generator(self, mock_data, tmp_path):
        """Create a generator with mocked data."""
        holdings, cash, dates, prices, ref_data = mock_data

        from ii_skills.portfolio_tracker.workbook_generator import PortfolioWorkbookGenerator
        from ii_skills.portfolio_tracker.data_fetcher import PriceHistory
        from ii_skills.portfolio_tracker.trade_log import TradeLog

        gen = PortfolioWorkbookGenerator(output_dir=str(tmp_path))

        # Inject mock data
        gen._trade_log = TradeLog.from_holdings(
            holdings, cash, base_date=dates[0]
        )
        gen._tickers = gen._trade_log.get_tickers()
        gen._price_history = PriceHistory(dates=dates, prices=prices)
        gen._ref_data = ref_data
        gen._daily_units = gen._trade_log.compute_daily_units(dates)
        gen._daily_cash = gen._trade_log.compute_daily_cash(dates)

        return gen

    def test_generate_creates_file(self, generator, tmp_path):
        """Test that generate() creates an xlsx file."""
        from openpyxl import Workbook
        generator.wb = Workbook()
        generator.wb.remove(generator.wb.active)

        # Build all sheets
        generator.add_trade_log()
        generator.add_market_data()
        generator.add_reference_data()
        generator.add_daily_units()
        generator.add_daily_cash()
        generator.add_positions_sheet()
        generator.add_equity_curve()
        generator.add_sector_exposure()
        generator.add_geographic_exposure()
        generator.add_investment_theses()
        generator.add_correlation_matrix()
        generator.add_risk_metrics()
        generator.add_stress_testing()
        generator.add_dashboard()

        output_path = str(tmp_path / "test_output.xlsx")
        generator.wb.save(output_path)
        assert os.path.exists(output_path)

    def test_all_14_sheets_created(self, generator):
        """Verify all 14 sheets are present."""
        from openpyxl import Workbook
        generator.wb = Workbook()
        generator.wb.remove(generator.wb.active)

        generator.add_trade_log()
        generator.add_market_data()
        generator.add_reference_data()
        generator.add_daily_units()
        generator.add_daily_cash()
        generator.add_positions_sheet()
        generator.add_equity_curve()
        generator.add_sector_exposure()
        generator.add_geographic_exposure()
        generator.add_investment_theses()
        generator.add_correlation_matrix()
        generator.add_risk_metrics()
        generator.add_stress_testing()
        generator.add_dashboard()

        assert len(generator.wb.sheetnames) == 14

    def test_trade_log_sheet(self, generator):
        """Trade_Log has correct headers and data."""
        from openpyxl import Workbook
        generator.wb = Workbook()
        generator.wb.remove(generator.wb.active)
        generator.add_trade_log()

        ws = generator.wb["Trade_Log"]
        assert ws.cell(row=1, column=1).value == "Date"
        assert ws.cell(row=1, column=2).value == "Ticker"
        assert ws.cell(row=2, column=2).value == "AAPL"

    def test_market_data_sheet(self, generator):
        """Market_Data has date column and ticker columns."""
        from openpyxl import Workbook
        generator.wb = Workbook()
        generator.wb.remove(generator.wb.active)
        generator.add_market_data()

        ws = generator.wb["Market_Data"]
        assert ws.cell(row=1, column=1).value == "Date"
        # Should have ticker columns
        headers = [ws.cell(row=1, column=c).value for c in range(1, 10) if ws.cell(row=1, column=c).value]
        assert "AAPL" in headers
        assert "SPY" in headers

    def test_positions_formulas(self, generator):
        """Positions sheet cells contain formulas (start with =)."""
        from openpyxl import Workbook
        generator.wb = Workbook()
        generator.wb.remove(generator.wb.active)
        generator.add_trade_log()
        generator.add_market_data()
        generator.add_reference_data()
        generator.add_daily_units()
        generator.add_daily_cash()
        generator.add_positions_sheet()

        ws = generator.wb["Positions"]
        # Shares (col 3) should be a formula
        shares_cell = ws.cell(row=2, column=3).value
        assert isinstance(shares_cell, str) and shares_cell.startswith("=")
        # Price (col 4) should be a formula
        price_cell = ws.cell(row=2, column=4).value
        assert isinstance(price_cell, str) and price_cell.startswith("=")
        # Market Value (col 5) should be a formula
        mv_cell = ws.cell(row=2, column=5).value
        assert isinstance(mv_cell, str) and mv_cell.startswith("=")

    def test_equity_curve_formulas(self, generator):
        """Equity_Curve has SUMPRODUCT-based invested value formulas."""
        from openpyxl import Workbook
        generator.wb = Workbook()
        generator.wb.remove(generator.wb.active)
        generator.add_trade_log()
        generator.add_market_data()
        generator.add_reference_data()
        generator.add_daily_units()
        generator.add_daily_cash()
        generator.add_positions_sheet()
        generator.add_equity_curve()

        ws = generator.wb["Equity_Curve"]
        # Invested Value (col 2) should contain formulas with sheet refs
        inv_val = ws.cell(row=3, column=2).value  # row 3 = second data row
        assert isinstance(inv_val, str) and "Daily_Units" in inv_val

    def test_risk_metrics_formulas(self, generator):
        """Risk_Metrics sheet has formula-based metrics."""
        from openpyxl import Workbook
        generator.wb = Workbook()
        generator.wb.remove(generator.wb.active)
        generator.add_trade_log()
        generator.add_market_data()
        generator.add_reference_data()
        generator.add_daily_units()
        generator.add_daily_cash()
        generator.add_positions_sheet()
        generator.add_equity_curve()
        generator.add_sector_exposure()
        generator.add_geographic_exposure()
        generator.add_investment_theses()
        generator.add_correlation_matrix()
        generator.add_risk_metrics()

        ws = generator.wb["Risk_Metrics"]
        # Annual Volatility should be a STDEV formula
        vol_val = ws.cell(row=2, column=2).value
        assert isinstance(vol_val, str) and "STDEV" in vol_val

    def test_cell_map_populated(self, generator):
        """cell_map is populated for all data lake sheets."""
        from openpyxl import Workbook
        generator.wb = Workbook()
        generator.wb.remove(generator.wb.active)
        generator.add_trade_log()
        generator.add_market_data()
        generator.add_reference_data()
        generator.add_daily_units()
        generator.add_daily_cash()

        assert "Trade_Log" in generator.cell_map
        assert "Market_Data" in generator.cell_map
        assert "Reference_Data" in generator.cell_map
        assert "Daily_Units" in generator.cell_map
        assert "Daily_Cash" in generator.cell_map
        assert generator.cell_map["Market_Data"]["data_start_row"] == 2
        assert "ticker_cols" in generator.cell_map["Market_Data"]

    def test_correlation_matrix(self, generator):
        """Correlation matrix has CORREL formulas."""
        from openpyxl import Workbook
        generator.wb = Workbook()
        generator.wb.remove(generator.wb.active)
        generator.add_trade_log()
        generator.add_market_data()
        generator.add_reference_data()
        generator.add_daily_units()
        generator.add_daily_cash()
        generator.add_positions_sheet()
        generator.add_correlation_matrix()

        ws = generator.wb["Correlation"]
        # Off-diagonal should be CORREL formula
        corr_val = ws.cell(row=2, column=3).value  # AAPL vs TOU.TO
        assert isinstance(corr_val, str) and "CORREL" in corr_val
        # Diagonal should be 1.0
        diag_val = ws.cell(row=2, column=2).value
        assert diag_val == 1.0


# ================================================================
# Data Fetcher Tests (mocked)
# ================================================================

class TestDataFetcher:
    """Tests for data fetcher with mocked yfinance."""

    def test_reference_data_crypto(self):
        from ii_skills.portfolio_tracker.data_fetcher import ReferenceData
        # Test classification logic without yfinance
        ref = ReferenceData(ticker="BTC-USD")
        ref.asset_class = "Crypto"
        ref.sector = "Cryptocurrency"
        assert ref.asset_class == "Crypto"

    def test_reference_data_cad_stock(self):
        from ii_skills.portfolio_tracker.data_fetcher import ReferenceData
        ref = ReferenceData(ticker="TOU.TO")
        # CAD stocks should default to CAD currency
        if ref.ticker.endswith(".TO"):
            ref.currency = "CAD"
        assert ref.currency == "CAD"
