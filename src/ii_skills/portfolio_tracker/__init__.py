"""
Portfolio Tracker Skill

Institutional-grade portfolio tracking workbook with Bloomberg Terminal
styling, formula-driven analytics, and 3-mode update system.

Capabilities:
- Generate comprehensive Excel workbook (14 sheets)
- Full sync (refresh all data + regenerate)
- Quick update (append latest prices)
- Manual trade entry
- Update investment theses
- Export risk reports
- Query positions and risk metrics
- Configure benchmark and allocation targets

Version: 1.0.0
"""

from typing import Dict, List, Optional
from pathlib import Path

from ii_skills import BaseSkill, register_skill

__version__ = "1.0.0"
__author__ = "II-Agent System"


@register_skill
class PortfolioTrackerSkill(BaseSkill):
    """Portfolio Tracker skill for ii-agent."""

    name = "portfolio_tracker"
    version = __version__
    description = "Institutional-grade portfolio tracking with formula-driven Excel workbooks"

    SKILL_DIR = Path(__file__).parent
    OUTPUT_DIR = SKILL_DIR / "output"

    def __init__(self, config: Optional[Dict] = None):
        super().__init__(config)
        self.OUTPUT_DIR.mkdir(exist_ok=True)
        # Defaults
        self._benchmark = "SPY"
        self._risk_free_rate = 0.045
        self._target_allocation = {
            "Equities": 0.50,
            "Commodities": 0.15,
            "Crypto": 0.10,
            "ETF": 0.15,
            "Cash": 0.10,
        }

    def validate_config(self) -> List[str]:
        """Validate configuration. Yahoo Finance doesn't require API keys."""
        return []

    def get_capabilities(self) -> List[str]:
        """Return list of skill capabilities."""
        return [
            "generate_workbook",
            "full_sync",
            "quick_update",
            "add_trade",
            "update_theses",
            "export_risk_report",
            "list_positions",
            "get_risk_metrics",
            "configure",
        ]

    def execute(self, action: str, **kwargs) -> Dict:
        """Execute a portfolio tracker action."""
        if action == "generate_workbook":
            return self._generate_workbook(**kwargs)
        elif action == "full_sync":
            return self._full_sync(**kwargs)
        elif action == "quick_update":
            return self._quick_update(**kwargs)
        elif action == "add_trade":
            return self._add_trade(**kwargs)
        elif action == "update_theses":
            return self._update_theses(**kwargs)
        elif action == "export_risk_report":
            return self._export_risk_report(**kwargs)
        elif action == "list_positions":
            return self._list_positions(**kwargs)
        elif action == "get_risk_metrics":
            return self._get_risk_metrics(**kwargs)
        elif action == "configure":
            return self._configure(**kwargs)
        else:
            raise NotImplementedError(f"Action '{action}' not implemented")

    def _generate_workbook(self, output_name: Optional[str] = None,
                           period: str = "1y", **kwargs) -> Dict:
        """Generate full portfolio tracker workbook."""
        from .workbook_generator import PortfolioWorkbookGenerator

        generator = PortfolioWorkbookGenerator(
            output_dir=str(self.OUTPUT_DIR),
            benchmark=self._benchmark,
            risk_free_rate=self._risk_free_rate,
            target_allocation=self._target_allocation,
        )

        output_path = generator.generate(
            output_name=output_name,
            period=period,
        )

        return {
            "success": True,
            "output_path": output_path,
            "sheets": generator.sheet_names,
        }

    def _full_sync(self, **kwargs) -> Dict:
        """Full sync: refresh all data and regenerate workbook."""
        from .updater import PortfolioUpdater

        updater = PortfolioUpdater(
            output_dir=str(self.OUTPUT_DIR),
            benchmark=self._benchmark,
            risk_free_rate=self._risk_free_rate,
            target_allocation=self._target_allocation,
        )
        return updater.full_sync(**kwargs)

    def _quick_update(self, workbook_path: Optional[str] = None, **kwargs) -> Dict:
        """Quick update: append latest prices to existing workbook."""
        from .updater import PortfolioUpdater

        updater = PortfolioUpdater(
            output_dir=str(self.OUTPUT_DIR),
            benchmark=self._benchmark,
            risk_free_rate=self._risk_free_rate,
            target_allocation=self._target_allocation,
        )
        return updater.quick_update(workbook_path=workbook_path, **kwargs)

    def _add_trade(self, ticker: str, action: str, quantity: float,
                   price: float, currency: str = "CAD",
                   account: str = "", fees: float = 0.0,
                   date: Optional[str] = None, **kwargs) -> Dict:
        """Add a manual trade entry."""
        from .updater import PortfolioUpdater

        updater = PortfolioUpdater(
            output_dir=str(self.OUTPUT_DIR),
            benchmark=self._benchmark,
            risk_free_rate=self._risk_free_rate,
            target_allocation=self._target_allocation,
        )
        return updater.add_trade(
            ticker=ticker, action=action, quantity=quantity,
            price=price, currency=currency, account=account,
            fees=fees, date=date, **kwargs
        )

    def _update_theses(self, ticker: str, thesis: str = "",
                       catalyst: str = "", edge: str = "",
                       workbook_path: Optional[str] = None, **kwargs) -> Dict:
        """Update investment thesis for a ticker."""
        from .updater import PortfolioUpdater

        updater = PortfolioUpdater(
            output_dir=str(self.OUTPUT_DIR),
            benchmark=self._benchmark,
            risk_free_rate=self._risk_free_rate,
            target_allocation=self._target_allocation,
        )
        return updater.update_theses(
            ticker=ticker, thesis=thesis, catalyst=catalyst,
            edge=edge, workbook_path=workbook_path, **kwargs
        )

    def _export_risk_report(self, **kwargs) -> Dict:
        """Export standalone risk report."""
        from .workbook_generator import PortfolioWorkbookGenerator
        from .risk_engine import PortfolioRiskMetrics

        generator = PortfolioWorkbookGenerator(
            output_dir=str(self.OUTPUT_DIR),
            benchmark=self._benchmark,
            risk_free_rate=self._risk_free_rate,
            target_allocation=self._target_allocation,
        )

        # Load data and compute metrics
        metrics, nav, positions = generator.compute_risk_data()

        return {
            "success": True,
            "metrics": {
                "annual_volatility": metrics.annual_volatility,
                "sharpe_ratio": metrics.sharpe_ratio,
                "sortino_ratio": metrics.sortino_ratio,
                "var_95": metrics.var_95,
                "var_99": metrics.var_99,
                "cvar_95": metrics.cvar_95,
                "max_drawdown": metrics.max_drawdown,
                "portfolio_beta": metrics.portfolio_beta,
                "total_return": metrics.total_return,
            },
            "nav": nav,
            "positions": positions,
        }

    def _list_positions(self, **kwargs) -> Dict:
        """Return current positions summary."""
        from .workbook_generator import PortfolioWorkbookGenerator

        generator = PortfolioWorkbookGenerator(
            output_dir=str(self.OUTPUT_DIR),
            benchmark=self._benchmark,
            risk_free_rate=self._risk_free_rate,
            target_allocation=self._target_allocation,
        )
        positions = generator.get_positions_summary()
        return {"success": True, "positions": positions}

    def _get_risk_metrics(self, **kwargs) -> Dict:
        """Return computed risk metrics."""
        return self._export_risk_report(**kwargs)

    def _configure(self, benchmark: Optional[str] = None,
                   risk_free_rate: Optional[float] = None,
                   target_allocation: Optional[Dict[str, float]] = None,
                   **kwargs) -> Dict:
        """Configure benchmark, risk-free rate, or target allocation."""
        if benchmark is not None:
            self._benchmark = benchmark
        if risk_free_rate is not None:
            self._risk_free_rate = risk_free_rate
        if target_allocation is not None:
            self._target_allocation = target_allocation

        return {
            "success": True,
            "benchmark": self._benchmark,
            "risk_free_rate": self._risk_free_rate,
            "target_allocation": self._target_allocation,
        }


def get_tracker(config: Optional[Dict] = None) -> PortfolioTrackerSkill:
    """Get an instance of the Portfolio Tracker skill."""
    skill = PortfolioTrackerSkill(config)
    skill.initialize()
    return skill
