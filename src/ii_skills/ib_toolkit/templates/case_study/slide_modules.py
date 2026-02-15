#!/usr/bin/env python3
"""
Modular Slide Generator for PE/IB Analysis

This module provides individual slide generation functions that can be
called independently based on the type of analysis requested.

Usage:
    from slide_modules import SlideGenerator

    gen = SlideGenerator()
    gen.add_industry_analysis(company_data)
    gen.add_competitive_analysis(company_data)
    gen.save("output.pptx")

Or for single-purpose outputs:
    gen = SlideGenerator()
    gen.add_debt_schedule(debt_data)
    gen.save("debt_analysis.pptx")
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from datetime import datetime
from dataclasses import dataclass, field
from typing import List, Dict, Optional, Any
import os
import math

# ============================================================
# STYLE CONSTANTS (Roberto's BCI/Northleaf style)
# ============================================================
NAVY = RGBColor(0, 54, 91)        # #00365B - Primary brand color
DARK_GRAY = RGBColor(0, 0, 0)     # #000000 - Body text
LIGHT_GRAY = RGBColor(245, 245, 245)
ACCENT_GREEN = RGBColor(34, 139, 34)
ACCENT_RED = RGBColor(178, 34, 34)
WHITE = RGBColor(255, 255, 255)
FONT_NAME = "Arial"

# ============================================================
# SLIDE MODULE REGISTRY
# ============================================================
# Maps user intents to available slide modules

SLIDE_MODULES = {
    # Industry Analysis
    "industry_analysis": {
        "name": "Industry Analysis",
        "description": "Market overview, TAM/SAM/SOM, growth drivers, industry dynamics",
        "slides": ["market_overview", "porters_five_forces", "industry_value_chain", "market_sizing"],
        "function": "add_industry_analysis"
    },

    # Competitive Analysis
    "competitive_analysis": {
        "name": "Competitive Analysis",
        "description": "Competitive landscape, market positioning, SWOT, competitive moat",
        "slides": ["competitive_landscape", "swot_analysis", "competitive_moat", "market_positioning"],
        "function": "add_competitive_analysis"
    },

    # Financial Analysis
    "financial_analysis": {
        "name": "Financial Analysis",
        "description": "Historical financials, projections, key metrics, quality of earnings",
        "slides": ["historical_summary", "projected_summary", "key_metrics", "margin_analysis"],
        "function": "add_financial_analysis"
    },

    # Valuation
    "valuation": {
        "name": "Valuation Analysis",
        "description": "Trading comps, transaction comps, DCF, football field",
        "slides": ["trading_comps", "transaction_comps", "dcf_summary", "football_field"],
        "function": "add_valuation"
    },

    # LBO / Transaction
    "lbo_analysis": {
        "name": "LBO / Transaction Analysis",
        "description": "Sources & uses, debt schedule, returns analysis, sensitivity",
        "slides": ["sources_uses", "debt_schedule", "returns_analysis", "sensitivity_matrix"],
        "function": "add_lbo_analysis"
    },

    # Debt Analysis
    "debt_analysis": {
        "name": "Debt / Capital Structure Analysis",
        "description": "Debt capacity, capital structure, coverage ratios, debt schedule",
        "slides": ["debt_capacity", "capital_structure", "coverage_analysis", "debt_schedule"],
        "function": "add_debt_analysis"
    },

    # Management Analysis
    "management_analysis": {
        "name": "Management / Team Analysis",
        "description": "Leadership team, organizational structure, compensation, track record",
        "slides": ["leadership_overview", "org_structure", "management_incentives", "track_record"],
        "function": "add_management_analysis"
    },

    # Investment Thesis
    "investment_thesis": {
        "name": "Investment Thesis & Risks",
        "description": "Thesis pillars, risk assessment, value creation, recommendation",
        "slides": ["thesis_summary", "risk_matrix", "value_creation_bridge", "recommendation"],
        "function": "add_investment_thesis"
    },

    # Company Overview
    "company_overview": {
        "name": "Company Overview",
        "description": "Business description, revenue breakdown, customer analysis, operations",
        "slides": ["business_description", "revenue_breakdown", "customer_analysis", "operations"],
        "function": "add_company_overview"
    },

    # ============================================================
    # INSTITUTIONAL MODULES (7+ Day Case)
    # ============================================================

    # Scenario Analysis
    "scenario_analysis": {
        "name": "Scenario Analysis",
        "description": "Bull/Bear/Base case comparison with probability-weighted returns",
        "slides": ["scenario_summary", "scenario_drivers", "probability_weighted_returns"],
        "function": "add_scenario_analysis"
    },

    # Management vs Buyer Case
    "management_vs_buyer": {
        "name": "Management vs Buyer Case",
        "description": "Side-by-side comparison of management projections vs buyer underwriting",
        "slides": ["case_comparison", "variance_analysis", "returns_bridge"],
        "function": "add_management_vs_buyer"
    },

    # DCF Valuation
    "dcf_valuation": {
        "name": "DCF Valuation",
        "description": "Discounted cash flow analysis with terminal value and sensitivity",
        "slides": ["dcf_summary", "fcf_bridge", "terminal_value_sensitivity"],
        "function": "add_dcf_valuation"
    },

    # Covenant Analysis
    "covenant_analysis": {
        "name": "Covenant Analysis",
        "description": "Debt covenant compliance, leverage ratios, coverage analysis",
        "slides": ["covenant_overview", "leverage_trajectory", "coverage_analysis"],
        "function": "add_covenant_analysis"
    },

    # ============================================================
    # DUE DILIGENCE & QUALITY MODULES
    # ============================================================

    # Quality of Earnings
    "quality_of_earnings": {
        "name": "Quality of Earnings",
        "description": "EBITDA normalization, adjustments, run-rate analysis",
        "slides": ["qoe_summary", "adjustment_detail", "run_rate_bridge"],
        "function": "add_quality_of_earnings"
    },

    # Working Capital
    "working_capital_analysis": {
        "name": "Working Capital Analysis",
        "description": "NWC components, days analysis, peg mechanism",
        "slides": ["nwc_components", "nwc_peg"],
        "function": "add_working_capital_analysis"
    },

    # Customer Quality
    "customer_quality": {
        "name": "Customer & Revenue Quality",
        "description": "Customer concentration, retention, unit economics",
        "slides": ["concentration", "retention_metrics", "unit_economics"],
        "function": "add_customer_quality_analysis"
    },

    # Credit Analysis
    "credit_analysis": {
        "name": "Credit & Debt Sizing",
        "description": "Credit ratios, debt capacity, stress testing",
        "slides": ["credit_ratios", "debt_capacity", "stress_test"],
        "function": "add_credit_analysis_slides"
    },

    # ============================================================
    # RETURNS & CAPITAL STRUCTURE MODULES
    # ============================================================

    # Dividend Recap
    "dividend_recap": {
        "name": "Dividend Recapitalization",
        "description": "Mid-hold dividend, capital return, returns impact",
        "slides": ["recap_overview", "returns_impact"],
        "function": "add_dividend_recap_slides"
    },

    # Refinancing
    "refinancing": {
        "name": "Refinancing Analysis",
        "description": "Rate savings, maturity extension, cost-benefit",
        "slides": ["structure_comparison", "savings_analysis"],
        "function": "add_refinancing_slides"
    },

    # Waterfall
    "waterfall": {
        "name": "Equity Waterfall",
        "description": "Cap table, distribution waterfall, LP/GP splits",
        "slides": ["cap_table", "waterfall"],
        "function": "add_waterfall_slides"
    },

    # Sponsor Economics
    "sponsor_economics": {
        "name": "Sponsor Economics",
        "description": "GP carry, management fees, fund-level returns",
        "slides": ["deal_economics", "fund_economics"],
        "function": "add_sponsor_economics_slides"
    },

    # ============================================================
    # TRANSACTION STRUCTURE MODULES
    # ============================================================

    # Add-on Analysis
    "addon_analysis": {
        "name": "Add-on / Bolt-on Analysis",
        "description": "Platform plus add-on, synergies, combined returns",
        "slides": ["addon_overview", "combined_financials", "accretion_analysis"],
        "function": "add_addon_analysis_slides"
    },

    # Synergy Model
    "synergy_model": {
        "name": "Synergy Analysis",
        "description": "Revenue and cost synergies, timing, realization",
        "slides": ["synergy_summary", "synergy_detail", "synergy_timeline"],
        "function": "add_synergy_slides"
    },

    # Carve-out Analysis
    "carveout_analysis": {
        "name": "Carve-out Analysis",
        "description": "Standalone costs, stranded costs, TSA requirements",
        "slides": ["carveout_overview", "standalone_adjustments", "separation_timeline"],
        "function": "add_carveout_slides"
    },

    # Earnout Model
    "earnout_model": {
        "name": "Earnout / Contingent Consideration",
        "description": "Performance milestones, probability-weighted value",
        "slides": ["earnout_structure", "milestone_analysis"],
        "function": "add_earnout_slides"
    },

    # Purchase Price Allocation
    "ppa_analysis": {
        "name": "Purchase Price Allocation",
        "description": "Asset valuation, goodwill, intangibles",
        "slides": ["ppa_summary", "intangibles_detail"],
        "function": "add_ppa_slides"
    },

    # ============================================================
    # VALUE CREATION MODULES
    # ============================================================

    # Value Creation Bridge
    "value_creation": {
        "name": "Value Creation Bridge",
        "description": "EBITDA growth, multiple expansion, deleveraging",
        "slides": ["value_bridge", "value_drivers"],
        "function": "add_value_creation_slides"
    },

    # 100-Day Plan
    "hundred_day_plan": {
        "name": "100-Day Plan",
        "description": "Post-close priorities, quick wins, milestones",
        "slides": ["plan_overview", "priority_matrix", "timeline"],
        "function": "add_hundred_day_slides"
    },

    # Exit Readiness
    "exit_readiness": {
        "name": "Exit Readiness Assessment",
        "description": "Exit options, timing, value maximization",
        "slides": ["exit_options", "readiness_scorecard"],
        "function": "add_exit_readiness_slides"
    },

    # Management Incentive Plan
    "mip_analysis": {
        "name": "Management Incentive Plan",
        "description": "MIP structure, vesting, payout scenarios",
        "slides": ["mip_structure", "payout_scenarios"],
        "function": "add_mip_slides"
    },

    # ============================================================
    # SPECIALIZED MODULES
    # ============================================================

    # Rollup Model
    "rollup_model": {
        "name": "Rollup / Platform Build",
        "description": "Multi-acquisition strategy, combined metrics",
        "slides": ["rollup_summary", "acquisition_timeline", "combined_metrics"],
        "function": "add_rollup_slides"
    },

    # Tax Analysis
    "tax_analysis": {
        "name": "Tax Analysis",
        "description": "Tax structure, NOLs, step-up, effective rate",
        "slides": ["tax_overview", "tax_shield_analysis"],
        "function": "add_tax_slides"
    },

    # Control Premium Analysis
    "control_premium": {
        "name": "Control Premium Analysis",
        "description": "Premium to unaffected, historical premiums",
        "slides": ["premium_analysis", "precedent_premiums"],
        "function": "add_control_premium_slides"
    },

    # ============================================================
    # CFA VISUAL SLIDES (Research Challenge Patterns)
    # ============================================================

    # What Must Be True
    "wmbt": {
        "name": "What Must Be True (WMBT)",
        "description": "Investment thesis conditions with thresholds, status, and consequences",
        "slides": ["wmbt_checklist"],
        "function": "add_wmbt_slide"
    },

    # Risk-Mitigant Pairing
    "risk_mitigant": {
        "name": "Risk-Mitigant Pairing",
        "description": "Risk assessment with probability, impact, mitigant, and residual risk",
        "slides": ["risk_mitigant_table"],
        "function": "add_risk_mitigant_slide"
    },

    # Valuation Blending
    "valuation_blending": {
        "name": "Valuation Blending",
        "description": "Weighted valuation across methods with blended target price",
        "slides": ["valuation_blend_table", "valuation_blend_summary"],
        "function": "add_valuation_blending_slide"
    },

    # Value Chain
    "value_chain": {
        "name": "Value Chain Analysis",
        "description": "Horizontal value chain flow with company positioning highlighted",
        "slides": ["value_chain_flow"],
        "function": "add_value_chain_slide"
    },

    # Risk Matrix
    "risk_matrix": {
        "name": "Risk Matrix",
        "description": "3x3 probability-impact matrix with color-coded risk categories",
        "slides": ["risk_matrix_grid"],
        "function": "add_risk_matrix_slide"
    },

    # TAM Funnel
    "tam_funnel": {
        "name": "TAM/SAM/SOM Funnel",
        "description": "Market sizing funnel with descending-width bars and sources",
        "slides": ["tam_funnel_chart"],
        "function": "add_tam_funnel_slide"
    },

    # PESTEL
    "pestel": {
        "name": "PESTEL Analysis",
        "description": "6-factor macro analysis grid (Political, Economic, Social, Technological, Environmental, Legal)",
        "slides": ["pestel_grid"],
        "function": "add_pestel_slide"
    },

    # Porter's Radar
    "porters_radar": {
        "name": "Porter's Five Forces Radar",
        "description": "Quantified pentagon with force ratings (1-5 scale)",
        "slides": ["porters_radar_chart"],
        "function": "add_porters_radar_slide"
    },

    # ============================================================
    # CFA EXCEL+SLIDE PAIRS
    # ============================================================

    # Tornado Sensitivity
    "tornado_sensitivity": {
        "name": "Tornado Sensitivity Chart",
        "description": "Single-variable sensitivity with horizontal bars showing IRR/MOIC range",
        "slides": ["tornado_chart"],
        "function": "add_tornado_sensitivity_slide"
    },

    # Football Field
    "football_field": {
        "name": "Football Field Valuation",
        "description": "Valuation range chart with horizontal bars per methodology",
        "slides": ["football_field_chart"],
        "function": "add_football_field_slide"
    },

    # SOTP Waterfall
    "sotp_waterfall": {
        "name": "SOTP Waterfall",
        "description": "Sum-of-the-parts waterfall with segment contributions and EV bridge",
        "slides": ["sotp_waterfall_chart"],
        "function": "add_sotp_waterfall_slide"
    },
}


@dataclass
class SlideData:
    """Container for slide-specific data."""
    title: str = ""
    subtitle: str = ""
    content: List[str] = field(default_factory=list)
    table_headers: List[str] = field(default_factory=list)
    table_rows: List[List[str]] = field(default_factory=list)
    source: str = ""
    notes: str = ""


class SlideGenerator:
    """
    Modular slide generator that creates specific analysis slides on demand.

    Example:
        gen = SlideGenerator(company_name="Acme Corp")
        gen.add_industry_analysis()
        gen.add_competitive_analysis()
        gen.save("acme_analysis.pptx")
    """

    def __init__(self, company_name: str = "[Company Name]", date_str: str = None):
        """Initialize the slide generator."""
        self.prs = Presentation()
        self.prs.slide_width = Inches(10)
        self.prs.slide_height = Inches(7.5)
        self.company_name = company_name
        self.date_str = date_str or datetime.now().strftime("%B %d, %Y")
        self.page_num = 0
        self._has_cover = False

    def _next_page(self) -> int:
        """Get next page number."""
        self.page_num += 1
        return self.page_num

    # ============================================================
    # CORE SLIDE BUILDERS
    # ============================================================

    def _add_cover_slide(self, title: str, subtitle: str = ""):
        """Add cover slide if not already present."""
        if self._has_cover:
            return

        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])

        # Title
        txBox = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(9.1), Inches(1))
        p = txBox.text_frame.paragraphs[0]
        p.text = title
        p.font.name = FONT_NAME
        p.font.size = Pt(28)
        p.font.bold = True
        p.font.color.rgb = DARK_GRAY

        # Subtitle
        if subtitle:
            txBox2 = slide.shapes.add_textbox(Inches(0.5), Inches(5.2), Inches(9.1), Inches(1))
            p2 = txBox2.text_frame.paragraphs[0]
            p2.text = subtitle
            p2.font.name = FONT_NAME
            p2.font.size = Pt(14)
            p2.font.bold = True
            p2.font.color.rgb = DARK_GRAY

        # Date
        txBox3 = slide.shapes.add_textbox(Inches(2.3), Inches(6.3), Inches(2), Inches(0.2))
        p3 = txBox3.text_frame.paragraphs[0]
        p3.text = self.date_str
        p3.font.name = FONT_NAME
        p3.font.size = Pt(11)
        p3.font.color.rgb = DARK_GRAY

        self._has_cover = True

    def _add_section_header(self, section_title: str, section_number: int = None):
        """Add a section divider slide."""
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])

        # Section number
        if section_number:
            txBox = slide.shapes.add_textbox(Inches(0.5), Inches(3.0), Inches(1.5), Inches(1))
            p = txBox.text_frame.paragraphs[0]
            p.text = f"{section_number:02d}"
            p.font.name = FONT_NAME
            p.font.size = Pt(48)
            p.font.bold = True
            p.font.color.rgb = NAVY
            title_left = 2.0
        else:
            title_left = 0.5

        # Section title
        txBox2 = slide.shapes.add_textbox(Inches(title_left), Inches(3.0), Inches(7.5), Inches(1.5))
        p2 = txBox2.text_frame.paragraphs[0]
        p2.text = section_title
        p2.font.name = FONT_NAME
        p2.font.size = Pt(32)
        p2.font.bold = True
        p2.font.color.rgb = NAVY

    def _add_content_slide(self, title: str, content: List[str], subtitle: str = None,
                          source: str = None) -> Any:
        """Add a content slide with bullet points."""
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        page = self._next_page()

        # Title
        txBox = slide.shapes.add_textbox(Inches(0.3), Inches(0.2), Inches(9.5), Inches(0.3))
        p = txBox.text_frame.paragraphs[0]
        p.text = title
        p.font.name = FONT_NAME
        p.font.size = Pt(16)
        p.font.bold = True
        p.font.color.rgb = NAVY

        # Subtitle
        content_top = 1.0
        if subtitle:
            txBox_sub = slide.shapes.add_textbox(Inches(0.3), Inches(0.6), Inches(9.5), Inches(0.5))
            p_sub = txBox_sub.text_frame.paragraphs[0]
            p_sub.text = subtitle
            p_sub.font.name = FONT_NAME
            p_sub.font.size = Pt(11)
            p_sub.font.color.rgb = DARK_GRAY
            content_top = 1.2

        # Content
        txBox2 = slide.shapes.add_textbox(Inches(0.3), Inches(content_top), Inches(9.4), Inches(5.5))
        tf2 = txBox2.text_frame
        tf2.word_wrap = True

        for i, item in enumerate(content):
            p = tf2.paragraphs[0] if i == 0 else tf2.add_paragraph()
            if isinstance(item, tuple):
                p.text = item[0]
                p.level = item[1]
            else:
                p.text = f"• {item}" if item and not item.startswith("•") else item
                p.level = 0
            p.font.name = FONT_NAME
            p.font.size = Pt(10)
            p.font.color.rgb = DARK_GRAY
            p.space_after = Pt(6)

        # Page number
        txBox_pg = slide.shapes.add_textbox(Inches(9.0), Inches(7.0), Inches(0.8), Inches(0.3))
        p_pg = txBox_pg.text_frame.paragraphs[0]
        p_pg.text = str(page)
        p_pg.font.name = FONT_NAME
        p_pg.font.size = Pt(10)
        p_pg.font.color.rgb = DARK_GRAY

        # Source
        if source:
            txBox_src = slide.shapes.add_textbox(Inches(0.3), Inches(7.0), Inches(8.5), Inches(0.3))
            p_src = txBox_src.text_frame.paragraphs[0]
            p_src.text = f"Source: {source}"
            p_src.font.name = FONT_NAME
            p_src.font.size = Pt(8)
            p_src.font.color.rgb = DARK_GRAY

        return slide

    def _add_table_slide(self, title: str, headers: List[str], rows: List[List[str]],
                        col_widths: List[float] = None, subtitle: str = None,
                        source: str = None) -> Any:
        """Add a slide with a data table."""
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        page = self._next_page()

        # Title
        txBox = slide.shapes.add_textbox(Inches(0.3), Inches(0.2), Inches(9.5), Inches(0.3))
        p = txBox.text_frame.paragraphs[0]
        p.text = title
        p.font.name = FONT_NAME
        p.font.size = Pt(16)
        p.font.bold = True
        p.font.color.rgb = NAVY

        # Subtitle
        table_top = 0.8
        if subtitle:
            txBox_sub = slide.shapes.add_textbox(Inches(0.3), Inches(0.6), Inches(9.5), Inches(0.5))
            p_sub = txBox_sub.text_frame.paragraphs[0]
            p_sub.text = subtitle
            p_sub.font.name = FONT_NAME
            p_sub.font.size = Pt(11)
            p_sub.font.color.rgb = DARK_GRAY
            table_top = 1.2

        # Table
        num_cols = len(headers)
        num_rows = len(rows) + 1

        if col_widths is None:
            col_widths = [Inches(9.4 / num_cols)] * num_cols
        else:
            col_widths = [Inches(w) for w in col_widths]

        table = slide.shapes.add_table(num_rows, num_cols, Inches(0.3), Inches(table_top),
                                       sum(col_widths), Inches(0.35 * num_rows)).table

        for i, width in enumerate(col_widths):
            table.columns[i].width = width

        # Header row
        for i, header in enumerate(headers):
            cell = table.cell(0, i)
            cell.text = header
            cell.fill.solid()
            cell.fill.fore_color.rgb = NAVY
            p = cell.text_frame.paragraphs[0]
            p.font.name = FONT_NAME
            p.font.size = Pt(10)
            p.font.bold = True
            p.font.color.rgb = WHITE
            p.alignment = PP_ALIGN.CENTER

        # Data rows
        for row_idx, row_data in enumerate(rows):
            for col_idx, cell_value in enumerate(row_data):
                cell = table.cell(row_idx + 1, col_idx)
                cell.text = str(cell_value)
                p = cell.text_frame.paragraphs[0]
                p.font.name = FONT_NAME
                p.font.size = Pt(10)
                p.font.color.rgb = DARK_GRAY
                p.alignment = PP_ALIGN.LEFT if col_idx == 0 else PP_ALIGN.CENTER
                if col_idx == 0:
                    p.font.bold = True

        # Page number
        txBox_pg = slide.shapes.add_textbox(Inches(9.0), Inches(7.0), Inches(0.8), Inches(0.3))
        p_pg = txBox_pg.text_frame.paragraphs[0]
        p_pg.text = str(page)
        p_pg.font.name = FONT_NAME
        p_pg.font.size = Pt(10)
        p_pg.font.color.rgb = DARK_GRAY

        # Source
        if source:
            txBox_src = slide.shapes.add_textbox(Inches(0.3), Inches(7.0), Inches(8.5), Inches(0.3))
            p_src = txBox_src.text_frame.paragraphs[0]
            p_src.text = f"Source: {source}"
            p_src.font.name = FONT_NAME
            p_src.font.size = Pt(8)
            p_src.font.color.rgb = DARK_GRAY

        return slide

    def _add_framework_slide(self, title: str, framework_type: str, items: List,
                            subtitle: str = None) -> Any:
        """Add a framework slide (2x2, five forces, etc.)."""
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        page = self._next_page()

        # Title
        txBox = slide.shapes.add_textbox(Inches(0.3), Inches(0.2), Inches(9.5), Inches(0.3))
        p = txBox.text_frame.paragraphs[0]
        p.text = title
        p.font.name = FONT_NAME
        p.font.size = Pt(16)
        p.font.bold = True
        p.font.color.rgb = NAVY

        content_top = 1.0
        if subtitle:
            txBox_sub = slide.shapes.add_textbox(Inches(0.3), Inches(0.6), Inches(9.5), Inches(0.5))
            p_sub = txBox_sub.text_frame.paragraphs[0]
            p_sub.text = subtitle
            p_sub.font.name = FONT_NAME
            p_sub.font.size = Pt(11)
            p_sub.font.color.rgb = DARK_GRAY
            content_top = 1.2

        if framework_type == "2x2":
            positions = [
                (0.3, content_top, 4.5, 2.7),
                (5.0, content_top, 4.5, 2.7),
                (0.3, content_top + 2.9, 4.5, 2.7),
                (5.0, content_top + 2.9, 4.5, 2.7),
            ]
            colors = [NAVY, ACCENT_GREEN, RGBColor(200, 150, 50), ACCENT_RED]

            for i, (item_title, item_content) in enumerate(items[:4]):
                left, top, width, height = positions[i]

                box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                            Inches(left), Inches(top), Inches(width), Inches(height))
                box.fill.solid()
                box.fill.fore_color.rgb = RGBColor(248, 250, 252)
                box.line.color.rgb = colors[i]
                box.line.width = Pt(1.5)

                # Box title
                txBox_item = slide.shapes.add_textbox(Inches(left + 0.15), Inches(top + 0.1),
                                                      Inches(width - 0.3), Inches(0.4))
                p_item = txBox_item.text_frame.paragraphs[0]
                p_item.text = item_title
                p_item.font.name = FONT_NAME
                p_item.font.size = Pt(11)
                p_item.font.bold = True
                p_item.font.color.rgb = colors[i]

                # Box content
                txBox2 = slide.shapes.add_textbox(Inches(left + 0.15), Inches(top + 0.5),
                                                  Inches(width - 0.3), Inches(height - 0.6))
                tf2 = txBox2.text_frame
                tf2.word_wrap = True
                for j, bullet in enumerate(item_content):
                    p2 = tf2.paragraphs[0] if j == 0 else tf2.add_paragraph()
                    p2.text = f"• {bullet}"
                    p2.font.name = FONT_NAME
                    p2.font.size = Pt(9)
                    p2.font.color.rgb = DARK_GRAY

        elif framework_type == "five_forces":
            # Center - Competitive Rivalry
            center = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(3.6), Inches(3.0),
                                           Inches(2.5), Inches(1.3))
            center.fill.solid()
            center.fill.fore_color.rgb = NAVY
            center.line.fill.background()

            txBox_center = slide.shapes.add_textbox(Inches(3.7), Inches(3.2), Inches(2.3), Inches(1))
            p_center = txBox_center.text_frame.paragraphs[0]
            p_center.text = "Competitive\nRivalry"
            p_center.font.name = FONT_NAME
            p_center.font.size = Pt(12)
            p_center.font.bold = True
            p_center.font.color.rgb = WHITE
            p_center.alignment = PP_ALIGN.CENTER

            # Four forces
            force_positions = [
                (3.6, 1.0, "Threat of\nNew Entrants"),
                (3.6, 5.0, "Threat of\nSubstitutes"),
                (0.5, 3.0, "Supplier\nPower"),
                (7.0, 3.0, "Buyer\nPower"),
            ]

            for left, top, text in force_positions:
                box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                            Inches(left), Inches(top), Inches(2.5), Inches(1.2))
                box.fill.solid()
                box.fill.fore_color.rgb = RGBColor(248, 250, 252)
                box.line.color.rgb = NAVY
                box.line.width = Pt(1.5)

                txBox_force = slide.shapes.add_textbox(Inches(left + 0.1), Inches(top + 0.2),
                                                       Inches(2.3), Inches(1))
                p_force = txBox_force.text_frame.paragraphs[0]
                p_force.text = text
                p_force.font.name = FONT_NAME
                p_force.font.size = Pt(11)
                p_force.font.bold = True
                p_force.font.color.rgb = NAVY
                p_force.alignment = PP_ALIGN.CENTER

        # Page number
        txBox_pg = slide.shapes.add_textbox(Inches(9.0), Inches(7.0), Inches(0.8), Inches(0.3))
        p_pg = txBox_pg.text_frame.paragraphs[0]
        p_pg.text = str(page)
        p_pg.font.name = FONT_NAME
        p_pg.font.size = Pt(10)
        p_pg.font.color.rgb = DARK_GRAY

        return slide

    # ============================================================
    # MODULE: INDUSTRY ANALYSIS
    # ============================================================

    def add_industry_analysis(self, data: Dict = None, include_cover: bool = True):
        """
        Add industry analysis slides.

        Args:
            data: Optional dict with keys: tam, sam, som, growth_rate, drivers, headwinds
            include_cover: Whether to add a section header
        """
        data = data or {}

        if include_cover:
            self._add_section_header("Industry Analysis")

        # Market Overview
        self._add_content_slide(
            "Market Overview",
            [
                "Total Addressable Market (TAM)",
                (f"Global market size: ${data.get('tam', 'XX')} billion (2024)", 1),
                (f"Expected CAGR: {data.get('growth_rate', 'X')}% (2024-2029)", 1),
                "",
                "Key Market Drivers",
                (data.get('drivers', ['[Driver 1]', '[Driver 2]', '[Driver 3]'])[0], 1),
                (data.get('drivers', ['[Driver 1]', '[Driver 2]', '[Driver 3]'])[1] if len(data.get('drivers', [])) > 1 else '[Driver 2]', 1),
                "",
                "Market Headwinds",
                (data.get('headwinds', ['[Risk 1]', '[Risk 2]'])[0], 1),
            ],
            subtitle=f"Industry dynamics for {self.company_name}'s target market"
        )

        # Porter's Five Forces
        self._add_framework_slide(
            "Porter's Five Forces Analysis",
            "five_forces",
            [],
            subtitle="Competitive intensity assessment"
        )

        # Industry Value Chain
        self._add_content_slide(
            "Industry Value Chain",
            [
                "Upstream (Suppliers)",
                ("[Raw materials / components suppliers]", 1),
                ("[Technology / IP providers]", 1),
                "",
                "Midstream (Operations)",
                ("[Manufacturing / service delivery]", 1),
                ("[Distribution / logistics]", 1),
                "",
                "Downstream (Customers)",
                ("[End customers / channel partners]", 1),
                "",
                f"Company Position: {self.company_name} operates in [position in value chain]",
            ],
            subtitle="Value chain positioning and integration opportunities"
        )

        return self

    # ============================================================
    # MODULE: COMPETITIVE ANALYSIS
    # ============================================================

    def add_competitive_analysis(self, data: Dict = None, include_cover: bool = True):
        """
        Add competitive analysis slides.

        Args:
            data: Optional dict with competitors list and SWOT items
            include_cover: Whether to add a section header
        """
        data = data or {}

        if include_cover:
            self._add_section_header("Competitive Positioning")

        # Competitive Landscape Table
        competitors = data.get('competitors', [
            [self.company_name, "$XXX", "XX%", "[Strength]", "[Weakness]"],
            ["Competitor A", "$XXX", "XX%", "[Strength]", "[Weakness]"],
            ["Competitor B", "$XXX", "XX%", "[Strength]", "[Weakness]"],
            ["Competitor C", "$XXX", "XX%", "[Strength]", "[Weakness]"],
        ])

        self._add_table_slide(
            "Competitive Landscape",
            ["Company", "Revenue", "Share", "Key Strengths", "Key Weaknesses"],
            competitors,
            col_widths=[1.8, 1.2, 1.0, 2.4, 2.4],
            subtitle="Market share and competitive positioning",
            source="Company filings, Capital IQ, industry reports"
        )

        # SWOT Analysis
        swot = data.get('swot', {
            'strengths': ["[Market leadership]", "[Proprietary technology]", "[Strong relationships]"],
            'weaknesses': ["[Customer concentration]", "[Geographic limitations]", "[Tech gaps]"],
            'opportunities': ["[New market expansion]", "[Product extension]", "[M&A potential]"],
            'threats': ["[New entrants]", "[Technology disruption]", "[Regulatory changes]"],
        })

        self._add_framework_slide(
            "SWOT Analysis",
            "2x2",
            [
                ("STRENGTHS", swot['strengths']),
                ("WEAKNESSES", swot['weaknesses']),
                ("OPPORTUNITIES", swot['opportunities']),
                ("THREATS", swot['threats']),
            ],
            subtitle=f"Strategic positioning assessment for {self.company_name}"
        )

        # Competitive Moat
        self._add_content_slide(
            "Competitive Moat Assessment",
            [
                "Sources of Competitive Advantage",
                "",
                "1. Switching Costs: [HIGH / MEDIUM / LOW]",
                ("   [Explanation of customer lock-in]", 1),
                "",
                "2. Network Effects: [HIGH / MEDIUM / LOW]",
                ("   [Explanation of network dynamics]", 1),
                "",
                "3. Cost Advantages: [HIGH / MEDIUM / LOW]",
                ("   [Explanation of cost position]", 1),
                "",
                "4. Intangible Assets: [HIGH / MEDIUM / LOW]",
                ("   [Brands, patents, regulatory licenses]", 1),
            ],
            subtitle="Durability of competitive advantages"
        )

        return self

    # ============================================================
    # MODULE: FINANCIAL ANALYSIS
    # ============================================================

    def add_financial_analysis(self, data: Dict = None, include_cover: bool = True):
        """
        Add financial analysis slides.

        Args:
            data: Optional dict with historical and projected financials
            include_cover: Whether to add a section header
        """
        data = data or {}

        if include_cover:
            self._add_section_header("Financial Analysis")

        # Historical Summary
        historical = data.get('historical', [
            ["Revenue", "$XXX", "$XXX", "$XXX", "X%"],
            ["% Growth", "X%", "X%", "X%", "—"],
            ["Gross Profit", "$XX", "$XX", "$XX", "X%"],
            ["% Margin", "XX%", "XX%", "XX%", "—"],
            ["EBITDA", "$XX", "$XX", "$XX", "X%"],
            ["% Margin", "XX%", "XX%", "XX%", "—"],
            ["Free Cash Flow", "$XX", "$XX", "$XX", "X%"],
        ])

        self._add_table_slide(
            "Historical Financial Summary",
            ["($M)", "FY2022", "FY2023", "FY2024", "CAGR"],
            historical,
            col_widths=[2.5, 1.5, 1.5, 1.5, 1.5],
            subtitle="3-year historical performance summary",
            source="Company filings, management data"
        )

        # Projected Summary
        projected = data.get('projected', [
            ["Revenue", "$XXX", "$XXX", "$XXX", "$XXX", "$XXX"],
            ["% Growth", "X%", "X%", "X%", "X%", "X%"],
            ["EBITDA", "$XX", "$XX", "$XX", "$XX", "$XX"],
            ["% Margin", "XX%", "XX%", "XX%", "XX%", "XX%"],
            ["Unlevered FCF", "$XX", "$XX", "$XX", "$XX", "$XX"],
        ])

        self._add_table_slide(
            "Projected Financial Summary",
            ["($M)", "FY2025E", "FY2026E", "FY2027E", "FY2028E", "FY2029E"],
            projected,
            col_widths=[1.9, 1.3, 1.3, 1.3, 1.3, 1.3],
            subtitle="5-year projected operating performance"
        )

        # Key Metrics Comparison
        self._add_table_slide(
            "Key Operating Metrics",
            ["Metric", "Company", "Peer Median", "Variance", "Commentary"],
            [
                ["Revenue Growth (3Y)", "X%", "X%", "+/- X%", "[vs peers]"],
                ["Gross Margin", "XX%", "XX%", "+/- X%", "[vs peers]"],
                ["EBITDA Margin", "XX%", "XX%", "+/- X%", "[vs peers]"],
                ["FCF Conversion", "XX%", "XX%", "+/- X%", "[quality]"],
            ],
            col_widths=[2.2, 1.3, 1.5, 1.2, 2.3],
            subtitle="Benchmarking vs. peer group median"
        )

        return self

    # ============================================================
    # MODULE: DEBT / CAPITAL STRUCTURE ANALYSIS
    # ============================================================

    def add_debt_analysis(self, data: Dict = None, include_cover: bool = True):
        """
        Add debt and capital structure analysis slides.

        Args:
            data: Optional dict with debt schedule, coverage ratios, etc.
            include_cover: Whether to add a section header
        """
        data = data or {}

        if include_cover:
            self._add_section_header("Capital Structure Analysis")

        # Capital Structure Overview
        self._add_table_slide(
            "Current Capital Structure",
            ["Tranche", "Amount ($M)", "Rate", "Maturity", "Covenants"],
            data.get('debt_tranches', [
                ["Senior Secured (TLB)", "$XX", "L+XXX", "20XX", "X.Xx Leverage"],
                ["Second Lien", "$XX", "L+XXX", "20XX", "X.Xx Leverage"],
                ["Revolving Credit", "$XX", "L+XXX", "20XX", "—"],
                ["Total Debt", "$XXX", "—", "—", "—"],
                ["Cash", "($XX)", "—", "—", "—"],
                ["Net Debt", "$XXX", "—", "—", "—"],
            ]),
            col_widths=[2.2, 1.5, 1.3, 1.3, 2.2],
            subtitle="Existing debt facilities and key terms"
        )

        # Debt Capacity Analysis
        self._add_table_slide(
            "Debt Capacity Analysis",
            ["Metric", "Current", "Conservative", "Moderate", "Aggressive"],
            data.get('debt_capacity', [
                ["EBITDA", "$XX", "$XX", "$XX", "$XX"],
                ["Total Debt / EBITDA", "X.Xx", "3.0x", "4.5x", "6.0x"],
                ["Senior Debt / EBITDA", "X.Xx", "2.0x", "3.0x", "4.0x"],
                ["Interest Coverage", "X.Xx", "3.0x", "2.5x", "2.0x"],
                ["Implied Debt Capacity", "$XXX", "$XXX", "$XXX", "$XXX"],
            ]),
            col_widths=[2.4, 1.5, 1.5, 1.5, 1.5],
            subtitle="Incremental debt capacity under various leverage scenarios"
        )

        # Coverage Analysis
        self._add_content_slide(
            "Coverage Ratio Analysis",
            [
                "Interest Coverage (EBITDA / Cash Interest)",
                ("Current: X.Xx | Pro Forma: X.Xx | Covenant: X.Xx", 1),
                ("Headroom: $XX million or XX%", 1),
                "",
                "Fixed Charge Coverage ((EBITDA - CapEx) / Fixed Charges)",
                ("Current: X.Xx | Pro Forma: X.Xx | Covenant: X.Xx", 1),
                ("Headroom: $XX million or XX%", 1),
                "",
                "Leverage (Total Debt / EBITDA)",
                ("Current: X.Xx | Pro Forma: X.Xx | Covenant: X.Xx", 1),
                ("Deleveraging trajectory: X.Xx by Year 3", 1),
                "",
                "Assessment: [COMFORTABLE / TIGHT / AT RISK]",
            ],
            subtitle="Covenant compliance and cushion analysis"
        )

        return self

    # ============================================================
    # MODULE: LBO / TRANSACTION ANALYSIS
    # ============================================================

    def add_lbo_analysis(self, data: Dict = None, include_cover: bool = True):
        """
        Add LBO and transaction analysis slides.

        Args:
            data: Optional dict with S&U, returns, sensitivities
            include_cover: Whether to add a section header
        """
        data = data or {}

        if include_cover:
            self._add_section_header("Transaction Analysis")

        # Sources & Uses
        self._add_table_slide(
            "Sources and Uses of Funds",
            ["Sources", "$M", "%", "", "Uses", "$M", "%"],
            data.get('sources_uses', [
                ["Senior Debt", "$XX", "XX%", "", "Purchase Price", "$XXX", "XX%"],
                ["Subordinated Debt", "$XX", "XX%", "", "Refinance Debt", "$XX", "XX%"],
                ["Rollover Equity", "$XX", "XX%", "", "Transaction Fees", "$XX", "XX%"],
                ["Sponsor Equity", "$XX", "XX%", "", "Financing Fees", "$XX", "XX%"],
                ["Total Sources", "$XXX", "100%", "", "Total Uses", "$XXX", "100%"],
            ]),
            col_widths=[1.6, 1.0, 0.9, 0.2, 1.6, 1.0, 0.9],
            subtitle="Transaction funding structure"
        )

        # Returns Analysis
        self._add_table_slide(
            "LBO Returns Analysis",
            ["Exit Yr", "EBITDA", "Multiple", "Exit EV", "Equity", "MOIC", "IRR"],
            data.get('returns', [
                ["Year 3", "$XX", "X.Xx", "$XXX", "$XXX", "X.Xx", "XX%"],
                ["Year 4", "$XX", "X.Xx", "$XXX", "$XXX", "X.Xx", "XX%"],
                ["Year 5", "$XX", "X.Xx", "$XXX", "$XXX", "X.Xx", "XX%"],
            ]),
            col_widths=[1.2, 1.2, 1.4, 1.4, 1.4, 1.2, 1.2],
            subtitle="Returns by exit year (base case multiple)"
        )

        # Sensitivity Matrix
        self._add_content_slide(
            "Returns Sensitivity Analysis",
            [
                "[Insert sensitivity matrix here]",
                "",
                "Key Sensitivities:",
                ("Entry Multiple: X.Xx - X.Xx (IRR range: XX% - XX%)", 1),
                ("Exit Multiple: X.Xx - X.Xx (IRR range: XX% - XX%)", 1),
                ("Revenue Growth: X% - X% (IRR range: XX% - XX%)", 1),
                ("EBITDA Margin: XX% - XX% (IRR range: XX% - XX%)", 1),
                "",
                "Downside Protection:",
                ("Break-even exit multiple: X.Xx (vs. entry of X.Xx)", 1),
                ("Capital protection in downside: X.Xx MOIC", 1),
            ],
            subtitle="IRR sensitivity to key value drivers"
        )

        return self

    # ============================================================
    # MODULE: MANAGEMENT ANALYSIS
    # ============================================================

    def add_management_analysis(self, data: Dict = None, include_cover: bool = True):
        """
        Add management and team analysis slides.

        Args:
            data: Optional dict with management team info
            include_cover: Whether to add a section header
        """
        data = data or {}

        if include_cover:
            self._add_section_header("Management Analysis")

        # Leadership Overview
        self._add_table_slide(
            "Leadership Team Overview",
            ["Name", "Title", "Tenure", "Background", "Ownership"],
            data.get('leadership', [
                ["[CEO Name]", "CEO", "X yrs", "[Prior experience]", "X%"],
                ["[CFO Name]", "CFO", "X yrs", "[Prior experience]", "X%"],
                ["[COO Name]", "COO", "X yrs", "[Prior experience]", "X%"],
                ["[CTO Name]", "CTO", "X yrs", "[Prior experience]", "X%"],
            ]),
            col_widths=[1.8, 1.2, 1.0, 3.2, 1.2],
            subtitle="Key executives and management depth"
        )

        # Management Assessment
        self._add_content_slide(
            "Management Assessment",
            [
                "Strengths",
                ("[Deep industry expertise and relationships]", 1),
                ("[Proven track record of execution]", 1),
                ("[Strong alignment through equity ownership]", 1),
                "",
                "Areas to Monitor",
                ("[Succession planning for key roles]", 1),
                ("[Bandwidth for growth initiatives]", 1),
                "",
                "Key Person Risk",
                ("CEO: [HIGH / MEDIUM / LOW] - [Mitigation plan]", 1),
                ("Technical Leadership: [HIGH / MEDIUM / LOW]", 1),
                "",
                "Incentive Alignment",
                ("Management rollover: XX% of existing equity", 1),
                ("Option pool: XX% fully diluted", 1),
            ],
            subtitle="Leadership quality and continuity assessment"
        )

        return self

    # ============================================================
    # MODULE: INVESTMENT THESIS
    # ============================================================

    def add_investment_thesis(self, data: Dict = None, include_cover: bool = True):
        """
        Add investment thesis and recommendation slides.

        Args:
            data: Optional dict with thesis pillars, risks, recommendation
            include_cover: Whether to add a section header
        """
        data = data or {}

        if include_cover:
            self._add_section_header("Investment Thesis & Risks")

        # Investment Thesis
        self._add_content_slide(
            "Investment Thesis",
            [
                "Pillar 1: [Attractive Market Dynamics]",
                ("[Supporting point A]", 1),
                ("[Supporting point B]", 1),
                "",
                "Pillar 2: [Defensible Competitive Position]",
                ("[Supporting point A]", 1),
                ("[Supporting point B]", 1),
                "",
                "Pillar 3: [Multiple Paths to Value Creation]",
                ("[Organic growth initiatives]", 1),
                ("[Margin expansion opportunities]", 1),
                ("[M&A / tuck-in potential]", 1),
            ],
            subtitle=f"Core investment rationale for {self.company_name}"
        )

        # Risk Matrix
        risks = data.get('risks', {
            'high_high': ["[Key risk 1]", "[Key risk 2]", "Mitigant: [Action]"],
            'high_low': ["[Tail risk 1]", "Mitigant: [Action]"],
            'low_high': ["[Operational risk 1]", "Mitigant: [Action]"],
            'low_low': ["[Minor risk]", "Monitor"],
        })

        self._add_framework_slide(
            "Risk Assessment Matrix",
            "2x2",
            [
                ("HIGH IMPACT / HIGH PROBABILITY", risks['high_high']),
                ("HIGH IMPACT / LOW PROBABILITY", risks['high_low']),
                ("LOW IMPACT / HIGH PROBABILITY", risks['low_high']),
                ("LOW IMPACT / LOW PROBABILITY", risks['low_low']),
            ],
            subtitle="Risk prioritization and mitigation strategies"
        )

        # Recommendation
        recommendation = data.get('recommendation', 'INVEST')
        self._add_content_slide(
            "Investment Recommendation",
            [
                f"RECOMMENDATION: {recommendation}",
                "",
                "Key Reasons:",
                ("[Reason 1: Most compelling factor]", 1),
                ("[Reason 2: Second factor]", 1),
                ("[Reason 3: Third factor]", 1),
                "",
                "Key Conditions / Next Steps:",
                ("[Condition 1: e.g., Complete management meetings]", 1),
                ("[Condition 2: e.g., Validate customer relationships]", 1),
                "",
                "Proposed Terms:",
                ("Entry Multiple: X.Xx LTM EBITDA | Equity Check: $XX million", 1),
                ("Target IRR: XX% | Target MOIC: X.Xx", 1),
            ],
            subtitle="Final investment recommendation and terms"
        )

        return self

    # ============================================================
    # MODULE: COMPANY OVERVIEW
    # ============================================================

    def add_company_overview(self, data: Dict = None, include_cover: bool = True):
        """
        Add company overview slides.

        Args:
            data: Optional dict with company info
            include_cover: Whether to add a section header
        """
        data = data or {}

        if include_cover:
            self._add_section_header("Company Overview")

        # Business Description
        self._add_content_slide(
            "Business Description",
            [
                "Company Background",
                (f"Founded in [YEAR], headquartered in [LOCATION]", 1),
                ("[Brief description of core business]", 1),
                ("[Number of employees, locations, geographies]", 1),
                "",
                "Key Business Segments",
                ("[Segment 1]: XX% of revenue - [Description]", 1),
                ("[Segment 2]: XX% of revenue - [Description]", 1),
                ("[Segment 3]: XX% of revenue - [Description]", 1),
            ],
            subtitle=f"Overview of {self.company_name}'s operations"
        )

        # Revenue Breakdown
        self._add_table_slide(
            "Revenue Breakdown",
            ["Segment", "Revenue ($M)", "% of Total", "Growth", "Margin"],
            data.get('segments', [
                ["Segment A", "$XX", "XX%", "+X%", "XX%"],
                ["Segment B", "$XX", "XX%", "+X%", "XX%"],
                ["Segment C", "$XX", "XX%", "+X%", "XX%"],
                ["Total", "$XXX", "100%", "+X%", "XX%"],
            ]),
            col_widths=[2.5, 1.5, 1.5, 1.5, 1.5],
            subtitle="LTM revenue breakdown by business segment"
        )

        # Customer Analysis
        self._add_content_slide(
            "Customer Analysis",
            [
                "Customer Base Profile",
                ("Total customers: [X,XXX]", 1),
                ("Average contract value: $[XX,XXX]", 1),
                ("Customer retention rate: [XX%]", 1),
                "",
                "Top Customer Concentration",
                ("Top 10 customers: XX% of revenue", 1),
                ("Largest customer: XX% of revenue", 1),
                ("Average relationship tenure: [X] years", 1),
            ],
            subtitle="Customer concentration and retention analysis"
        )

        return self

    # ============================================================
    # MODULE: VALUATION
    # ============================================================

    def add_valuation(self, data: Dict = None, include_cover: bool = True):
        """
        Add valuation analysis slides.

        Args:
            data: Optional dict with comps and valuation data
            include_cover: Whether to add a section header
        """
        data = data or {}

        if include_cover:
            self._add_section_header("Valuation")

        # Trading Comps
        self._add_table_slide(
            "Public Comparable Companies",
            ["Company", "EV ($M)", "EV/Rev", "EV/EBITDA", "Growth", "Margin"],
            data.get('trading_comps', [
                ["Comp A", "$X,XXX", "X.Xx", "XX.Xx", "X%", "XX%"],
                ["Comp B", "$X,XXX", "X.Xx", "XX.Xx", "X%", "XX%"],
                ["Comp C", "$X,XXX", "X.Xx", "XX.Xx", "X%", "XX%"],
                ["Mean", "—", "X.Xx", "XX.Xx", "X%", "XX%"],
                ["Median", "—", "X.Xx", "XX.Xx", "X%", "XX%"],
            ]),
            col_widths=[1.8, 1.3, 1.3, 1.5, 1.3, 1.3],
            subtitle="Trading comparables as of [Date]",
            source="Capital IQ, Company filings"
        )

        # Transaction Comps
        self._add_table_slide(
            "Precedent Transactions",
            ["Date", "Target", "Acquirer", "EV ($M)", "EV/Rev", "EV/EBITDA"],
            data.get('transaction_comps', [
                ["MM/YY", "Target A", "Buyer A", "$XXX", "X.Xx", "XX.Xx"],
                ["MM/YY", "Target B", "Buyer B", "$XXX", "X.Xx", "XX.Xx"],
                ["Mean", "—", "—", "—", "X.Xx", "XX.Xx"],
            ]),
            col_widths=[1.0, 1.8, 1.8, 1.3, 1.3, 1.5],
            subtitle="Selected M&A transactions",
            source="Capital IQ, PitchBook"
        )

        # Football Field
        self._add_content_slide(
            "Valuation Summary",
            [
                "[Insert football field chart here]",
                "",
                "Methodology Summary:",
                ("Trading Comps: $XXX - $XXX million", 1),
                ("Transaction Comps: $XXX - $XXX million", 1),
                ("DCF Analysis: $XXX - $XXX million", 1),
                ("LBO Analysis: $XXX - $XXX million (at XX% IRR)", 1),
                "",
                "Selected Valuation Range: $XXX - $XXX million",
                "Implied Multiple: X.Xx - X.Xx LTM EBITDA",
            ],
            subtitle="Valuation range across methodologies"
        )

        return self

    # ============================================================
    # INSTITUTIONAL MODULES (7+ Day Case)
    # ============================================================

    def add_scenario_analysis(self, data: Dict = None, include_cover: bool = True):
        """
        Add scenario analysis slides (Bull/Bear/Base case comparison).

        Args:
            data: Optional dict with scenario assumptions and returns
            include_cover: Whether to add a section header
        """
        data = data or {}

        if include_cover:
            self._add_section_header("Scenario Analysis")

        # Scenario Summary
        self._add_table_slide(
            "Scenario Comparison Summary",
            ["Metric", "Bear Case", "Base Case", "Bull Case"],
            data.get('scenario_summary', [
                ["Probability", "20%", "60%", "20%"],
                ["Revenue CAGR", "3%", "6%", "10%"],
                ["Exit EBITDA Margin", "18%", "22%", "26%"],
                ["Exit Multiple", "7.0x", "8.0x", "9.0x"],
                ["Exit Year", "5", "5", "5"],
                ["MOIC", "1.5x", "2.2x", "3.0x"],
                ["IRR", "8%", "17%", "25%"],
            ]),
            col_widths=[2.5, 2.0, 2.0, 2.0],
            subtitle="Key assumptions and returns by scenario"
        )

        # Scenario Drivers
        self._add_content_slide(
            "Scenario Key Drivers",
            [
                "Bear Case Assumptions (20% probability)",
                ("Revenue growth limited to GDP + inflation", 1),
                ("Margin compression from competitive pressure", 1),
                ("Multiple contraction to historical trough", 1),
                "",
                "Base Case Assumptions (60% probability)",
                ("Revenue growth aligned with industry", 1),
                ("Margin improvement from operational initiatives", 1),
                ("Exit multiple at entry (no expansion)", 1),
                "",
                "Bull Case Assumptions (20% probability)",
                ("Market share gains driving above-market growth", 1),
                ("Full margin potential realized", 1),
                ("Multiple expansion to premium peer levels", 1),
            ],
            subtitle="Key assumptions driving each scenario"
        )

        # Probability-Weighted Returns
        self._add_table_slide(
            "Probability-Weighted Returns",
            ["Scenario", "Probability", "MOIC", "IRR", "Weighted MOIC", "Weighted IRR"],
            data.get('weighted_returns', [
                ["Bear Case", "20%", "1.5x", "8%", "0.30x", "1.6%"],
                ["Base Case", "60%", "2.2x", "17%", "1.32x", "10.2%"],
                ["Bull Case", "20%", "3.0x", "25%", "0.60x", "5.0%"],
                ["Blended", "100%", "—", "—", "2.22x", "16.8%"],
            ]),
            col_widths=[1.8, 1.3, 1.3, 1.3, 1.5, 1.5],
            subtitle="Expected returns under probability-weighted scenarios"
        )

        # Downside Protection
        self._add_content_slide(
            "Downside Protection Analysis",
            [
                "Bear Case Return Profile",
                ("Bear case MOIC: 1.5x", 1),
                ("Implies return of capital plus modest return", 1),
                ("Key downside drivers:", 1),
                ("  - Revenue shortfall: [X]% impact", 2),
                ("  - Margin compression: [X]% impact", 2),
                ("  - Multiple contraction: [X]% impact", 2),
                "",
                "Mitigants",
                ("Asset value floor: $[XXX]M (Y.Yx coverage)", 1),
                ("Debt paydown provides equity accretion", 1),
                ("Contractual revenue base of [X]%", 1),
                ("Management incentive alignment", 1),
            ],
            subtitle="Downside scenario analysis and protection measures"
        )

        return self

    def add_management_vs_buyer(self, data: Dict = None, include_cover: bool = True):
        """
        Add management vs buyer case comparison slides.

        Args:
            data: Optional dict with management and buyer projections
            include_cover: Whether to add a section header
        """
        data = data or {}

        if include_cover:
            self._add_section_header("Management vs Buyer Case")

        # Case Comparison Table
        self._add_table_slide(
            "Projection Case Comparison",
            ["Year", "1", "2", "3", "4", "5", "CAGR"],
            data.get('case_comparison', [
                ["Management Revenue", "$XXX", "$XXX", "$XXX", "$XXX", "$XXX", "X%"],
                ["Buyer Revenue", "$XXX", "$XXX", "$XXX", "$XXX", "$XXX", "X%"],
                ["  Variance", "($X)", "($X)", "($X)", "($X)", "($X)", "—"],
                ["", "", "", "", "", "", ""],
                ["Management EBITDA", "$XX", "$XX", "$XX", "$XX", "$XX", "X%"],
                ["Buyer EBITDA", "$XX", "$XX", "$XX", "$XX", "$XX", "X%"],
                ["  Variance", "($X)", "($X)", "($X)", "($X)", "($X)", "—"],
            ]),
            col_widths=[2.2, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0],
            subtitle="Management plan vs buyer underwriting assumptions"
        )

        # Variance Drivers
        self._add_content_slide(
            "Key Variance Drivers",
            [
                "Revenue Haircut Rationale",
                ("New product launch timing: [X]% haircut", 1),
                ("Market growth assumptions: [X]% haircut", 1),
                ("Pricing assumptions: [X]% haircut", 1),
                "",
                "Margin Haircut Rationale",
                ("Cost savings timing: [X] bps haircut", 1),
                ("Synergy realization: [X] bps haircut", 1),
                ("Investment requirements: [X] bps haircut", 1),
                "",
                "Total Impact",
                ("Year 5 Revenue variance: [X]%", 1),
                ("Year 5 EBITDA variance: [X]%", 1),
                ("IRR impact: [X] percentage points", 1),
            ],
            subtitle="Justification for buyer case adjustments"
        )

        # Returns Bridge
        self._add_table_slide(
            "Returns Comparison (Year 5 Exit)",
            ["Metric", "Management Case", "Buyer Case", "Difference"],
            data.get('returns_comparison', [
                ["Exit EBITDA", "$XXM", "$XXM", "($XM)"],
                ["Exit EV", "$XXXM", "$XXXM", "($XXM)"],
                ["Exit Equity", "$XXXM", "$XXXM", "($XXM)"],
                ["MOIC", "X.Xx", "X.Xx", "(X.Xx)"],
                ["IRR", "XX%", "XX%", "(X%)"],
            ]),
            col_widths=[2.5, 2.0, 2.0, 2.0],
            subtitle="Return impact of buyer case haircuts"
        )

        return self

    def add_dcf_valuation(self, data: Dict = None, include_cover: bool = True):
        """
        Add DCF valuation slides.

        Args:
            data: Optional dict with DCF inputs and outputs
            include_cover: Whether to add a section header
        """
        data = data or {}

        if include_cover:
            self._add_section_header("DCF Valuation")

        # DCF Summary
        self._add_table_slide(
            "DCF Valuation Summary",
            ["Component", "Value ($M)", "Notes"],
            data.get('dcf_summary', [
                ["PV of Projection Period FCF", "$XXX", "Years 1-5"],
                ["PV of Terminal Value", "$XXX", "Exit multiple / perpetuity"],
                ["Enterprise Value", "$XXX", "—"],
                ["Less: Net Debt", "($XX)", "As of close"],
                ["Equity Value", "$XXX", "—"],
                ["Implied EV/EBITDA", "X.Xx", "Based on LTM"],
            ]),
            col_widths=[3.5, 2.0, 3.0],
            subtitle="DCF-derived valuation"
        )

        # FCF Bridge
        self._add_table_slide(
            "Free Cash Flow Build",
            ["Year", "1", "2", "3", "4", "5", "Terminal"],
            data.get('fcf_build', [
                ["EBITDA", "$XX", "$XX", "$XX", "$XX", "$XX", "$XX"],
                ["(-) D&A", "($X)", "($X)", "($X)", "($X)", "($X)", "($X)"],
                ["EBIT", "$XX", "$XX", "$XX", "$XX", "$XX", "$XX"],
                ["(-) Taxes", "($X)", "($X)", "($X)", "($X)", "($X)", "($X)"],
                ["(+) D&A", "$X", "$X", "$X", "$X", "$X", "$X"],
                ["(-) CapEx", "($X)", "($X)", "($X)", "($X)", "($X)", "($X)"],
                ["(-) NWC", "($X)", "($X)", "($X)", "($X)", "($X)", "$0"],
                ["UFCF", "$XX", "$XX", "$XX", "$XX", "$XX", "$XX"],
            ]),
            col_widths=[2.2, 1.0, 1.0, 1.0, 1.0, 1.0, 1.3],
            subtitle="Unlevered free cash flow projection"
        )

        # Terminal Value Sensitivity
        self._add_table_slide(
            "DCF Sensitivity Analysis",
            ["", "8.0x", "8.5x", "9.0x", "9.5x", "10.0x"],
            data.get('dcf_sensitivity', [
                ["WACC 8.0%", "$XXX", "$XXX", "$XXX", "$XXX", "$XXX"],
                ["WACC 9.0%", "$XXX", "$XXX", "$XXX", "$XXX", "$XXX"],
                ["WACC 10.0%", "$XXX", "$XXX", "$XXX", "$XXX", "$XXX"],
                ["WACC 11.0%", "$XXX", "$XXX", "$XXX", "$XXX", "$XXX"],
                ["WACC 12.0%", "$XXX", "$XXX", "$XXX", "$XXX", "$XXX"],
            ]),
            col_widths=[1.8, 1.4, 1.4, 1.4, 1.4, 1.4],
            subtitle="Enterprise value sensitivity to WACC and exit multiple"
        )

        return self

    def add_covenant_analysis(self, data: Dict = None, include_cover: bool = True):
        """
        Add covenant analysis slides.

        Args:
            data: Optional dict with covenant projections
            include_cover: Whether to add a section header
        """
        data = data or {}

        if include_cover:
            self._add_section_header("Covenant Analysis")

        # Covenant Overview
        self._add_table_slide(
            "Covenant Package Overview",
            ["Covenant", "Threshold", "Test Frequency", "Cure Rights"],
            data.get('covenant_package', [
                ["Max Leverage Ratio", "X.Xx Total Debt / EBITDA", "Quarterly", "Yes - equity cure"],
                ["Min Interest Coverage", "X.Xx EBITDA / Interest", "Quarterly", "Yes - equity cure"],
                ["Min Fixed Charge Coverage", "X.Xx (EBITDA-CapEx) / Debt Service", "Quarterly", "Limited"],
                ["Max CapEx", "$XXM annually", "Annual", "Carry-forward"],
                ["Dividend Restrictions", "Available basket + RP growth", "Ongoing", "N/A"],
            ]),
            col_widths=[2.5, 2.5, 1.5, 2.0],
            subtitle="Summary of financial covenant package"
        )

        # Leverage Trajectory
        self._add_table_slide(
            "Leverage Ratio Trajectory",
            ["Year", "Close", "1", "2", "3", "4", "5"],
            data.get('leverage_trajectory', [
                ["Total Debt", "$XXX", "$XXX", "$XXX", "$XXX", "$XXX", "$XXX"],
                ["EBITDA", "$XX", "$XX", "$XX", "$XX", "$XX", "$XX"],
                ["Leverage Ratio", "X.Xx", "X.Xx", "X.Xx", "X.Xx", "X.Xx", "X.Xx"],
                ["Covenant", "X.Xx", "X.Xx", "X.Xx", "X.Xx", "X.Xx", "X.Xx"],
                ["Headroom", "X.Xx", "X.Xx", "X.Xx", "X.Xx", "X.Xx", "X.Xx"],
                ["Status", "Pass", "Pass", "Pass", "Pass", "Pass", "Pass"],
            ]),
            col_widths=[2.0, 1.2, 1.2, 1.2, 1.2, 1.2, 1.2],
            subtitle="Leverage covenant compliance over projection period"
        )

        # Coverage Analysis
        self._add_table_slide(
            "Interest Coverage Trajectory",
            ["Year", "Close", "1", "2", "3", "4", "5"],
            data.get('coverage_trajectory', [
                ["EBITDA", "$XX", "$XX", "$XX", "$XX", "$XX", "$XX"],
                ["Interest Expense", "$X", "$X", "$X", "$X", "$X", "$X"],
                ["Coverage Ratio", "X.Xx", "X.Xx", "X.Xx", "X.Xx", "X.Xx", "X.Xx"],
                ["Covenant", "X.Xx", "X.Xx", "X.Xx", "X.Xx", "X.Xx", "X.Xx"],
                ["Headroom", "X.Xx", "X.Xx", "X.Xx", "X.Xx", "X.Xx", "X.Xx"],
                ["Status", "Pass", "Pass", "Pass", "Pass", "Pass", "Pass"],
            ]),
            col_widths=[2.0, 1.2, 1.2, 1.2, 1.2, 1.2, 1.2],
            subtitle="Interest coverage covenant compliance"
        )

        # Covenant Stress Test
        self._add_content_slide(
            "Covenant Stress Test",
            [
                "Downside Scenario Impact on Covenants",
                "",
                "Leverage Ratio Stress Test:",
                ("Base case: X.Xx (passes with X.Xx headroom)", 1),
                ("Bear case: X.Xx (passes with X.Xx headroom)", 1),
                ("Break-even EBITDA: $XXM (XX% below plan)", 1),
                "",
                "Interest Coverage Stress Test:",
                ("Base case: X.Xx (passes with X.Xx headroom)", 1),
                ("Bear case: X.Xx (passes with X.Xx headroom)", 1),
                ("Break-even EBITDA: $XXM (XX% below plan)", 1),
                "",
                "Key Takeaway: Covenant package provides [adequate/tight] flexibility",
            ],
            subtitle="Covenant compliance under stress scenarios"
        )

        return self

    # ============================================================
    # DUE DILIGENCE & QUALITY MODULES
    # ============================================================

    def add_quality_of_earnings(self, data: Dict = None, include_cover: bool = True):
        """
        Add Quality of Earnings analysis slides.

        Args:
            data: Optional dict with QoE data
            include_cover: Whether to add a section header
        """
        data = data or {}

        if include_cover:
            self._add_section_header("Quality of Earnings")

        # QoE Summary
        self._add_table_slide(
            "EBITDA Normalization Summary",
            ["Adjustment", "LTM ($M)", "FY-1 ($M)", "Commentary"],
            data.get('adjustments', [
                ["Reported EBITDA", "$XX.X", "$XX.X", "Per management accounts"],
                ["Owner comp normalization", "+$X.X", "+$X.X", "Above-market compensation"],
                ["One-time legal costs", "+$X.X", "+$X.X", "Settlement costs"],
                ["Non-recurring consulting", "+$X.X", "+$X.X", "Carve-out related"],
                ["Stock-based compensation", "+$X.X", "+$X.X", "Non-cash expense"],
                ["Run-rate cost savings", "+$X.X", "—", "In-progress initiatives"],
                ["Adjusted EBITDA", "$XX.X", "$XX.X", "Normalized earnings"],
            ]),
            col_widths=[2.5, 1.5, 1.5, 3.0],
            subtitle="EBITDA adjustments and normalization"
        )

        # Adjustment Categories
        self._add_content_slide(
            "Key QoE Findings",
            [
                "Revenue Quality",
                ("XX% of revenue is recurring/contractual", 1),
                ("Top 10 customers represent XX% of revenue", 1),
                ("No material revenue recognition issues identified", 1),
                "",
                "Cost Structure",
                ("Gross margin stable at XX% over 3 years", 1),
                ("SG&A includes $X.XM of one-time items", 1),
                ("Identified $X.XM of run-rate cost savings", 1),
                "",
                "Key Adjustments",
                ("Total add-backs: $X.XM (XX% of reported EBITDA)", 1),
                ("Largest adjustment: [description]", 1),
                ("All adjustments supported by documentation", 1),
            ],
            subtitle="Summary of due diligence findings"
        )

        # Run-Rate Bridge
        self._add_table_slide(
            "Run-Rate EBITDA Bridge",
            ["Item", "Amount ($M)", "Timing", "Confidence"],
            data.get('run_rate', [
                ["Adjusted LTM EBITDA", "$XX.X", "—", "—"],
                ["Full-year price increase", "+$X.X", "Q1 implementation", "High"],
                ["New customer annualization", "+$X.X", "Signed contracts", "High"],
                ["Planned cost reductions", "+$X.X", "In progress", "Medium"],
                ["Lost customer impact", "($X.X)", "Known churn", "High"],
                ["Run-Rate EBITDA", "$XX.X", "—", "—"],
            ]),
            col_widths=[3.0, 1.5, 2.0, 1.5],
            subtitle="Bridge from adjusted to run-rate EBITDA"
        )

        return self

    def add_working_capital_analysis(self, data: Dict = None, include_cover: bool = True):
        """
        Add Working Capital analysis slides.

        Args:
            data: Optional dict with NWC data
            include_cover: Whether to add a section header
        """
        data = data or {}

        if include_cover:
            self._add_section_header("Working Capital Analysis")

        # NWC Components
        self._add_table_slide(
            "Working Capital Components",
            ["Component", "LTM ($M)", "FY-1", "FY-2", "Avg", "Days"],
            data.get('nwc_components', [
                ["Accounts Receivable", "$XX", "$XX", "$XX", "$XX", "XX"],
                ["Inventory", "$XX", "$XX", "$XX", "$XX", "XX"],
                ["Prepaid Expenses", "$X", "$X", "$X", "$X", "—"],
                ["Accounts Payable", "($XX)", "($XX)", "($XX)", "($XX)", "XX"],
                ["Accrued Expenses", "($X)", "($X)", "($X)", "($X)", "—"],
                ["Deferred Revenue", "($X)", "($X)", "($X)", "($X)", "—"],
                ["Net Working Capital", "$XX", "$XX", "$XX", "$XX", "—"],
            ]),
            col_widths=[2.5, 1.2, 1.2, 1.2, 1.2, 1.2],
            subtitle="Historical NWC components and days analysis"
        )

        # NWC Peg Analysis
        self._add_content_slide(
            "NWC Peg Mechanism",
            [
                "Target NWC Calculation",
                ("Methodology: Trailing 12-month average", 1),
                ("Target NWC: $XX.XM", 1),
                ("Estimated closing NWC: $XX.XM", 1),
                "",
                "Purchase Agreement Mechanism",
                ("Estimated adjustment: $X.XM [to buyer/seller]", 1),
                ("True-up period: 90 days post-close", 1),
                ("Dispute resolution: Independent accountant", 1),
                "",
                "Key Considerations",
                ("Seasonality impact: Q4 NWC typically XX% higher", 1),
                ("Excluded items: Cash, debt, transaction expenses", 1),
                ("Collar: +/- $X.XM de minimis threshold", 1),
            ],
            subtitle="NWC target and adjustment mechanism"
        )

        return self

    def add_customer_quality_analysis(self, data: Dict = None, include_cover: bool = True):
        """
        Add Customer/Revenue Quality analysis slides.

        Args:
            data: Optional dict with customer data
            include_cover: Whether to add a section header
        """
        data = data or {}

        if include_cover:
            self._add_section_header("Customer & Revenue Quality")

        # Customer Concentration
        self._add_table_slide(
            "Customer Concentration Analysis",
            ["Customer", "Revenue ($M)", "% of Total", "Tenure", "Contract"],
            data.get('customers', [
                ["Customer A", "$XX", "XX%", "X yrs", "Multi-year"],
                ["Customer B", "$XX", "XX%", "X yrs", "Annual"],
                ["Customer C", "$XX", "XX%", "X yrs", "Multi-year"],
                ["Customer D", "$XX", "XX%", "X yrs", "Annual"],
                ["Customer E", "$XX", "XX%", "X yrs", "M-t-M"],
                ["Other (XXX customers)", "$XXX", "XX%", "Avg X yrs", "Various"],
                ["Total", "$XXX", "100%", "—", "—"],
            ]),
            col_widths=[2.5, 1.5, 1.2, 1.2, 1.5],
            subtitle="Top customer analysis and concentration risk"
        )

        # Retention Metrics
        self._add_table_slide(
            "Revenue Retention & Churn",
            ["Metric", "FY-2", "FY-1", "LTM", "Trend"],
            data.get('retention', [
                ["Gross Revenue Retention", "XX%", "XX%", "XX%", "→"],
                ["Net Revenue Retention", "XXX%", "XXX%", "XXX%", "↑"],
                ["Logo Churn Rate", "X%", "X%", "X%", "↓"],
                ["Dollar Churn Rate", "X%", "X%", "X%", "↓"],
            ]),
            col_widths=[3.0, 1.5, 1.5, 1.5, 1.0],
            subtitle="Customer retention and churn analysis"
        )

        # Unit Economics
        self._add_content_slide(
            "Unit Economics Summary",
            [
                "Customer Acquisition",
                ("Customer Acquisition Cost (CAC): $X,XXX", 1),
                ("CAC Payback Period: XX months", 1),
                ("Sales efficiency: X.Xx (LTV/CAC)", 1),
                "",
                "Customer Lifetime Value",
                ("Average Revenue Per User: $XX,XXX", 1),
                ("Gross Margin: XX%", 1),
                ("Annual Churn: X%", 1),
                ("LTV: $XXX,XXX", 1),
                "",
                "Assessment",
                ("LTV/CAC of X.Xx indicates [healthy/marginal] unit economics", 1),
                ("[Revenue quality is strong due to X, Y, Z]", 1),
            ],
            subtitle="Customer acquisition and lifetime value analysis"
        )

        return self

    def add_credit_analysis_slides(self, data: Dict = None, include_cover: bool = True):
        """
        Add Credit/Debt Sizing analysis slides.

        Args:
            data: Optional dict with credit data
            include_cover: Whether to add a section header
        """
        data = data or {}

        if include_cover:
            self._add_section_header("Credit Analysis")

        # Credit Metrics
        self._add_table_slide(
            "Key Credit Ratios",
            ["Ratio", "Current", "Covenant", "Headroom", "Status"],
            data.get('ratios', [
                ["Total Debt / EBITDA", "X.Xx", "X.Xx", "X.Xx", "Pass"],
                ["Net Debt / EBITDA", "X.Xx", "X.Xx", "X.Xx", "Pass"],
                ["EBITDA / Interest", "X.Xx", "X.Xx", "X.Xx", "Pass"],
                ["(EBITDA-CapEx) / Interest", "X.Xx", "X.Xx", "X.Xx", "Pass"],
                ["Fixed Charge Coverage", "X.Xx", "X.Xx", "X.Xx", "Pass"],
            ]),
            col_widths=[2.5, 1.3, 1.3, 1.3, 1.3],
            subtitle="Current leverage and coverage metrics"
        )

        # Debt Capacity
        self._add_table_slide(
            "Debt Capacity Analysis",
            ["Constraint", "Max Debt ($M)", "Implied Multiple", "Binding?"],
            data.get('capacity', [
                ["Leverage Ratio (5.5x)", "$XXX", "5.5x", ""],
                ["Interest Coverage (2.0x)", "$XXX", "X.Xx", ""],
                ["Fixed Charge (1.5x)", "$XXX", "X.Xx", "✓"],
                ["Senior Secured (4.0x)", "$XXX", "4.0x", ""],
            ]),
            col_widths=[3.0, 2.0, 2.0, 1.5],
            subtitle="Maximum debt capacity under various constraints"
        )

        # Stress Test
        self._add_table_slide(
            "EBITDA Stress Test",
            ["Decline", "EBITDA ($M)", "Leverage", "Coverage", "Status"],
            data.get('stress_test', [
                ["Base Case", "$XX", "X.Xx", "X.Xx", "Pass"],
                ["-10%", "$XX", "X.Xx", "X.Xx", "Pass"],
                ["-20%", "$XX", "X.Xx", "X.Xx", "Pass"],
                ["-30%", "$XX", "X.Xx", "X.Xx", "Tight"],
                ["-40%", "$XX", "X.Xx", "X.Xx", "Breach"],
            ]),
            col_widths=[1.5, 1.5, 1.5, 1.5, 1.5],
            subtitle="Covenant compliance under EBITDA stress scenarios"
        )

        return self

    # ============================================================
    # RETURNS & CAPITAL STRUCTURE MODULES
    # ============================================================

    def add_dividend_recap_slides(self, data: Dict = None, include_cover: bool = True):
        """
        Add Dividend Recapitalization analysis slides.

        Args:
            data: Optional dict with recap data
            include_cover: Whether to add a section header
        """
        data = data or {}

        if include_cover:
            self._add_section_header("Dividend Recapitalization")

        # Recap Structure
        self._add_content_slide(
            "Dividend Recap Overview",
            [
                "Transaction Summary",
                ("Timing: Year [X] of hold period", 1),
                ("New debt raised: $XXM", 1),
                ("Net dividend to equity: $XXM", 1),
                "",
                "Capital Structure Impact",
                ("Pre-recap leverage: X.Xx", 1),
                ("Post-recap leverage: X.Xx", 1),
                ("Interest rate on new debt: X.X%", 1),
                "",
                "Rationale",
                ("Return capital to LPs while maintaining ownership", 1),
                ("De-risk investment ahead of exit", 1),
                ("Take advantage of favorable debt markets", 1),
            ],
            subtitle="Mid-hold dividend recapitalization"
        )

        # Returns Impact
        self._add_table_slide(
            "Returns Impact Analysis",
            ["Metric", "Without Recap", "With Recap", "Impact"],
            data.get('returns_impact', [
                ["Initial Equity", "$XXXM", "$XXXM", "—"],
                ["Dividend (Year X)", "—", "$XXM", "+$XXM"],
                ["Exit Equity", "$XXXM", "$XXXM", "($XXM)"],
                ["Total Proceeds", "$XXXM", "$XXXM", "+$XXM"],
                ["MOIC", "X.Xx", "X.Xx", "+X.Xx"],
                ["IRR", "XX%", "XX%", "+X%"],
            ]),
            col_widths=[2.5, 2.0, 2.0, 1.5],
            subtitle="Comparison of returns with and without recap"
        )

        return self

    def add_refinancing_slides(self, data: Dict = None, include_cover: bool = True):
        """
        Add Refinancing Analysis slides.

        Args:
            data: Optional dict with refinancing data
            include_cover: Whether to add a section header
        """
        data = data or {}

        if include_cover:
            self._add_section_header("Refinancing Analysis")

        # Current vs New Structure
        self._add_table_slide(
            "Debt Structure Comparison",
            ["Tranche", "Current ($M)", "Rate", "New ($M)", "Rate"],
            data.get('structure', [
                ["Term Loan A", "$XXX", "X.X%", "—", "—"],
                ["Term Loan B", "$XXX", "X.X%", "$XXX", "X.X%"],
                ["Senior Notes", "$XX", "X.X%", "—", "—"],
                ["Total", "$XXX", "X.X%", "$XXX", "X.X%"],
            ]),
            col_widths=[2.0, 1.5, 1.2, 1.5, 1.2],
            subtitle="Current debt structure vs. refinancing proposal"
        )

        # Savings Analysis
        self._add_content_slide(
            "Refinancing Economics",
            [
                "Annual Savings",
                ("Current interest expense: $X.XM", 1),
                ("New interest expense: $X.XM", 1),
                ("Annual savings: $X.XM", 1),
                ("Rate reduction: XXX bps", 1),
                "",
                "Transaction Costs",
                ("Call premium / make-whole: $X.XM", 1),
                ("Arrangement fee: $X.XM", 1),
                ("Legal & advisory: $X.XM", 1),
                ("Total costs: $X.XM", 1),
                "",
                "Payback Analysis",
                ("Payback period: X.X years", 1),
                ("Recommendation: [Proceed / Wait]", 1),
            ],
            subtitle="Cost-benefit analysis of refinancing"
        )

        return self

    def add_waterfall_slides(self, data: Dict = None, include_cover: bool = True):
        """
        Add Cap Table / Waterfall analysis slides.

        Args:
            data: Optional dict with waterfall data
            include_cover: Whether to add a section header
        """
        data = data or {}

        if include_cover:
            self._add_section_header("Equity Waterfall")

        # Cap Table
        self._add_table_slide(
            "Capitalization Table",
            ["Investor", "Invested ($M)", "Ownership", "Security"],
            data.get('cap_table', [
                ["Sponsor Fund", "$XXX", "XX%", "Common"],
                ["Co-Investors", "$XX", "XX%", "Common"],
                ["Management", "$X", "X%", "Common"],
                ["Rollover Equity", "$X", "X%", "Common"],
                ["Total", "$XXX", "100%", "—"],
            ]),
            col_widths=[2.5, 2.0, 1.5, 2.0],
            subtitle="Ownership structure at close"
        )

        # Waterfall
        self._add_table_slide(
            "Distribution Waterfall by Exit Value",
            ["Component", "$200M", "$300M", "$400M", "$500M"],
            data.get('waterfall', [
                ["Return of Capital", "$XXX", "$XXX", "$XXX", "$XXX"],
                ["Preferred Return", "$XX", "$XX", "$XX", "$XX"],
                ["GP Catch-Up", "$X", "$XX", "$XX", "$XX"],
                ["Remaining (80/20)", "$X", "$XX", "$XX", "$XX"],
                ["Total to LPs", "$XXX", "$XXX", "$XXX", "$XXX"],
                ["Total to GP (Carry)", "$X", "$XX", "$XX", "$XX"],
                ["LP MOIC", "X.Xx", "X.Xx", "X.Xx", "X.Xx"],
            ]),
            col_widths=[2.5, 1.5, 1.5, 1.5, 1.5],
            subtitle="LP and GP distributions at various exit values"
        )

        return self

    def add_sponsor_economics_slides(self, data: Dict = None, include_cover: bool = True):
        """
        Add Sponsor Economics analysis slides.

        Args:
            data: Optional dict with sponsor data
            include_cover: Whether to add a section header
        """
        data = data or {}

        if include_cover:
            self._add_section_header("Sponsor Economics")

        # Deal Economics
        self._add_content_slide(
            "GP Economics - This Deal",
            [
                "Deal Investment",
                ("Deal equity: $XXXM (XX% of fund)", 1),
                ("Expected MOIC: X.Xx", 1),
                ("Expected IRR: XX%", 1),
                "",
                "GP Economics",
                ("Carried interest (20% of profit): $XXM", 1),
                ("GP co-invest return: $X.XM", 1),
                ("Total GP economics: $XXM", 1),
                "",
                "Fund Context",
                ("Fund size: $XXXM", 1),
                ("This deal as % of fund: XX%", 1),
                ("Hurdle rate: X%", 1),
            ],
            subtitle="GP carried interest and co-investment returns"
        )

        # Fund Level
        self._add_table_slide(
            "Fund-Level Economics (Illustrative)",
            ["Component", "Amount ($M)", "% of Fund"],
            data.get('fund_economics', [
                ["Fund Size (Committed)", "$XXX", "1.0x"],
                ["Gross Proceeds", "$X,XXX", "X.Xx"],
                ["Gross Profit", "$XXX", "—"],
                ["", "", ""],
                ["Management Fees (total)", "$XX", "X%"],
                ["Carried Interest (20%)", "$XXX", "XX%"],
                ["GP Co-Invest Profit", "$XX", "X%"],
                ["", "", ""],
                ["Total GP Revenue", "$XXX", "—"],
            ]),
            col_widths=[3.5, 2.0, 1.5],
            subtitle="Illustrative fund-level GP economics"
        )

        return self

    # ============================================================
    # TRANSACTION STRUCTURE MODULES
    # ============================================================

    def add_addon_analysis_slides(self, data: Dict = None, include_cover: bool = True):
        """
        Add Add-on / Bolt-on acquisition analysis slides.

        Args:
            data: Optional dict with add-on data
            include_cover: Whether to add a section header
        """
        data = data or {}

        if include_cover:
            self._add_section_header("Add-on Analysis")

        # Add-on Overview
        self._add_table_slide(
            "Add-on Transaction Summary",
            ["Metric", "Platform", "Add-on", "Combined", "Accretion"],
            data.get('addon_summary', [
                ["Revenue", "$XXX", "$XX", "$XXX", "+XX%"],
                ["EBITDA", "$XX", "$X", "$XX", "+XX%"],
                ["Margin", "XX%", "XX%", "XX%", "+X bps"],
                ["Purchase Price", "—", "$XXM", "—", "—"],
                ["Implied Multiple", "—", "X.Xx", "—", "—"],
            ]),
            col_widths=[2.0, 1.5, 1.5, 1.5, 1.5],
            subtitle="Platform + add-on combination metrics"
        )

        # Combined Financials
        self._add_table_slide(
            "Combined Pro Forma Financials",
            ["Year", "1", "2", "3", "4", "5"],
            data.get('combined_financials', [
                ["Platform Revenue", "$XXX", "$XXX", "$XXX", "$XXX", "$XXX"],
                ["Add-on Revenue", "$XX", "$XX", "$XX", "$XX", "$XX"],
                ["Revenue Synergies", "$X", "$X", "$XX", "$XX", "$XX"],
                ["Combined Revenue", "$XXX", "$XXX", "$XXX", "$XXX", "$XXX"],
                ["Combined EBITDA", "$XX", "$XX", "$XX", "$XX", "$XX"],
                ["% Margin", "XX%", "XX%", "XX%", "XX%", "XX%"],
            ]),
            col_widths=[2.5, 1.3, 1.3, 1.3, 1.3, 1.3],
            subtitle="5-year pro forma with synergies"
        )

        # Accretion Analysis
        self._add_content_slide(
            "Returns Accretion Analysis",
            [
                "Platform-Only Returns",
                ("Entry: X.Xx EBITDA, Exit: X.Xx EBITDA", 1),
                ("MOIC: X.Xx, IRR: XX%", 1),
                "",
                "Platform + Add-on Returns",
                ("Combined entry: X.Xx EBITDA", 1),
                ("Combined exit: X.Xx EBITDA", 1),
                ("MOIC: X.Xx (+X.Xx), IRR: XX% (+X%)", 1),
                "",
                "Key Value Drivers",
                ("Multiple arbitrage: Buy at X.Xx, sell at X.Xx", 1),
                ("Revenue synergies: $XXM run-rate by Year 3", 1),
                ("Cost synergies: $XM from back-office consolidation", 1),
            ],
            subtitle="Impact of add-on acquisition on returns"
        )

        return self

    def add_synergy_slides(self, data: Dict = None, include_cover: bool = True):
        """
        Add Synergy Analysis slides.

        Args:
            data: Optional dict with synergy data
            include_cover: Whether to add a section header
        """
        data = data or {}

        if include_cover:
            self._add_section_header("Synergy Analysis")

        # Synergy Summary
        self._add_table_slide(
            "Synergy Summary",
            ["Category", "Run-Rate ($M)", "% of Target", "Timing", "Confidence"],
            data.get('synergy_summary', [
                ["Revenue Synergies", "", "", "", ""],
                ["  Cross-sell opportunities", "$X.X", "X%", "Year 2-3", "Medium"],
                ["  Pricing optimization", "$X.X", "X%", "Year 1", "High"],
                ["Cost Synergies", "", "", "", ""],
                ["  Headcount reduction", "$X.X", "X%", "Year 1", "High"],
                ["  Procurement savings", "$X.X", "X%", "Year 1-2", "Medium"],
                ["  Facility consolidation", "$X.X", "X%", "Year 2", "High"],
                ["Total Synergies", "$XX.X", "XX%", "—", "—"],
            ]),
            col_widths=[2.8, 1.5, 1.2, 1.2, 1.2],
            subtitle="Run-rate synergy targets by category"
        )

        # Synergy Detail
        self._add_content_slide(
            "Synergy Detail & Assumptions",
            [
                "Revenue Synergies ($X.XM run-rate)",
                ("Cross-sell: X% of target customer base converts", 1),
                ("Price increase: X% on overlapping products", 1),
                ("New product launch: $X.XM incremental revenue", 1),
                "",
                "Cost Synergies ($X.XM run-rate)",
                ("Corporate: X FTEs eliminated ($X.XM)", 1),
                ("Procurement: X% savings on $XXM spend", 1),
                ("IT/Systems: Consolidation to single platform ($X.XM)", 1),
                "",
                "One-Time Costs to Achieve",
                ("Severance: $X.XM (X FTEs × $XXK avg)", 1),
                ("Integration consulting: $X.XM", 1),
                ("Systems migration: $X.XM", 1),
            ],
            subtitle="Key assumptions underlying synergy targets"
        )

        # Synergy Timeline
        self._add_table_slide(
            "Synergy Realization Timeline",
            ["Synergy Type", "Year 1", "Year 2", "Year 3", "Run-Rate"],
            data.get('synergy_timeline', [
                ["Revenue Synergies", "XX%", "XX%", "100%", "$X.XM"],
                ["Cost Synergies", "XX%", "XX%", "100%", "$X.XM"],
                ["Total Synergies", "$X.XM", "$XX.XM", "$XX.XM", "$XX.XM"],
                ["One-Time Costs", "($X.XM)", "($X.XM)", "—", "($X.XM)"],
                ["Net Benefit", "($X.XM)", "$X.XM", "$XX.XM", "—"],
            ]),
            col_widths=[2.5, 1.5, 1.5, 1.5, 1.5],
            subtitle="Phased synergy realization schedule"
        )

        return self

    def add_carveout_slides(self, data: Dict = None, include_cover: bool = True):
        """
        Add Carve-out Analysis slides.

        Args:
            data: Optional dict with carve-out data
            include_cover: Whether to add a section header
        """
        data = data or {}

        if include_cover:
            self._add_section_header("Carve-out Analysis")

        # Carve-out Overview
        self._add_content_slide(
            "Carve-out Transaction Overview",
            [
                "Transaction Structure",
                ("Seller: [Parent Company]", 1),
                ("Target: [Division/Business Unit]", 1),
                ("Revenue: $XXXM (XX% of parent)", 1),
                ("EBITDA: $XXM (XX% margin)", 1),
                "",
                "Key Carve-out Considerations",
                ("Shared services: $X.XM allocated from parent", 1),
                ("Transition services required: XX months", 1),
                ("Stranded costs at parent: $X.XM", 1),
                "",
                "Separation Complexity",
                ("IT systems: [Integrated / Separate]", 1),
                ("Facilities: [Shared / Dedicated]", 1),
                ("Employees: XXX dedicated, XX shared", 1),
            ],
            subtitle="Key elements of the carve-out transaction"
        )

        # Standalone Adjustments
        self._add_table_slide(
            "Standalone Operating Adjustments",
            ["Item", "As Reported ($M)", "Adjustment", "Standalone ($M)"],
            data.get('standalone_adjustments', [
                ["Revenue", "$XXX", "—", "$XXX"],
                ["COGS", "($XX)", "—", "($XX)"],
                ["Gross Profit", "$XX", "—", "$XX"],
                ["Corporate Allocation", "($X)", "+$X", "$0"],
                ["Shared IT Systems", "($X)", "+$X", "($X)"],
                ["Standalone G&A", "—", "+($X)", "($X)"],
                ["Insurance", "($X)", "+($X)", "($X)"],
                ["Public Company Costs", "—", "+($X)", "($X)"],
                ["Standalone EBITDA", "$XX", "($X)", "$XX"],
            ]),
            col_widths=[2.8, 2.0, 1.5, 2.0],
            subtitle="Bridge from reported to standalone economics"
        )

        # Separation Timeline
        self._add_table_slide(
            "Separation & TSA Timeline",
            ["Workstream", "Q1", "Q2", "Q3", "Q4", "Owner"],
            data.get('separation_timeline', [
                ["Legal entity setup", "●", "", "", "", "Legal"],
                ["IT separation", "●", "●", "●", "", "IT"],
                ["HR transition", "●", "●", "", "", "HR"],
                ["Finance standalone", "●", "●", "●", "", "Finance"],
                ["Facility separation", "", "●", "●", "●", "Ops"],
                ["TSA exit", "", "", "●", "●", "PMO"],
            ]),
            col_widths=[2.5, 1.0, 1.0, 1.0, 1.0, 1.5],
            subtitle="Key separation milestones and TSA duration"
        )

        return self

    def add_earnout_slides(self, data: Dict = None, include_cover: bool = True):
        """
        Add Earnout / Contingent Consideration slides.

        Args:
            data: Optional dict with earnout data
            include_cover: Whether to add a section header
        """
        data = data or {}

        if include_cover:
            self._add_section_header("Earnout Analysis")

        # Earnout Structure
        self._add_table_slide(
            "Earnout Structure",
            ["Milestone", "Target", "Earnout ($M)", "Period", "Probability"],
            data.get('earnout_structure', [
                ["Revenue Target Y1", "$XXXM", "$XX", "FY1", "XX%"],
                ["Revenue Target Y2", "$XXXM", "$XX", "FY2", "XX%"],
                ["EBITDA Target Y1", "$XXM", "$XX", "FY1", "XX%"],
                ["EBITDA Target Y2", "$XXM", "$XX", "FY2", "XX%"],
                ["Customer Retention", ">XX%", "$X", "FY1", "XX%"],
                ["Total Earnout", "—", "$XXM", "—", "—"],
            ]),
            col_widths=[2.5, 1.5, 1.5, 1.2, 1.2],
            subtitle="Performance milestones and contingent payments"
        )

        # Milestone Analysis
        self._add_content_slide(
            "Earnout Valuation & Risk Assessment",
            [
                "Probability-Weighted Value",
                ("Total potential earnout: $XXM", 1),
                ("Probability-weighted value: $X.XM", 1),
                ("Discount rate applied: X%", 1),
                ("Fair value at close: $X.XM", 1),
                "",
                "Milestone Achievement Risk",
                ("Revenue targets: Management projects XX% probability", 1),
                ("EBITDA targets: Sensitive to margin assumptions", 1),
                ("Retention target: Historical retention supports achievement", 1),
                "",
                "Key Considerations",
                ("Earnout caps total consideration at $XXXM", 1),
                ("Accounting: Contingent liability on balance sheet", 1),
                ("Management alignment: Seller executives staying X years", 1),
            ],
            subtitle="Fair value assessment of earnout consideration"
        )

        return self

    def add_ppa_slides(self, data: Dict = None, include_cover: bool = True):
        """
        Add Purchase Price Allocation slides.

        Args:
            data: Optional dict with PPA data
            include_cover: Whether to add a section header
        """
        data = data or {}

        if include_cover:
            self._add_section_header("Purchase Price Allocation")

        # PPA Summary
        self._add_table_slide(
            "Purchase Price Allocation Summary",
            ["Asset/Liability", "Book Value ($M)", "Fair Value ($M)", "Step-Up"],
            data.get('ppa_summary', [
                ["Current Assets", "$XX", "$XX", "—"],
                ["PP&E", "$XX", "$XX", "$X"],
                ["Identified Intangibles", "—", "$XX", "$XX"],
                ["  Customer relationships", "—", "$XX", ""],
                ["  Technology/IP", "—", "$X", ""],
                ["  Trade names", "—", "$X", ""],
                ["  Non-compete agreements", "—", "$X", ""],
                ["Goodwill", "—", "$XXX", "$XXX"],
                ["Deferred Tax Liability", "—", "($XX)", "($XX)"],
                ["Total Purchase Price", "—", "$XXX", "—"],
            ]),
            col_widths=[3.0, 2.0, 2.0, 1.5],
            subtitle="Preliminary purchase price allocation"
        )

        # Intangibles Detail
        self._add_table_slide(
            "Identified Intangible Assets",
            ["Intangible", "Fair Value ($M)", "Life (Yrs)", "Annual Amort."],
            data.get('intangibles', [
                ["Customer Relationships", "$XX", "XX", "$X.X"],
                ["Technology / Patents", "$X", "X", "$X.X"],
                ["Trade Names", "$X", "Indef.", "—"],
                ["Non-Compete Agreements", "$X", "X", "$X.X"],
                ["Backlog", "$X", "X", "$X.X"],
                ["Total Amortizing", "$XX", "—", "$X.X"],
            ]),
            col_widths=[3.0, 2.0, 1.5, 2.0],
            subtitle="Intangible asset detail and amortization schedule"
        )

        return self

    # ============================================================
    # VALUE CREATION MODULES
    # ============================================================

    def add_value_creation_slides(self, data: Dict = None, include_cover: bool = True):
        """
        Add Value Creation Bridge slides.

        Args:
            data: Optional dict with value creation data
            include_cover: Whether to add a section header
        """
        data = data or {}

        if include_cover:
            self._add_section_header("Value Creation Analysis")

        # Value Bridge
        self._add_table_slide(
            "Value Creation Bridge",
            ["Component", "Entry Value ($M)", "Exit Value ($M)", "Value Created"],
            data.get('value_bridge', [
                ["Entry Equity Value", "$XXX", "—", "—"],
                ["EBITDA Growth", "—", "$XXX", "+$XXX"],
                ["  Revenue growth contribution", "", "", "(+$XX)"],
                ["  Margin expansion contribution", "", "", "(+$XX)"],
                ["Multiple Expansion", "—", "$XX", "+$XX"],
                ["Debt Paydown", "—", "$XX", "+$XX"],
                ["Exit Equity Value", "—", "$XXX", "—"],
                ["Total Value Created", "—", "—", "$XXX"],
            ]),
            col_widths=[3.0, 2.0, 2.0, 1.5],
            subtitle="Attribution of equity value creation"
        )

        # Value Drivers
        self._add_content_slide(
            "Value Creation Drivers",
            [
                "EBITDA Growth (XX% of value created)",
                ("Revenue CAGR: XX% ($XXM → $XXXM)", 1),
                ("Margin expansion: XX% → XX% (+XXX bps)", 1),
                ("Key initiatives: [pricing, new products, efficiency]", 1),
                "",
                "Multiple Expansion (XX% of value created)",
                ("Entry multiple: X.Xx", 1),
                ("Exit multiple: X.Xx (+X.Xx turn)", 1),
                ("Drivers: Scale, market position, growth profile", 1),
                "",
                "Deleveraging (XX% of value created)",
                ("Entry leverage: X.Xx", 1),
                ("Exit leverage: X.Xx", 1),
                ("Cumulative debt paydown: $XXM", 1),
            ],
            subtitle="Key drivers of equity value creation"
        )

        return self

    def add_hundred_day_slides(self, data: Dict = None, include_cover: bool = True):
        """
        Add 100-Day Plan slides.

        Args:
            data: Optional dict with 100-day plan data
            include_cover: Whether to add a section header
        """
        data = data or {}

        if include_cover:
            self._add_section_header("100-Day Plan")

        # Plan Overview
        self._add_content_slide(
            "100-Day Plan Overview",
            [
                "Phase 1: Days 1-30 (Foundation)",
                ("Finalize management team and reporting structure", 1),
                ("Establish board cadence and KPI reporting", 1),
                ("Complete integration planning (if add-on)", 1),
                "",
                "Phase 2: Days 31-60 (Quick Wins)",
                ("Implement pricing initiatives ($X.XM impact)", 1),
                ("Launch cost reduction program", 1),
                ("Begin commercial excellence review", 1),
                "",
                "Phase 3: Days 61-100 (Strategic Initiatives)",
                ("Finalize strategic plan and budget", 1),
                ("Launch technology/systems roadmap", 1),
                ("Complete organizational design", 1),
            ],
            subtitle="Post-close value creation priorities"
        )

        # Priority Matrix
        self._add_table_slide(
            "Initiative Priority Matrix",
            ["Initiative", "Impact ($M)", "Effort", "Timeline", "Owner"],
            data.get('priorities', [
                ["Pricing optimization", "$X.X", "Low", "Q1", "CCO"],
                ["Procurement savings", "$X.X", "Medium", "Q1-Q2", "COO"],
                ["Sales force effectiveness", "$X.X", "Medium", "Q2", "CCO"],
                ["Working capital optimization", "$X.X", "Low", "Q1", "CFO"],
                ["Headcount rationalization", "$X.X", "High", "Q1", "CHRO"],
                ["System consolidation", "$X.X", "High", "Q2-Q3", "CTO"],
            ]),
            col_widths=[2.8, 1.2, 1.2, 1.2, 1.2],
            subtitle="Prioritized initiatives by impact and effort"
        )

        # Timeline
        self._add_table_slide(
            "100-Day Milestone Timeline",
            ["Milestone", "Day 30", "Day 60", "Day 100", "Status"],
            data.get('milestones', [
                ["Management team in place", "●", "", "", ""],
                ["KPI dashboard live", "●", "", "", ""],
                ["Pricing changes implemented", "", "●", "", ""],
                ["Cost savings identified", "", "●", "", ""],
                ["Strategic plan approved", "", "", "●", ""],
                ["Budget finalized", "", "", "●", ""],
            ]),
            col_widths=[3.0, 1.2, 1.2, 1.2, 1.0],
            subtitle="Key milestones and accountability"
        )

        return self

    def add_exit_readiness_slides(self, data: Dict = None, include_cover: bool = True):
        """
        Add Exit Readiness Assessment slides.

        Args:
            data: Optional dict with exit readiness data
            include_cover: Whether to add a section header
        """
        data = data or {}

        if include_cover:
            self._add_section_header("Exit Readiness")

        # Exit Options
        self._add_table_slide(
            "Exit Options Analysis",
            ["Exit Route", "Probability", "Timing", "Expected Value", "Considerations"],
            data.get('exit_options', [
                ["Strategic Sale", "XX%", "Year X", "$XXX-XXXM", "Synergy premium potential"],
                ["Sponsor-to-Sponsor", "XX%", "Year X", "$XXX-XXXM", "Clean process"],
                ["IPO", "XX%", "Year X+", "$XXX-XXXM", "Market dependent"],
                ["Dividend Recap", "XX%", "Year X", "$XXM dividend", "Partial liquidity"],
            ]),
            col_widths=[2.0, 1.2, 1.0, 1.8, 2.5],
            subtitle="Potential exit routes and expected outcomes"
        )

        # Readiness Scorecard
        self._add_table_slide(
            "Exit Readiness Scorecard",
            ["Dimension", "Current", "Target", "Gap", "Action Required"],
            data.get('scorecard', [
                ["Financial Performance", "●●●○○", "●●●●●", "2", "Continue EBITDA growth"],
                ["Management Team", "●●●●○", "●●●●●", "1", "Hire CFO"],
                ["Systems & Reporting", "●●●○○", "●●●●○", "1", "Implement ERP"],
                ["Customer Concentration", "●●○○○", "●●●●○", "2", "Diversify top 5"],
                ["Growth Story", "●●●●○", "●●●●●", "1", "Execute M&A pipeline"],
                ["Market Position", "●●●●○", "●●●●●", "1", "Maintain share gains"],
            ]),
            col_widths=[2.2, 1.3, 1.3, 0.8, 2.8],
            subtitle="Assessment of key exit readiness dimensions"
        )

        return self

    def add_mip_slides(self, data: Dict = None, include_cover: bool = True):
        """
        Add Management Incentive Plan (MIP) slides.

        Args:
            data: Optional dict with MIP data
            include_cover: Whether to add a section header
        """
        data = data or {}

        if include_cover:
            self._add_section_header("Management Incentive Plan")

        # MIP Structure
        self._add_content_slide(
            "MIP Structure Overview",
            [
                "Pool Size & Allocation",
                ("Total MIP pool: XX% of fully diluted equity", 1),
                ("CEO allocation: XX%", 1),
                ("CFO allocation: X%", 1),
                ("Other executives: X%", 1),
                ("Future hires reserve: X%", 1),
                "",
                "Vesting Terms",
                ("Time vesting: XX% over X years (cliff + monthly)", 1),
                ("Performance vesting: XX% tied to exit returns", 1),
                ("Acceleration: Double trigger on CIC", 1),
                "",
                "Key Terms",
                ("Strike price: Fair market value at grant ($X.XX/share)", 1),
                ("Instrument: [Options / Profits Interests / RSUs]", 1),
            ],
            subtitle="Equity incentive structure for management team"
        )

        # Payout Scenarios
        self._add_table_slide(
            "MIP Payout Scenarios",
            ["Exit Value", "MOIC", "MIP Pool Value", "CEO Payout", "% of Exit"],
            data.get('mip_payouts', [
                ["$200M", "1.5x", "$X.XM", "$X.XM", "X.X%"],
                ["$300M", "2.0x", "$X.XM", "$X.XM", "X.X%"],
                ["$400M", "2.5x", "$XX.XM", "$X.XM", "X.X%"],
                ["$500M", "3.0x", "$XX.XM", "$XX.XM", "X.X%"],
                ["$600M", "3.5x", "$XX.XM", "$XX.XM", "X.X%"],
            ]),
            col_widths=[1.8, 1.2, 1.8, 1.8, 1.4],
            subtitle="Management equity value at various exit scenarios"
        )

        return self

    # ============================================================
    # SPECIALIZED MODULES
    # ============================================================

    def add_rollup_slides(self, data: Dict = None, include_cover: bool = True):
        """
        Add Rollup / Platform Build slides.

        Args:
            data: Optional dict with rollup data
            include_cover: Whether to add a section header
        """
        data = data or {}

        if include_cover:
            self._add_section_header("Rollup Strategy")

        # Rollup Summary
        self._add_table_slide(
            "Rollup Acquisition Summary",
            ["Acquisition", "Close Date", "Revenue ($M)", "EBITDA ($M)", "Multiple"],
            data.get('rollup_summary', [
                ["Platform", "Year 0", "$XXX", "$XX", "X.Xx"],
                ["Add-on #1", "Year 1", "$XX", "$X", "X.Xx"],
                ["Add-on #2", "Year 2", "$XX", "$X", "X.Xx"],
                ["Add-on #3", "Year 3", "$XX", "$X", "X.Xx"],
                ["Combined", "—", "$XXX", "$XX", "X.Xx (blended)"],
            ]),
            col_widths=[2.0, 1.2, 1.5, 1.5, 1.5],
            subtitle="Platform + add-on acquisition history"
        )

        # Acquisition Timeline
        self._add_content_slide(
            "Rollup Strategy & Pipeline",
            [
                "Strategy Overview",
                ("Target market: [description]", 1),
                ("Fragmentation: Top X players control XX% of market", 1),
                ("Typical target profile: $X-XXM revenue, XX%+ margins", 1),
                "",
                "Acquisition Criteria",
                ("Geography: [Regions of focus]", 1),
                ("Valuation: X.Xx - X.Xx EBITDA", 1),
                ("Synergy potential: Back-office, procurement", 1),
                "",
                "Pipeline Status",
                ("Active discussions: X targets ($XXM combined revenue)", 1),
                ("LOI stage: X targets", 1),
                ("Dry powder available: $XXM", 1),
            ],
            subtitle="Acquisition strategy and current pipeline"
        )

        # Combined Metrics
        self._add_table_slide(
            "Combined Platform Metrics",
            ["Metric", "Platform Only", "With Add-ons", "Target (Exit)"],
            data.get('combined_metrics', [
                ["Revenue", "$XXX", "$XXX", "$XXX"],
                ["Revenue CAGR", "X%", "XX%", "XX%"],
                ["EBITDA", "$XX", "$XX", "$XX"],
                ["EBITDA Margin", "XX%", "XX%", "XX%"],
                ["Leverage", "X.Xx", "X.Xx", "X.Xx"],
                ["Blended Entry Multiple", "X.Xx", "X.Xx", "—"],
                ["Exit Multiple Target", "—", "—", "X.Xx"],
            ]),
            col_widths=[2.8, 1.8, 1.8, 1.8],
            subtitle="Platform value creation through rollup"
        )

        return self

    def add_tax_slides(self, data: Dict = None, include_cover: bool = True):
        """
        Add Tax Analysis slides.

        Args:
            data: Optional dict with tax data
            include_cover: Whether to add a section header
        """
        data = data or {}

        if include_cover:
            self._add_section_header("Tax Analysis")

        # Tax Overview
        self._add_content_slide(
            "Tax Structure Overview",
            [
                "Transaction Tax Structure",
                ("Deal structure: [Asset / Stock purchase]", 1),
                ("Tax jurisdiction: [US / Canada / Other]", 1),
                ("Step-up available: [Yes / No / Partial]", 1),
                "",
                "Key Tax Attributes",
                ("Net Operating Losses: $XXM (expires 20XX)", 1),
                ("Section 382 limitation: $X.XM annually", 1),
                ("Tax credits available: $X.XM", 1),
                "",
                "Effective Tax Rate Projection",
                ("Statutory rate: XX%", 1),
                ("Estimated effective rate: XX%", 1),
                ("Drivers: [NOL utilization, R&D credits, etc.]", 1),
            ],
            subtitle="Tax considerations and structure"
        )

        # Tax Shield Analysis
        self._add_table_slide(
            "Tax Shield Value Analysis",
            ["Year", "1", "2", "3", "4", "5", "Total"],
            data.get('tax_shield', [
                ["Amortization (step-up)", "$X", "$X", "$X", "$X", "$X", "$XX"],
                ["Interest deduction", "$X", "$X", "$X", "$X", "$X", "$XX"],
                ["NOL utilization", "$X", "$X", "$X", "$X", "$X", "$XX"],
                ["Total deductions", "$XX", "$XX", "$XX", "$XX", "$XX", "$XXX"],
                ["Tax shield (XX%)", "$X", "$X", "$X", "$X", "$X", "$XX"],
                ["PV of tax shield", "—", "—", "—", "—", "—", "$XX"],
            ]),
            col_widths=[2.2, 1.0, 1.0, 1.0, 1.0, 1.0, 1.2],
            subtitle="Present value of tax attributes"
        )

        return self

    def add_control_premium_slides(self, data: Dict = None, include_cover: bool = True):
        """
        Add Control Premium Analysis slides.

        Args:
            data: Optional dict with control premium data
            include_cover: Whether to add a section header
        """
        data = data or {}

        if include_cover:
            self._add_section_header("Control Premium Analysis")

        # Premium Analysis
        self._add_table_slide(
            "Offer Premium Analysis",
            ["Metric", "Unaffected Price", "Offer Price", "Premium"],
            data.get('premium_analysis', [
                ["Share Price", "$XX.XX", "$XX.XX", "XX%"],
                ["30-Day VWAP", "$XX.XX", "—", "XX%"],
                ["52-Week High", "$XX.XX", "—", "XX%"],
                ["Equity Value", "$XXXM", "$XXXM", "XX%"],
                ["Enterprise Value", "$XXXM", "$XXXM", "XX%"],
                ["EV/EBITDA", "X.Xx", "X.Xx", "+X.Xx turn"],
            ]),
            col_widths=[2.5, 2.0, 2.0, 1.5],
            subtitle="Offer premium relative to unaffected price"
        )

        # Precedent Premiums
        self._add_table_slide(
            "Precedent Transaction Premiums",
            ["Transaction", "Date", "Deal Value", "1-Day Premium", "30-Day Premium"],
            data.get('precedent_premiums', [
                ["Comparable Deal 1", "20XX", "$X.XB", "XX%", "XX%"],
                ["Comparable Deal 2", "20XX", "$X.XB", "XX%", "XX%"],
                ["Comparable Deal 3", "20XX", "$XXX M", "XX%", "XX%"],
                ["Comparable Deal 4", "20XX", "$XXX M", "XX%", "XX%"],
                ["Comparable Deal 5", "20XX", "$XXX M", "XX%", "XX%"],
                ["Median", "—", "—", "XX%", "XX%"],
                ["This Transaction", "—", "$XXXM", "XX%", "XX%"],
            ]),
            col_widths=[2.5, 1.0, 1.5, 1.5, 1.5],
            subtitle="Control premiums paid in comparable transactions"
        )

        return self

    # ============================================================
    # FULL DECK GENERATION
    # ============================================================

    def add_full_deck(self, data: Dict = None):
        """
        Generate a complete investment memo deck with all modules.

        Args:
            data: Optional dict with all data for population
        """
        data = data or {}

        # Cover
        self._add_cover_slide(
            f"{self.company_name} Investment Analysis",
            "Private Equity Case Study"
        )

        # Table of Contents
        self._add_content_slide(
            "Table of Contents",
            [
                "1. Executive Summary",
                "2. Company Overview",
                "3. Industry Analysis",
                "4. Competitive Positioning",
                "5. Financial Analysis",
                "6. Valuation",
                "7. Transaction Analysis",
                "8. Investment Thesis & Risks",
            ]
        )

        # Add all modules
        self.add_company_overview(data.get('company'), include_cover=True)
        self.add_industry_analysis(data.get('industry'), include_cover=True)
        self.add_competitive_analysis(data.get('competitive'), include_cover=True)
        self.add_financial_analysis(data.get('financial'), include_cover=True)
        self.add_valuation(data.get('valuation'), include_cover=True)
        self.add_lbo_analysis(data.get('lbo'), include_cover=True)
        self.add_investment_thesis(data.get('thesis'), include_cover=True)

        return self

    # ============================================================
    # CFA VISUAL SLIDES (Research Challenge Patterns)
    # ============================================================

    def add_wmbt_slide(self, data: Dict = None, include_cover: bool = True):
        """
        Add "What Must Be True" (WMBT) slide — investment thesis conditions
        with thresholds, current status, and consequences.

        Args:
            data: Optional dict with keys: conditions (list of dicts with
                  condition, threshold, current, status, consequence)
            include_cover: Whether to add a section header
        """
        data = data or {}

        if include_cover:
            self._add_section_header("What Must Be True")

        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        page = self._next_page()

        # Title
        txBox = slide.shapes.add_textbox(Inches(0.3), Inches(0.2), Inches(9.5), Inches(0.3))
        p = txBox.text_frame.paragraphs[0]
        p.text = f"What Must Be True — {self.company_name}"
        p.font.name = FONT_NAME
        p.font.size = Pt(16)
        p.font.bold = True
        p.font.color.rgb = NAVY

        # Subtitle
        txBox_sub = slide.shapes.add_textbox(Inches(0.3), Inches(0.6), Inches(9.5), Inches(0.3))
        p_sub = txBox_sub.text_frame.paragraphs[0]
        p_sub.text = "Key conditions that must hold for the investment thesis to succeed"
        p_sub.font.name = FONT_NAME
        p_sub.font.size = Pt(11)
        p_sub.font.color.rgb = DARK_GRAY

        conditions = data.get('conditions', [
            {"condition": "Revenue growth sustains >10%", "threshold": ">10%", "current": "12.3%", "status": "PASS", "consequence": "Terminal value drops 20-30%"},
            {"condition": "EBITDA margins expand to 25%+", "threshold": ">25%", "current": "23.1%", "status": "AT RISK", "consequence": "Returns fall below 15% IRR hurdle"},
            {"condition": "Working capital stable at 8-10% of rev", "threshold": "<10%", "current": "9.2%", "status": "PASS", "consequence": "FCF conversion deteriorates"},
            {"condition": "No customer concentration >20%", "threshold": "<20%", "current": "18.5%", "status": "AT RISK", "consequence": "Revenue vulnerability; lower multiple"},
            {"condition": "Management team retained post-close", "threshold": "Key 5 retained", "current": "4 of 5 confirmed", "status": "PASS", "consequence": "Execution risk; transition cost $2-5M"},
            {"condition": "Debt covenants maintained", "threshold": "<4.5x leverage", "current": "3.8x", "status": "PASS", "consequence": "Covenant breach triggers acceleration"},
        ])

        # Table: Condition | Threshold | Current | Status | Consequence
        headers = ["Condition", "Threshold", "Current", "Status", "If False..."]
        num_cols = 5
        num_rows = len(conditions) + 1
        col_widths_raw = [3.0, 1.2, 1.0, 0.9, 3.3]
        col_widths = [Inches(w) for w in col_widths_raw]

        table = slide.shapes.add_table(num_rows, num_cols, Inches(0.3), Inches(1.1),
                                       sum(col_widths), Inches(0.38 * num_rows)).table

        for i, width in enumerate(col_widths):
            table.columns[i].width = width

        # Header row
        for i, header in enumerate(headers):
            cell = table.cell(0, i)
            cell.text = header
            cell.fill.solid()
            cell.fill.fore_color.rgb = NAVY
            p = cell.text_frame.paragraphs[0]
            p.font.name = FONT_NAME
            p.font.size = Pt(9)
            p.font.bold = True
            p.font.color.rgb = WHITE
            p.alignment = PP_ALIGN.CENTER

        # Data rows with status color coding
        STATUS_COLORS = {
            "PASS": RGBColor(198, 239, 206),      # Light green
            "AT RISK": RGBColor(255, 235, 156),    # Light amber
            "FAIL": RGBColor(255, 199, 206),       # Light red
        }

        for row_idx, cond in enumerate(conditions):
            values = [cond.get('condition', ''), cond.get('threshold', ''),
                      cond.get('current', ''), cond.get('status', ''),
                      cond.get('consequence', '')]
            for col_idx, val in enumerate(values):
                cell = table.cell(row_idx + 1, col_idx)
                cell.text = str(val)
                p = cell.text_frame.paragraphs[0]
                p.font.name = FONT_NAME
                p.font.size = Pt(9)
                p.font.color.rgb = DARK_GRAY
                p.alignment = PP_ALIGN.LEFT if col_idx in (0, 4) else PP_ALIGN.CENTER
                if col_idx == 0:
                    p.font.bold = True

                # Color code status column
                if col_idx == 3:
                    status = str(val).upper()
                    if status in STATUS_COLORS:
                        cell.fill.solid()
                        cell.fill.fore_color.rgb = STATUS_COLORS[status]
                        p.font.bold = True

        # Page number
        txBox_pg = slide.shapes.add_textbox(Inches(9.0), Inches(7.0), Inches(0.8), Inches(0.3))
        p_pg = txBox_pg.text_frame.paragraphs[0]
        p_pg.text = str(page)
        p_pg.font.name = FONT_NAME
        p_pg.font.size = Pt(10)
        p_pg.font.color.rgb = DARK_GRAY

        # Source
        txBox_src = slide.shapes.add_textbox(Inches(0.3), Inches(7.0), Inches(8.5), Inches(0.3))
        p_src = txBox_src.text_frame.paragraphs[0]
        p_src.text = f"Source: {data.get('source', 'Management projections, buyer due diligence')}"
        p_src.font.name = FONT_NAME
        p_src.font.size = Pt(8)
        p_src.font.color.rgb = DARK_GRAY

        return self

    def add_risk_mitigant_slide(self, data: Dict = None, include_cover: bool = True):
        """
        Add Risk-Mitigant Pairing slide — risks matched with mitigants and residual risk.

        Args:
            data: Optional dict with keys: risks (list of dicts with
                  risk, category, probability, impact, mitigant, residual)
            include_cover: Whether to add a section header
        """
        data = data or {}

        if include_cover:
            self._add_section_header("Risk Assessment")

        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        page = self._next_page()

        # Title
        txBox = slide.shapes.add_textbox(Inches(0.3), Inches(0.2), Inches(9.5), Inches(0.3))
        p = txBox.text_frame.paragraphs[0]
        p.text = f"Risk-Mitigant Pairing — {self.company_name}"
        p.font.name = FONT_NAME
        p.font.size = Pt(16)
        p.font.bold = True
        p.font.color.rgb = NAVY

        risks = data.get('risks', [
            {"risk": "Customer concentration (top 3 = 45%)", "category": "Revenue", "probability": "Medium", "impact": "High", "mitigant": "Diversification plan; no single >20%", "residual": "Medium"},
            {"risk": "Key person dependency (CEO/CTO)", "category": "Operational", "probability": "Low", "impact": "High", "mitigant": "Retention packages; succession plan", "residual": "Low"},
            {"risk": "Technology obsolescence", "category": "Strategic", "probability": "Medium", "impact": "Medium", "mitigant": "R&D investment >5% of revenue", "residual": "Low"},
            {"risk": "Regulatory change (data privacy)", "category": "Regulatory", "probability": "Medium", "impact": "Medium", "mitigant": "Compliance program; GDPR-ready", "residual": "Low"},
            {"risk": "Integration risk (add-on strategy)", "category": "Execution", "probability": "High", "impact": "Medium", "mitigant": "Dedicated integration PMO", "residual": "Medium"},
            {"risk": "Interest rate increase", "category": "Financial", "probability": "Low", "impact": "Low", "mitigant": "Fixed-rate tranches; hedging", "residual": "Low"},
        ])

        # Table
        headers = ["Risk", "Category", "Prob.", "Impact", "Mitigant", "Residual"]
        num_cols = 6
        num_rows = len(risks) + 1
        col_widths_raw = [2.5, 0.9, 0.7, 0.7, 2.7, 0.8]
        col_widths = [Inches(w) for w in col_widths_raw]

        table = slide.shapes.add_table(num_rows, num_cols, Inches(0.3), Inches(0.8),
                                       sum(col_widths), Inches(0.38 * num_rows)).table

        for i, width in enumerate(col_widths):
            table.columns[i].width = width

        # Header
        for i, header in enumerate(headers):
            cell = table.cell(0, i)
            cell.text = header
            cell.fill.solid()
            cell.fill.fore_color.rgb = NAVY
            p = cell.text_frame.paragraphs[0]
            p.font.name = FONT_NAME
            p.font.size = Pt(9)
            p.font.bold = True
            p.font.color.rgb = WHITE
            p.alignment = PP_ALIGN.CENTER

        # Color map for probability/impact/residual
        LEVEL_COLORS = {
            "High": RGBColor(255, 199, 206),
            "Medium": RGBColor(255, 235, 156),
            "Low": RGBColor(198, 239, 206),
        }

        for row_idx, risk in enumerate(risks):
            values = [risk.get('risk', ''), risk.get('category', ''),
                      risk.get('probability', ''), risk.get('impact', ''),
                      risk.get('mitigant', ''), risk.get('residual', '')]
            for col_idx, val in enumerate(values):
                cell = table.cell(row_idx + 1, col_idx)
                cell.text = str(val)
                p = cell.text_frame.paragraphs[0]
                p.font.name = FONT_NAME
                p.font.size = Pt(9)
                p.font.color.rgb = DARK_GRAY
                p.alignment = PP_ALIGN.LEFT if col_idx in (0, 4) else PP_ALIGN.CENTER
                if col_idx == 0:
                    p.font.bold = True

                # Color code probability, impact, residual columns
                if col_idx in (2, 3, 5):
                    level = str(val).strip().title()
                    if level in LEVEL_COLORS:
                        cell.fill.solid()
                        cell.fill.fore_color.rgb = LEVEL_COLORS[level]

        # Page number
        txBox_pg = slide.shapes.add_textbox(Inches(9.0), Inches(7.0), Inches(0.8), Inches(0.3))
        p_pg = txBox_pg.text_frame.paragraphs[0]
        p_pg.text = str(page)
        p_pg.font.name = FONT_NAME
        p_pg.font.size = Pt(10)
        p_pg.font.color.rgb = DARK_GRAY

        # Source
        txBox_src = slide.shapes.add_textbox(Inches(0.3), Inches(7.0), Inches(8.5), Inches(0.3))
        p_src = txBox_src.text_frame.paragraphs[0]
        p_src.text = f"Source: {data.get('source', 'Due diligence findings, management discussions')}"
        p_src.font.name = FONT_NAME
        p_src.font.size = Pt(8)
        p_src.font.color.rgb = DARK_GRAY

        return self

    def add_valuation_blending_slide(self, data: Dict = None, include_cover: bool = True):
        """
        Add Valuation Blending slide — weighted valuation across methods with blended target.

        Args:
            data: Optional dict with keys: methods (list of dicts with method, low, mid, high, weight),
                  target_price, current_price
            include_cover: Whether to add a section header
        """
        data = data or {}

        if include_cover:
            self._add_section_header("Valuation Summary")

        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        page = self._next_page()

        # Title
        txBox = slide.shapes.add_textbox(Inches(0.3), Inches(0.2), Inches(9.5), Inches(0.3))
        p = txBox.text_frame.paragraphs[0]
        p.text = f"Blended Valuation — {self.company_name}"
        p.font.name = FONT_NAME
        p.font.size = Pt(16)
        p.font.bold = True
        p.font.color.rgb = NAVY

        # Subtitle
        txBox_sub = slide.shapes.add_textbox(Inches(0.3), Inches(0.6), Inches(9.5), Inches(0.3))
        p_sub = txBox_sub.text_frame.paragraphs[0]
        p_sub.text = "Probability-weighted valuation across methodologies"
        p_sub.font.name = FONT_NAME
        p_sub.font.size = Pt(11)
        p_sub.font.color.rgb = DARK_GRAY

        methods = data.get('methods', [
            {"method": "DCF (Base Case)", "low": "$42.00", "mid": "$48.50", "high": "$55.00", "weight": "35%"},
            {"method": "Trading Comps (EV/EBITDA)", "low": "$40.00", "mid": "$46.00", "high": "$52.00", "weight": "25%"},
            {"method": "Transaction Comps", "low": "$44.00", "mid": "$50.00", "high": "$56.00", "weight": "20%"},
            {"method": "Dividend Discount Model", "low": "$38.00", "mid": "$44.00", "high": "$50.00", "weight": "10%"},
            {"method": "Sum-of-the-Parts", "low": "$45.00", "mid": "$51.00", "high": "$57.00", "weight": "10%"},
        ])

        # Method table
        headers = ["Methodology", "Low", "Mid", "High", "Weight"]
        num_cols = 5
        num_rows = len(methods) + 2  # +1 header, +1 blended row
        col_widths_raw = [3.0, 1.3, 1.3, 1.3, 1.0]
        col_widths = [Inches(w) for w in col_widths_raw]

        table = slide.shapes.add_table(num_rows, num_cols, Inches(0.5), Inches(1.1),
                                       sum(col_widths), Inches(0.38 * num_rows)).table

        for i, width in enumerate(col_widths):
            table.columns[i].width = width

        # Header
        for i, header in enumerate(headers):
            cell = table.cell(0, i)
            cell.text = header
            cell.fill.solid()
            cell.fill.fore_color.rgb = NAVY
            p = cell.text_frame.paragraphs[0]
            p.font.name = FONT_NAME
            p.font.size = Pt(10)
            p.font.bold = True
            p.font.color.rgb = WHITE
            p.alignment = PP_ALIGN.CENTER

        # Data rows
        for row_idx, method in enumerate(methods):
            values = [method.get('method', ''), method.get('low', ''),
                      method.get('mid', ''), method.get('high', ''),
                      method.get('weight', '')]
            for col_idx, val in enumerate(values):
                cell = table.cell(row_idx + 1, col_idx)
                cell.text = str(val)
                p = cell.text_frame.paragraphs[0]
                p.font.name = FONT_NAME
                p.font.size = Pt(10)
                p.font.color.rgb = DARK_GRAY
                p.alignment = PP_ALIGN.LEFT if col_idx == 0 else PP_ALIGN.CENTER
                if col_idx == 0:
                    p.font.bold = True

        # Blended row (last row)
        blended_values = [
            "Blended Valuation",
            data.get('blended_low', '$42.10'),
            data.get('blended_mid', '$48.05'),
            data.get('blended_high', '$54.10'),
            "100%",
        ]
        for col_idx, val in enumerate(blended_values):
            cell = table.cell(num_rows - 1, col_idx)
            cell.text = str(val)
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(230, 240, 250)
            p = cell.text_frame.paragraphs[0]
            p.font.name = FONT_NAME
            p.font.size = Pt(10)
            p.font.bold = True
            p.font.color.rgb = NAVY
            p.alignment = PP_ALIGN.LEFT if col_idx == 0 else PP_ALIGN.CENTER

        # Equation bar below table
        eq_top = 1.1 + 0.38 * num_rows + 0.3
        eq_box = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(eq_top), Inches(8.9), Inches(0.8))
        eq_box.fill.solid()
        eq_box.fill.fore_color.rgb = RGBColor(248, 250, 252)
        eq_box.line.color.rgb = NAVY
        eq_box.line.width = Pt(1.5)

        txBox_eq = slide.shapes.add_textbox(Inches(0.7), Inches(eq_top + 0.1), Inches(8.5), Inches(0.6))
        tf_eq = txBox_eq.text_frame
        tf_eq.word_wrap = True
        p_eq = tf_eq.paragraphs[0]
        p_eq.text = f"Target Price = {data.get('blended_mid', '$48.05')}  |  Current Price: {data.get('current_price', '$41.25')}  |  Upside: {data.get('upside', '16.5%')}"
        p_eq.font.name = FONT_NAME
        p_eq.font.size = Pt(12)
        p_eq.font.bold = True
        p_eq.font.color.rgb = NAVY
        p_eq.alignment = PP_ALIGN.CENTER

        p2 = tf_eq.add_paragraph()
        p2.text = data.get('recommendation', 'BUY — Target price implies 16.5% upside with favorable risk/reward')
        p2.font.name = FONT_NAME
        p2.font.size = Pt(10)
        p2.font.color.rgb = ACCENT_GREEN
        p2.font.bold = True
        p2.alignment = PP_ALIGN.CENTER

        # Page number
        txBox_pg = slide.shapes.add_textbox(Inches(9.0), Inches(7.0), Inches(0.8), Inches(0.3))
        p_pg = txBox_pg.text_frame.paragraphs[0]
        p_pg.text = str(page)
        p_pg.font.name = FONT_NAME
        p_pg.font.size = Pt(10)
        p_pg.font.color.rgb = DARK_GRAY

        return self

    def add_value_chain_slide(self, data: Dict = None, include_cover: bool = True):
        """
        Add Value Chain slide — horizontal flow with boxes and arrows,
        company position highlighted.

        Args:
            data: Optional dict with keys: stages (list of dicts with name, description,
                  margin), highlight_index (int, 0-based index of company's position)
            include_cover: Whether to add a section header
        """
        data = data or {}

        if include_cover:
            self._add_section_header("Value Chain Analysis")

        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        page = self._next_page()

        # Title
        txBox = slide.shapes.add_textbox(Inches(0.3), Inches(0.2), Inches(9.5), Inches(0.3))
        p = txBox.text_frame.paragraphs[0]
        p.text = f"Industry Value Chain — {self.company_name}"
        p.font.name = FONT_NAME
        p.font.size = Pt(16)
        p.font.bold = True
        p.font.color.rgb = NAVY

        stages = data.get('stages', [
            {"name": "Raw Materials", "description": "Commodity inputs, mining, extraction", "margin": "8-12%"},
            {"name": "Components", "description": "Processing, sub-assembly, specialty parts", "margin": "15-20%"},
            {"name": "Manufacturing", "description": "Assembly, quality control, packaging", "margin": "12-18%"},
            {"name": "Distribution", "description": "Logistics, warehousing, fulfillment", "margin": "5-8%"},
            {"name": "End Market", "description": "Retail, B2B sales, service & support", "margin": "20-30%"},
        ])
        highlight_idx = data.get('highlight_index', 2)  # Company position

        num_stages = len(stages)
        total_width = 8.8
        gap = 0.15
        arrow_width = 0.25
        box_width = (total_width - (num_stages - 1) * (gap + arrow_width)) / num_stages
        box_height = 1.8
        box_top = 1.5
        start_left = 0.6

        for i, stage in enumerate(stages):
            left = start_left + i * (box_width + gap + arrow_width)
            is_highlight = (i == highlight_idx)

            # Box
            box = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                Inches(left), Inches(box_top), Inches(box_width), Inches(box_height))

            if is_highlight:
                box.fill.solid()
                box.fill.fore_color.rgb = NAVY
                text_color = WHITE
                desc_color = RGBColor(200, 220, 240)
            else:
                box.fill.solid()
                box.fill.fore_color.rgb = RGBColor(248, 250, 252)
                box.line.color.rgb = NAVY
                box.line.width = Pt(1)
                text_color = NAVY
                desc_color = DARK_GRAY

            # Stage name
            txBox_name = slide.shapes.add_textbox(
                Inches(left + 0.08), Inches(box_top + 0.1),
                Inches(box_width - 0.16), Inches(0.4))
            p_name = txBox_name.text_frame.paragraphs[0]
            p_name.text = stage.get('name', '')
            p_name.font.name = FONT_NAME
            p_name.font.size = Pt(9)
            p_name.font.bold = True
            p_name.font.color.rgb = text_color
            p_name.alignment = PP_ALIGN.CENTER

            # Description
            txBox_desc = slide.shapes.add_textbox(
                Inches(left + 0.08), Inches(box_top + 0.5),
                Inches(box_width - 0.16), Inches(0.8))
            tf_desc = txBox_desc.text_frame
            tf_desc.word_wrap = True
            p_desc = tf_desc.paragraphs[0]
            p_desc.text = stage.get('description', '')
            p_desc.font.name = FONT_NAME
            p_desc.font.size = Pt(8)
            p_desc.font.color.rgb = desc_color

            # Margin
            txBox_margin = slide.shapes.add_textbox(
                Inches(left + 0.08), Inches(box_top + box_height - 0.4),
                Inches(box_width - 0.16), Inches(0.3))
            p_margin = txBox_margin.text_frame.paragraphs[0]
            p_margin.text = f"Margin: {stage.get('margin', 'N/A')}"
            p_margin.font.name = FONT_NAME
            p_margin.font.size = Pt(8)
            p_margin.font.bold = True
            p_margin.font.color.rgb = text_color
            p_margin.alignment = PP_ALIGN.CENTER

            # Arrow (except after last box)
            if i < num_stages - 1:
                arrow_left = left + box_width + gap / 2
                arrow = slide.shapes.add_shape(
                    MSO_SHAPE.RIGHT_ARROW,
                    Inches(arrow_left), Inches(box_top + box_height / 2 - 0.15),
                    Inches(arrow_width), Inches(0.3))
                arrow.fill.solid()
                arrow.fill.fore_color.rgb = NAVY
                arrow.line.fill.background()

        # Company position label
        highlight_left = start_left + highlight_idx * (box_width + gap + arrow_width)
        txBox_label = slide.shapes.add_textbox(
            Inches(highlight_left), Inches(box_top + box_height + 0.15),
            Inches(box_width), Inches(0.3))
        p_label = txBox_label.text_frame.paragraphs[0]
        p_label.text = f"{self.company_name}"
        p_label.font.name = FONT_NAME
        p_label.font.size = Pt(9)
        p_label.font.bold = True
        p_label.font.color.rgb = NAVY
        p_label.alignment = PP_ALIGN.CENTER

        # Key takeaways section
        takeaways = data.get('takeaways', [
            "Company operates in the highest-margin segment of the value chain",
            "Vertical integration opportunity upstream could improve margin by 200-400bps",
            "Limited substitution risk due to proprietary technology and switching costs",
        ])

        txBox_take = slide.shapes.add_textbox(Inches(0.5), Inches(4.2), Inches(9.0), Inches(2.5))
        tf_take = txBox_take.text_frame
        tf_take.word_wrap = True
        p_header = tf_take.paragraphs[0]
        p_header.text = "Key Takeaways"
        p_header.font.name = FONT_NAME
        p_header.font.size = Pt(11)
        p_header.font.bold = True
        p_header.font.color.rgb = NAVY

        for takeaway in takeaways:
            p_t = tf_take.add_paragraph()
            p_t.text = f"• {takeaway}"
            p_t.font.name = FONT_NAME
            p_t.font.size = Pt(9)
            p_t.font.color.rgb = DARK_GRAY
            p_t.space_after = Pt(4)

        # Page number
        txBox_pg = slide.shapes.add_textbox(Inches(9.0), Inches(7.0), Inches(0.8), Inches(0.3))
        p_pg = txBox_pg.text_frame.paragraphs[0]
        p_pg.text = str(page)
        p_pg.font.name = FONT_NAME
        p_pg.font.size = Pt(10)
        p_pg.font.color.rgb = DARK_GRAY

        return self

    def add_risk_matrix_slide(self, data: Dict = None, include_cover: bool = True):
        """
        Add Risk Matrix slide — 3x3 probability-impact grid with color-coded cells.

        Args:
            data: Optional dict with keys: risks (list of dicts with name, probability (1-3),
                  impact (1-3)), labels optional
            include_cover: Whether to add a section header
        """
        data = data or {}

        if include_cover:
            self._add_section_header("Risk Assessment")

        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        page = self._next_page()

        # Title
        txBox = slide.shapes.add_textbox(Inches(0.3), Inches(0.2), Inches(9.5), Inches(0.3))
        p = txBox.text_frame.paragraphs[0]
        p.text = f"Risk Matrix — {self.company_name}"
        p.font.name = FONT_NAME
        p.font.size = Pt(16)
        p.font.bold = True
        p.font.color.rgb = NAVY

        risks = data.get('risks', [
            {"name": "Customer concentration", "probability": 2, "impact": 3},
            {"name": "Key person risk", "probability": 1, "impact": 3},
            {"name": "Technology disruption", "probability": 2, "impact": 2},
            {"name": "Regulatory change", "probability": 2, "impact": 2},
            {"name": "Integration execution", "probability": 3, "impact": 2},
            {"name": "Interest rate risk", "probability": 1, "impact": 1},
            {"name": "FX exposure", "probability": 2, "impact": 1},
            {"name": "Supply chain disruption", "probability": 1, "impact": 2},
        ])

        # 3x3 grid colors (impact cols, probability rows — row 0 = high prob)
        # [probability=High][impact=Low/Med/High]
        GRID_COLORS = [
            [RGBColor(255, 235, 156), RGBColor(255, 199, 206), RGBColor(255, 99, 71)],    # High prob
            [RGBColor(198, 239, 206), RGBColor(255, 235, 156), RGBColor(255, 199, 206)],   # Med prob
            [RGBColor(198, 239, 206), RGBColor(198, 239, 206), RGBColor(255, 235, 156)],   # Low prob
        ]

        grid_left = 2.0
        grid_top = 1.2
        cell_w = 2.0
        cell_h = 1.6
        label_width = 1.2

        # Y-axis label (Probability)
        txBox_y = slide.shapes.add_textbox(
            Inches(grid_left - 1.3), Inches(grid_top + 1.0), Inches(1.0), Inches(2.0))
        p_y = txBox_y.text_frame.paragraphs[0]
        p_y.text = "Probability"
        p_y.font.name = FONT_NAME
        p_y.font.size = Pt(12)
        p_y.font.bold = True
        p_y.font.color.rgb = NAVY

        # X-axis label (Impact)
        txBox_x = slide.shapes.add_textbox(
            Inches(grid_left + 1.5), Inches(grid_top + 3 * cell_h + 0.15), Inches(2.0), Inches(0.3))
        p_x = txBox_x.text_frame.paragraphs[0]
        p_x.text = "Impact"
        p_x.font.name = FONT_NAME
        p_x.font.size = Pt(12)
        p_x.font.bold = True
        p_x.font.color.rgb = NAVY
        p_x.alignment = PP_ALIGN.CENTER

        # Row labels (High=0, Medium=1, Low=2)
        prob_labels = ["High", "Medium", "Low"]
        for r, label in enumerate(prob_labels):
            txBox_rl = slide.shapes.add_textbox(
                Inches(grid_left - label_width - 0.1), Inches(grid_top + r * cell_h + cell_h / 2 - 0.15),
                Inches(label_width), Inches(0.3))
            p_rl = txBox_rl.text_frame.paragraphs[0]
            p_rl.text = label
            p_rl.font.name = FONT_NAME
            p_rl.font.size = Pt(10)
            p_rl.font.bold = True
            p_rl.font.color.rgb = NAVY
            p_rl.alignment = PP_ALIGN.RIGHT

        # Column labels (Low=0, Medium=1, High=2)
        impact_labels = ["Low", "Medium", "High"]
        for c, label in enumerate(impact_labels):
            txBox_cl = slide.shapes.add_textbox(
                Inches(grid_left + c * cell_w), Inches(grid_top - 0.3),
                Inches(cell_w), Inches(0.3))
            p_cl = txBox_cl.text_frame.paragraphs[0]
            p_cl.text = label
            p_cl.font.name = FONT_NAME
            p_cl.font.size = Pt(10)
            p_cl.font.bold = True
            p_cl.font.color.rgb = NAVY
            p_cl.alignment = PP_ALIGN.CENTER

        # Draw grid cells
        for r in range(3):
            for c in range(3):
                box = slide.shapes.add_shape(
                    MSO_SHAPE.RECTANGLE,
                    Inches(grid_left + c * cell_w), Inches(grid_top + r * cell_h),
                    Inches(cell_w), Inches(cell_h))
                box.fill.solid()
                box.fill.fore_color.rgb = GRID_COLORS[r][c]
                box.line.color.rgb = WHITE
                box.line.width = Pt(2)

        # Place risk labels in cells
        # probability: 1=Low(row 2), 2=Med(row 1), 3=High(row 0)
        # impact: 1=Low(col 0), 2=Med(col 1), 3=High(col 2)
        cell_risks = {}  # (row, col) -> list of risk names
        for risk in risks:
            prob = risk.get('probability', 1)
            impact = risk.get('impact', 1)
            row = 3 - prob  # Convert: 3->0(High), 2->1(Med), 1->2(Low)
            col = impact - 1  # Convert: 1->0(Low), 2->1(Med), 3->2(High)
            row = max(0, min(2, row))
            col = max(0, min(2, col))
            key = (row, col)
            if key not in cell_risks:
                cell_risks[key] = []
            cell_risks[key].append(risk.get('name', ''))

        for (r, c), names in cell_risks.items():
            txBox_risk = slide.shapes.add_textbox(
                Inches(grid_left + c * cell_w + 0.1),
                Inches(grid_top + r * cell_h + 0.1),
                Inches(cell_w - 0.2), Inches(cell_h - 0.2))
            tf_risk = txBox_risk.text_frame
            tf_risk.word_wrap = True
            for j, name in enumerate(names):
                p_r = tf_risk.paragraphs[0] if j == 0 else tf_risk.add_paragraph()
                p_r.text = f"• {name}"
                p_r.font.name = FONT_NAME
                p_r.font.size = Pt(8)
                p_r.font.bold = True
                p_r.font.color.rgb = DARK_GRAY

        # Page number
        txBox_pg = slide.shapes.add_textbox(Inches(9.0), Inches(7.0), Inches(0.8), Inches(0.3))
        p_pg = txBox_pg.text_frame.paragraphs[0]
        p_pg.text = str(page)
        p_pg.font.name = FONT_NAME
        p_pg.font.size = Pt(10)
        p_pg.font.color.rgb = DARK_GRAY

        return self

    def add_tam_funnel_slide(self, data: Dict = None, include_cover: bool = True):
        """
        Add TAM/SAM/SOM Funnel slide — descending-width bars for market sizing.

        Args:
            data: Optional dict with keys: tam, sam, som (dicts with value, description, source),
                  company_share
            include_cover: Whether to add a section header
        """
        data = data or {}

        if include_cover:
            self._add_section_header("Market Sizing")

        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        page = self._next_page()

        # Title
        txBox = slide.shapes.add_textbox(Inches(0.3), Inches(0.2), Inches(9.5), Inches(0.3))
        p = txBox.text_frame.paragraphs[0]
        p.text = f"TAM / SAM / SOM Analysis — {self.company_name}"
        p.font.name = FONT_NAME
        p.font.size = Pt(16)
        p.font.bold = True
        p.font.color.rgb = NAVY

        # Funnel layers
        layers = [
            {
                "label": "TAM (Total Addressable Market)",
                "value": data.get('tam', {}).get('value', '$45B'),
                "description": data.get('tam', {}).get('description', 'Global market for [industry] solutions'),
                "source": data.get('tam', {}).get('source', 'Industry report, 2025'),
                "color": RGBColor(0, 54, 91),         # Navy
                "width": 8.0,
            },
            {
                "label": "SAM (Serviceable Addressable Market)",
                "value": data.get('sam', {}).get('value', '$12B'),
                "description": data.get('sam', {}).get('description', 'North American mid-market segment'),
                "source": data.get('sam', {}).get('source', 'Management estimate'),
                "color": RGBColor(50, 100, 150),       # Medium blue
                "width": 5.6,
            },
            {
                "label": "SOM (Serviceable Obtainable Market)",
                "value": data.get('som', {}).get('value', '$1.8B'),
                "description": data.get('som', {}).get('description', 'Realistic capture based on go-to-market'),
                "source": data.get('som', {}).get('source', 'Bottom-up analysis'),
                "color": RGBColor(100, 150, 200),      # Light blue
                "width": 3.2,
            },
        ]

        center_x = 5.0
        bar_height = 1.3
        gap = 0.2
        start_top = 1.0

        for i, layer in enumerate(layers):
            top = start_top + i * (bar_height + gap)
            width = layer['width']
            left = center_x - width / 2

            # Bar
            bar = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                Inches(left), Inches(top), Inches(width), Inches(bar_height))
            bar.fill.solid()
            bar.fill.fore_color.rgb = layer['color']
            bar.line.fill.background()

            # Label + value on the bar
            txBox_bar = slide.shapes.add_textbox(
                Inches(left + 0.2), Inches(top + 0.15), Inches(width - 0.4), Inches(0.4))
            p_bar = txBox_bar.text_frame.paragraphs[0]
            p_bar.text = f"{layer['label']}  —  {layer['value']}"
            p_bar.font.name = FONT_NAME
            p_bar.font.size = Pt(11)
            p_bar.font.bold = True
            p_bar.font.color.rgb = WHITE
            p_bar.alignment = PP_ALIGN.CENTER

            # Description on the bar
            txBox_desc = slide.shapes.add_textbox(
                Inches(left + 0.2), Inches(top + 0.55), Inches(width - 0.4), Inches(0.5))
            tf_desc = txBox_desc.text_frame
            tf_desc.word_wrap = True
            p_desc = tf_desc.paragraphs[0]
            p_desc.text = layer['description']
            p_desc.font.name = FONT_NAME
            p_desc.font.size = Pt(9)
            p_desc.font.color.rgb = RGBColor(200, 220, 240)
            p_desc.alignment = PP_ALIGN.CENTER

        # Company share callout
        share_top = start_top + 3 * (bar_height + gap) + 0.2
        share_box = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(2.5), Inches(share_top), Inches(5.0), Inches(0.7))
        share_box.fill.solid()
        share_box.fill.fore_color.rgb = RGBColor(248, 250, 252)
        share_box.line.color.rgb = NAVY
        share_box.line.width = Pt(1.5)

        txBox_share = slide.shapes.add_textbox(
            Inches(2.7), Inches(share_top + 0.1), Inches(4.6), Inches(0.5))
        tf_share = txBox_share.text_frame
        tf_share.word_wrap = True
        p_share = tf_share.paragraphs[0]
        p_share.text = f"{self.company_name} Current Revenue: {data.get('company_revenue', '$320M')}  |  SOM Penetration: {data.get('som_penetration', '17.8%')}"
        p_share.font.name = FONT_NAME
        p_share.font.size = Pt(10)
        p_share.font.bold = True
        p_share.font.color.rgb = NAVY
        p_share.alignment = PP_ALIGN.CENTER

        # Key assumptions
        assumptions = data.get('assumptions', [
            "TAM growing at 7-9% CAGR driven by digital transformation and regulatory tailwinds",
            "SAM defined by geography (North America) and segment (mid-market, $50-500M revenue)",
            "SOM assumes 15-20% market share achievable within 5-year hold period",
        ])

        txBox_assume = slide.shapes.add_textbox(Inches(0.5), Inches(share_top + 1.0), Inches(9.0), Inches(1.5))
        tf_assume = txBox_assume.text_frame
        tf_assume.word_wrap = True
        p_ah = tf_assume.paragraphs[0]
        p_ah.text = "Key Assumptions"
        p_ah.font.name = FONT_NAME
        p_ah.font.size = Pt(10)
        p_ah.font.bold = True
        p_ah.font.color.rgb = NAVY

        for assumption in assumptions:
            p_a = tf_assume.add_paragraph()
            p_a.text = f"• {assumption}"
            p_a.font.name = FONT_NAME
            p_a.font.size = Pt(9)
            p_a.font.color.rgb = DARK_GRAY

        # Page number
        txBox_pg = slide.shapes.add_textbox(Inches(9.0), Inches(7.0), Inches(0.8), Inches(0.3))
        p_pg = txBox_pg.text_frame.paragraphs[0]
        p_pg.text = str(page)
        p_pg.font.name = FONT_NAME
        p_pg.font.size = Pt(10)
        p_pg.font.color.rgb = DARK_GRAY

        return self

    def add_pestel_slide(self, data: Dict = None, include_cover: bool = True):
        """
        Add PESTEL Analysis slide — 3x2 grid of macro factors.

        Args:
            data: Optional dict with keys for each factor: political, economic, social,
                  technological, environmental, legal (each a list of bullet strings)
            include_cover: Whether to add a section header
        """
        data = data or {}

        if include_cover:
            self._add_section_header("Macro Environment")

        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        page = self._next_page()

        # Title
        txBox = slide.shapes.add_textbox(Inches(0.3), Inches(0.2), Inches(9.5), Inches(0.3))
        p = txBox.text_frame.paragraphs[0]
        p.text = f"PESTEL Analysis — {self.company_name}"
        p.font.name = FONT_NAME
        p.font.size = Pt(16)
        p.font.bold = True
        p.font.color.rgb = NAVY

        factors = [
            ("Political", "P", data.get('political', [
                "Regulatory stability in core markets",
                "Trade policy impact on supply chain",
                "Government subsidy programs",
            ]), RGBColor(0, 54, 91)),
            ("Economic", "E", data.get('economic', [
                "GDP growth outlook: 2-3%",
                "Interest rate environment",
                "Inflation impact on margins",
            ]), RGBColor(34, 100, 34)),
            ("Social", "S", data.get('social', [
                "Demographic shifts (aging population)",
                "Consumer behavior trends",
                "Workforce availability",
            ]), RGBColor(150, 80, 0)),
            ("Technological", "T", data.get('technological', [
                "AI/automation disruption potential",
                "Digital transformation spend",
                "Cybersecurity requirements",
            ]), RGBColor(120, 50, 150)),
            ("Environmental", "E", data.get('environmental', [
                "Carbon regulation compliance",
                "ESG reporting requirements",
                "Supply chain sustainability",
            ]), RGBColor(0, 120, 120)),
            ("Legal", "L", data.get('legal', [
                "Data privacy (GDPR, CCPA)",
                "IP protection regime",
                "Employment law changes",
            ]), RGBColor(178, 34, 34)),
        ]

        # 3x2 grid layout
        col_count = 3
        row_count = 2
        box_width = 2.9
        box_height = 2.7
        h_gap = 0.15
        v_gap = 0.15
        grid_left = 0.4
        grid_top = 0.8

        for idx, (name, letter, bullets, color) in enumerate(factors):
            row = idx // col_count
            col = idx % col_count
            left = grid_left + col * (box_width + h_gap)
            top = grid_top + row * (box_height + v_gap)

            # Box
            box = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                Inches(left), Inches(top), Inches(box_width), Inches(box_height))
            box.fill.solid()
            box.fill.fore_color.rgb = RGBColor(248, 250, 252)
            box.line.color.rgb = color
            box.line.width = Pt(1.5)

            # Factor header
            txBox_h = slide.shapes.add_textbox(
                Inches(left + 0.15), Inches(top + 0.1),
                Inches(box_width - 0.3), Inches(0.35))
            p_h = txBox_h.text_frame.paragraphs[0]
            p_h.text = name
            p_h.font.name = FONT_NAME
            p_h.font.size = Pt(11)
            p_h.font.bold = True
            p_h.font.color.rgb = color

            # Bullets
            txBox_b = slide.shapes.add_textbox(
                Inches(left + 0.15), Inches(top + 0.5),
                Inches(box_width - 0.3), Inches(box_height - 0.6))
            tf_b = txBox_b.text_frame
            tf_b.word_wrap = True
            for j, bullet in enumerate(bullets):
                p_b = tf_b.paragraphs[0] if j == 0 else tf_b.add_paragraph()
                p_b.text = f"• {bullet}"
                p_b.font.name = FONT_NAME
                p_b.font.size = Pt(8)
                p_b.font.color.rgb = DARK_GRAY
                p_b.space_after = Pt(3)

        # Page number
        txBox_pg = slide.shapes.add_textbox(Inches(9.0), Inches(7.0), Inches(0.8), Inches(0.3))
        p_pg = txBox_pg.text_frame.paragraphs[0]
        p_pg.text = str(page)
        p_pg.font.name = FONT_NAME
        p_pg.font.size = Pt(10)
        p_pg.font.color.rgb = DARK_GRAY

        return self

    def add_porters_radar_slide(self, data: Dict = None, include_cover: bool = True):
        """
        Add Porter's Five Forces Radar slide — quantified pentagon chart with
        force ratings on a 1-5 scale.

        Args:
            data: Optional dict with keys: forces (dict with supplier_power, buyer_power,
                  new_entrants, substitutes, rivalry — each 1-5), descriptions (dict)
            include_cover: Whether to add a section header
        """
        data = data or {}

        if include_cover:
            self._add_section_header("Competitive Forces")

        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        page = self._next_page()

        # Title
        txBox = slide.shapes.add_textbox(Inches(0.3), Inches(0.2), Inches(9.5), Inches(0.3))
        p = txBox.text_frame.paragraphs[0]
        p.text = f"Porter's Five Forces — {self.company_name}"
        p.font.name = FONT_NAME
        p.font.size = Pt(16)
        p.font.bold = True
        p.font.color.rgb = NAVY

        force_defaults = {
            "rivalry": 4, "new_entrants": 2, "substitutes": 3,
            "buyer_power": 3, "supplier_power": 2,
        }
        forces = data.get('forces', force_defaults)

        desc_defaults = {
            "rivalry": "Intense: fragmented market with 5+ major players",
            "new_entrants": "Low: high barriers (capital, regulation, relationships)",
            "substitutes": "Moderate: alternative solutions exist but inferior",
            "buyer_power": "Moderate: concentrated buyer base but high switching costs",
            "supplier_power": "Low: multiple suppliers, commodity inputs",
        }
        descriptions = data.get('descriptions', desc_defaults)

        # Pentagon radar chart parameters
        force_names = ["Competitive\nRivalry", "Threat of\nNew Entrants", "Supplier\nPower",
                       "Threat of\nSubstitutes", "Buyer\nPower"]
        force_keys = ["rivalry", "new_entrants", "supplier_power", "substitutes", "buyer_power"]
        ratings = [forces.get(k, 3) for k in force_keys]

        cx, cy = 3.3, 3.6  # Center of pentagon (inches)
        max_r = 1.8         # Max radius (inches)

        # Calculate pentagon vertices for each ring (1-5)
        angles = [math.pi / 2 - i * 2 * math.pi / 5 for i in range(5)]

        # Draw concentric pentagons (gridlines)
        for ring in range(1, 6):
            scale = ring / 5.0
            points = []
            for angle in angles:
                x = cx + max_r * scale * math.cos(angle)
                y = cy - max_r * scale * math.sin(angle)
                points.append((x, y))

            # Draw lines between consecutive vertices
            for j in range(5):
                x1, y1 = points[j]
                x2, y2 = points[(j + 1) % 5]
                line = slide.shapes.add_connector(
                    1,  # MSO_CONNECTOR.STRAIGHT
                    Inches(x1), Inches(y1), Inches(x2), Inches(y2))
                line.line.color.rgb = RGBColor(220, 220, 220)
                line.line.width = Pt(0.5)

        # Draw axis lines from center to each vertex
        for angle in angles:
            x_end = cx + max_r * math.cos(angle)
            y_end = cy - max_r * math.sin(angle)
            line = slide.shapes.add_connector(
                1, Inches(cx), Inches(cy), Inches(x_end), Inches(y_end))
            line.line.color.rgb = RGBColor(200, 200, 200)
            line.line.width = Pt(0.5)

        # Draw the rating polygon (filled)
        # We'll use individual lines for the data shape since python-pptx
        # freeform is complex — draw thick navy lines connecting data points
        data_points = []
        for i, angle in enumerate(angles):
            scale = ratings[i] / 5.0
            x = cx + max_r * scale * math.cos(angle)
            y = cy - max_r * scale * math.sin(angle)
            data_points.append((x, y))

        for j in range(5):
            x1, y1 = data_points[j]
            x2, y2 = data_points[(j + 1) % 5]
            line = slide.shapes.add_connector(
                1, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
            line.line.color.rgb = NAVY
            line.line.width = Pt(2.5)

        # Draw data point markers
        marker_size = 0.15
        for x, y in data_points:
            marker = slide.shapes.add_shape(
                MSO_SHAPE.OVAL,
                Inches(x - marker_size / 2), Inches(y - marker_size / 2),
                Inches(marker_size), Inches(marker_size))
            marker.fill.solid()
            marker.fill.fore_color.rgb = NAVY
            marker.line.fill.background()

        # Force labels at each vertex (outside the pentagon)
        label_offsets = [
            (0, -0.55),     # Top (rivalry) - above
            (0.6, -0.3),   # Upper-right (new entrants)
            (0.5, 0.25),   # Lower-right (supplier power)
            (-0.5, 0.25),  # Lower-left (substitutes)
            (-0.6, -0.3),  # Upper-left (buyer power)
        ]

        for i, angle in enumerate(angles):
            x_label = cx + (max_r + 0.35) * math.cos(angle) + label_offsets[i][0]
            y_label = cy - (max_r + 0.35) * math.sin(angle) + label_offsets[i][1]

            txBox_label = slide.shapes.add_textbox(
                Inches(x_label - 0.6), Inches(y_label),
                Inches(1.5), Inches(0.5))
            tf_label = txBox_label.text_frame
            tf_label.word_wrap = True
            p_label = tf_label.paragraphs[0]
            p_label.text = f"{force_names[i]}\n({ratings[i]}/5)"
            p_label.font.name = FONT_NAME
            p_label.font.size = Pt(9)
            p_label.font.bold = True
            p_label.font.color.rgb = NAVY
            p_label.alignment = PP_ALIGN.CENTER

        # Descriptions panel (right side)
        txBox_desc = slide.shapes.add_textbox(Inches(6.0), Inches(1.0), Inches(3.8), Inches(5.5))
        tf_desc = txBox_desc.text_frame
        tf_desc.word_wrap = True

        p_dh = tf_desc.paragraphs[0]
        p_dh.text = "Force Assessment"
        p_dh.font.name = FONT_NAME
        p_dh.font.size = Pt(11)
        p_dh.font.bold = True
        p_dh.font.color.rgb = NAVY

        desc_labels = ["Rivalry", "New Entrants", "Supplier Power", "Substitutes", "Buyer Power"]
        for i, key in enumerate(force_keys):
            p_name = tf_desc.add_paragraph()
            p_name.text = f"{desc_labels[i]} ({ratings[i]}/5)"
            p_name.font.name = FONT_NAME
            p_name.font.size = Pt(9)
            p_name.font.bold = True
            p_name.font.color.rgb = NAVY
            p_name.space_before = Pt(6)

            p_detail = tf_desc.add_paragraph()
            p_detail.text = descriptions.get(key, 'N/A')
            p_detail.font.name = FONT_NAME
            p_detail.font.size = Pt(8)
            p_detail.font.color.rgb = DARK_GRAY

        # Page number
        txBox_pg = slide.shapes.add_textbox(Inches(9.0), Inches(7.0), Inches(0.8), Inches(0.3))
        p_pg = txBox_pg.text_frame.paragraphs[0]
        p_pg.text = str(page)
        p_pg.font.name = FONT_NAME
        p_pg.font.size = Pt(10)
        p_pg.font.color.rgb = DARK_GRAY

        return self

    # ============================================================
    # CFA EXCEL+SLIDE PAIRS
    # ============================================================

    def add_tornado_sensitivity_slide(self, data: Dict = None, include_cover: bool = True):
        """
        Add Tornado Sensitivity Chart slide — horizontal bars showing IRR range
        for each variable, sorted by impact.

        Args:
            data: Optional dict with keys: variables (list of dicts with name,
                  low_irr, base_irr, high_irr), base_irr
            include_cover: Whether to add a section header
        """
        data = data or {}

        if include_cover:
            self._add_section_header("Sensitivity Analysis")

        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        page = self._next_page()

        # Title
        txBox = slide.shapes.add_textbox(Inches(0.3), Inches(0.2), Inches(9.5), Inches(0.3))
        p = txBox.text_frame.paragraphs[0]
        p.text = f"Tornado Sensitivity — {self.company_name}"
        p.font.name = FONT_NAME
        p.font.size = Pt(16)
        p.font.bold = True
        p.font.color.rgb = NAVY

        txBox_sub = slide.shapes.add_textbox(Inches(0.3), Inches(0.55), Inches(9.5), Inches(0.3))
        p_sub = txBox_sub.text_frame.paragraphs[0]
        p_sub.text = "Single-variable sensitivity on Sponsor IRR (base case shown as center line)"
        p_sub.font.name = FONT_NAME
        p_sub.font.size = Pt(10)
        p_sub.font.color.rgb = DARK_GRAY

        base_irr = data.get('base_irr', 22.0)
        variables = data.get('variables', [
            {"name": "Exit Multiple", "low_irr": 16.0, "high_irr": 29.0},
            {"name": "Entry Multiple", "low_irr": 18.0, "high_irr": 27.5},
            {"name": "Revenue Growth", "low_irr": 18.5, "high_irr": 26.0},
            {"name": "EBITDA Margin", "low_irr": 19.0, "high_irr": 25.5},
            {"name": "Leverage (Turns)", "low_irr": 20.0, "high_irr": 24.5},
            {"name": "Interest Rate", "low_irr": 20.5, "high_irr": 23.5},
            {"name": "Capex % Revenue", "low_irr": 21.0, "high_irr": 23.0},
        ])

        # Sort by impact range (widest first)
        variables.sort(key=lambda v: (v.get('high_irr', base_irr) - v.get('low_irr', base_irr)), reverse=True)

        # Chart area
        chart_left = 2.2
        chart_right = 9.2
        chart_width = chart_right - chart_left
        chart_top = 1.1
        bar_height = 0.55
        bar_gap = 0.12

        # IRR range for scaling
        all_irrs = [v.get('low_irr', base_irr) for v in variables] + [v.get('high_irr', base_irr) for v in variables]
        min_irr = min(all_irrs) - 2
        max_irr = max(all_irrs) + 2
        irr_range = max_irr - min_irr

        def irr_to_x(irr):
            return chart_left + (irr - min_irr) / irr_range * chart_width

        # Base case vertical line
        base_x = irr_to_x(base_irr)
        base_line = slide.shapes.add_connector(
            1, Inches(base_x), Inches(chart_top - 0.1),
            Inches(base_x), Inches(chart_top + len(variables) * (bar_height + bar_gap) + 0.1))
        base_line.line.color.rgb = DARK_GRAY
        base_line.line.width = Pt(1.5)

        # Base IRR label
        txBox_base = slide.shapes.add_textbox(
            Inches(base_x - 0.4), Inches(chart_top - 0.35), Inches(0.8), Inches(0.25))
        p_base = txBox_base.text_frame.paragraphs[0]
        p_base.text = f"Base: {base_irr:.1f}%"
        p_base.font.name = FONT_NAME
        p_base.font.size = Pt(8)
        p_base.font.bold = True
        p_base.font.color.rgb = DARK_GRAY
        p_base.alignment = PP_ALIGN.CENTER

        for i, var in enumerate(variables):
            top = chart_top + i * (bar_height + bar_gap)
            low_irr = var.get('low_irr', base_irr)
            high_irr = var.get('high_irr', base_irr)

            # Variable label (left side)
            txBox_name = slide.shapes.add_textbox(
                Inches(0.2), Inches(top + 0.05), Inches(1.9), Inches(bar_height - 0.1))
            p_name = txBox_name.text_frame.paragraphs[0]
            p_name.text = var.get('name', '')
            p_name.font.name = FONT_NAME
            p_name.font.size = Pt(9)
            p_name.font.bold = True
            p_name.font.color.rgb = DARK_GRAY
            p_name.alignment = PP_ALIGN.RIGHT

            # Low bar (left of base) — red
            if low_irr < base_irr:
                x_start = irr_to_x(low_irr)
                x_end = irr_to_x(base_irr)
                bar_w = x_end - x_start
                bar = slide.shapes.add_shape(
                    MSO_SHAPE.RECTANGLE,
                    Inches(x_start), Inches(top), Inches(bar_w), Inches(bar_height))
                bar.fill.solid()
                bar.fill.fore_color.rgb = RGBColor(255, 120, 120)
                bar.line.fill.background()

                # Low value label
                txBox_low = slide.shapes.add_textbox(
                    Inches(x_start - 0.4), Inches(top + 0.05),
                    Inches(0.4), Inches(bar_height - 0.1))
                p_low = txBox_low.text_frame.paragraphs[0]
                p_low.text = f"{low_irr:.1f}%"
                p_low.font.name = FONT_NAME
                p_low.font.size = Pt(7)
                p_low.font.color.rgb = ACCENT_RED
                p_low.alignment = PP_ALIGN.RIGHT

            # High bar (right of base) — green
            if high_irr > base_irr:
                x_start = irr_to_x(base_irr)
                x_end = irr_to_x(high_irr)
                bar_w = x_end - x_start
                bar = slide.shapes.add_shape(
                    MSO_SHAPE.RECTANGLE,
                    Inches(x_start), Inches(top), Inches(bar_w), Inches(bar_height))
                bar.fill.solid()
                bar.fill.fore_color.rgb = RGBColor(120, 200, 120)
                bar.line.fill.background()

                # High value label
                txBox_high = slide.shapes.add_textbox(
                    Inches(x_end), Inches(top + 0.05),
                    Inches(0.4), Inches(bar_height - 0.1))
                p_high = txBox_high.text_frame.paragraphs[0]
                p_high.text = f"{high_irr:.1f}%"
                p_high.font.name = FONT_NAME
                p_high.font.size = Pt(7)
                p_high.font.color.rgb = ACCENT_GREEN
                p_high.alignment = PP_ALIGN.LEFT

        # Page number
        txBox_pg = slide.shapes.add_textbox(Inches(9.0), Inches(7.0), Inches(0.8), Inches(0.3))
        p_pg = txBox_pg.text_frame.paragraphs[0]
        p_pg.text = str(page)
        p_pg.font.name = FONT_NAME
        p_pg.font.size = Pt(10)
        p_pg.font.color.rgb = DARK_GRAY

        return self

    def add_football_field_slide(self, data: Dict = None, include_cover: bool = True):
        """
        Add Football Field Valuation slide — horizontal range bars for each methodology.

        Args:
            data: Optional dict with keys: methods (list of dicts with name, low, mid, high),
                  current_price
            include_cover: Whether to add a section header
        """
        data = data or {}

        if include_cover:
            self._add_section_header("Valuation Summary")

        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        page = self._next_page()

        # Title
        txBox = slide.shapes.add_textbox(Inches(0.3), Inches(0.2), Inches(9.5), Inches(0.3))
        p = txBox.text_frame.paragraphs[0]
        p.text = f"Football Field Valuation — {self.company_name}"
        p.font.name = FONT_NAME
        p.font.size = Pt(16)
        p.font.bold = True
        p.font.color.rgb = NAVY

        methods = data.get('methods', [
            {"name": "DCF (WACC 9-11%)", "low": 38.0, "mid": 48.5, "high": 58.0},
            {"name": "Trading Comps (EV/EBITDA)", "low": 36.0, "mid": 46.0, "high": 52.0},
            {"name": "Transaction Comps", "low": 40.0, "mid": 50.0, "high": 60.0},
            {"name": "Dividend Discount Model", "low": 34.0, "mid": 44.0, "high": 50.0},
            {"name": "52-Week Range", "low": 32.0, "mid": 41.0, "high": 49.0},
        ])

        current_price = data.get('current_price', 41.25)

        # Chart area
        chart_left = 3.0
        chart_right = 9.2
        chart_width = chart_right - chart_left
        chart_top = 1.0
        bar_height = 0.7
        bar_gap = 0.3

        # Value range for scaling
        all_vals = [m.get('low', 0) for m in methods] + [m.get('high', 0) for m in methods]
        min_val = min(all_vals) - 5
        max_val = max(all_vals) + 5
        val_range = max_val - min_val

        def val_to_x(val):
            return chart_left + (val - min_val) / val_range * chart_width

        # Color palette for bars
        bar_colors = [
            NAVY,
            RGBColor(50, 100, 150),
            RGBColor(100, 150, 200),
            RGBColor(34, 139, 34),
            RGBColor(150, 100, 50),
        ]

        for i, method in enumerate(methods):
            top = chart_top + i * (bar_height + bar_gap)
            low = method.get('low', 0)
            mid = method.get('mid', 0)
            high = method.get('high', 0)
            color = bar_colors[i % len(bar_colors)]

            # Method label
            txBox_name = slide.shapes.add_textbox(
                Inches(0.2), Inches(top + 0.1), Inches(2.7), Inches(bar_height - 0.2))
            p_name = txBox_name.text_frame.paragraphs[0]
            p_name.text = method.get('name', '')
            p_name.font.name = FONT_NAME
            p_name.font.size = Pt(9)
            p_name.font.bold = True
            p_name.font.color.rgb = DARK_GRAY
            p_name.alignment = PP_ALIGN.RIGHT

            # Range bar (low to high)
            x_low = val_to_x(low)
            x_high = val_to_x(high)
            bar_w = x_high - x_low

            bar = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                Inches(x_low), Inches(top), Inches(bar_w), Inches(bar_height))
            bar.fill.solid()
            bar.fill.fore_color.rgb = color
            bar.line.fill.background()

            # Mid-point marker (white diamond)
            x_mid = val_to_x(mid)
            marker_size = 0.18
            marker = slide.shapes.add_shape(
                MSO_SHAPE.DIAMOND,
                Inches(x_mid - marker_size / 2), Inches(top + bar_height / 2 - marker_size / 2),
                Inches(marker_size), Inches(marker_size))
            marker.fill.solid()
            marker.fill.fore_color.rgb = WHITE
            marker.line.color.rgb = color
            marker.line.width = Pt(1)

            # Value labels on bar
            txBox_vals = slide.shapes.add_textbox(
                Inches(x_low + 0.05), Inches(top + 0.05),
                Inches(bar_w - 0.1), Inches(bar_height - 0.1))
            p_vals = txBox_vals.text_frame.paragraphs[0]
            p_vals.text = f"${low:.0f}    —    ${mid:.0f}    —    ${high:.0f}"
            p_vals.font.name = FONT_NAME
            p_vals.font.size = Pt(8)
            p_vals.font.bold = True
            p_vals.font.color.rgb = WHITE
            p_vals.alignment = PP_ALIGN.CENTER

        # Current price vertical line
        if current_price:
            x_current = val_to_x(current_price)
            total_height = len(methods) * (bar_height + bar_gap)
            current_line = slide.shapes.add_connector(
                1, Inches(x_current), Inches(chart_top - 0.2),
                Inches(x_current), Inches(chart_top + total_height))
            current_line.line.color.rgb = ACCENT_RED
            current_line.line.width = Pt(2)

            txBox_current = slide.shapes.add_textbox(
                Inches(x_current - 0.5), Inches(chart_top - 0.4), Inches(1.0), Inches(0.2))
            p_current = txBox_current.text_frame.paragraphs[0]
            p_current.text = f"Current: ${current_price:.2f}"
            p_current.font.name = FONT_NAME
            p_current.font.size = Pt(8)
            p_current.font.bold = True
            p_current.font.color.rgb = ACCENT_RED
            p_current.alignment = PP_ALIGN.CENTER

        # Page number
        txBox_pg = slide.shapes.add_textbox(Inches(9.0), Inches(7.0), Inches(0.8), Inches(0.3))
        p_pg = txBox_pg.text_frame.paragraphs[0]
        p_pg.text = str(page)
        p_pg.font.name = FONT_NAME
        p_pg.font.size = Pt(10)
        p_pg.font.color.rgb = DARK_GRAY

        return self

    def add_sotp_waterfall_slide(self, data: Dict = None, include_cover: bool = True):
        """
        Add SOTP Waterfall slide — stacked segments with EV bridge to equity.

        Args:
            data: Optional dict with keys: segments (list of dicts with name, ev, color),
                  adjustments (list of dicts with name, value — negative for deductions),
                  equity_value, shares_outstanding
            include_cover: Whether to add a section header
        """
        data = data or {}

        if include_cover:
            self._add_section_header("Sum-of-the-Parts")

        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        page = self._next_page()

        # Title
        txBox = slide.shapes.add_textbox(Inches(0.3), Inches(0.2), Inches(9.5), Inches(0.3))
        p = txBox.text_frame.paragraphs[0]
        p.text = f"SOTP Waterfall — {self.company_name}"
        p.font.name = FONT_NAME
        p.font.size = Pt(16)
        p.font.bold = True
        p.font.color.rgb = NAVY

        segments = data.get('segments', [
            {"name": "Core Business", "ev": 850},
            {"name": "Growth Division", "ev": 320},
            {"name": "International", "ev": 180},
            {"name": "Other / Corporate", "ev": 50},
        ])

        adjustments = data.get('adjustments', [
            {"name": "Less: Net Debt", "value": -420},
            {"name": "Less: Minority Interest", "value": -30},
            {"name": "Plus: Cash & Equivalents", "value": 85},
        ])

        total_ev = sum(s.get('ev', 0) for s in segments)
        equity_value = data.get('equity_value', total_ev + sum(a.get('value', 0) for a in adjustments))
        shares = data.get('shares_outstanding', 20.0)

        # Waterfall chart
        chart_left = 0.8
        chart_bottom = 5.5
        chart_height = 3.5  # max bar height in inches
        bar_width = 0.9
        gap = 0.15

        # All items: segments + total EV + adjustments + equity value
        all_items = []
        for s in segments:
            all_items.append({"name": s['name'], "value": s['ev'], "type": "segment"})
        all_items.append({"name": "Total EV", "value": total_ev, "type": "total"})
        for a in adjustments:
            all_items.append({"name": a['name'], "value": a['value'], "type": "adjustment"})
        all_items.append({"name": "Equity Value", "value": equity_value, "type": "equity"})

        max_val = max(total_ev, equity_value) * 1.1

        def val_to_height(val):
            return abs(val) / max_val * chart_height

        # Segment colors
        seg_colors = [
            NAVY,
            RGBColor(50, 100, 150),
            RGBColor(100, 150, 200),
            RGBColor(150, 180, 210),
        ]

        # Draw waterfall bars
        running_top = chart_bottom  # Where the next segment stacks

        for i, item in enumerate(all_items):
            left = chart_left + i * (bar_width + gap)
            val = item['value']
            h = val_to_height(val)

            if item['type'] == 'segment':
                # Stacking segments from bottom up
                bar_top = running_top - h
                color = seg_colors[i % len(seg_colors)]
                running_top = bar_top  # Next segment starts here

                bar = slide.shapes.add_shape(
                    MSO_SHAPE.RECTANGLE,
                    Inches(left), Inches(bar_top), Inches(bar_width), Inches(h))
                bar.fill.solid()
                bar.fill.fore_color.rgb = color
                bar.line.fill.background()

            elif item['type'] == 'total':
                # Total EV — full bar from bottom
                h_total = val_to_height(total_ev)
                bar_top = chart_bottom - h_total

                bar = slide.shapes.add_shape(
                    MSO_SHAPE.RECTANGLE,
                    Inches(left), Inches(bar_top), Inches(bar_width), Inches(h_total))
                bar.fill.solid()
                bar.fill.fore_color.rgb = RGBColor(0, 80, 130)
                bar.line.fill.background()

                running_top = bar_top  # Reset for adjustments

            elif item['type'] == 'adjustment':
                if val < 0:
                    # Deduction: bar drops down from running_top
                    bar_top = running_top
                    bar = slide.shapes.add_shape(
                        MSO_SHAPE.RECTANGLE,
                        Inches(left), Inches(bar_top), Inches(bar_width), Inches(h))
                    bar.fill.solid()
                    bar.fill.fore_color.rgb = ACCENT_RED
                    bar.line.fill.background()
                    running_top = bar_top + h
                else:
                    # Addition: bar goes up from running_top
                    bar_top = running_top - h
                    bar = slide.shapes.add_shape(
                        MSO_SHAPE.RECTANGLE,
                        Inches(left), Inches(bar_top), Inches(bar_width), Inches(h))
                    bar.fill.solid()
                    bar.fill.fore_color.rgb = ACCENT_GREEN
                    bar.line.fill.background()
                    running_top = bar_top

            elif item['type'] == 'equity':
                # Equity value — full bar from bottom
                h_eq = val_to_height(equity_value)
                bar_top = chart_bottom - h_eq

                bar = slide.shapes.add_shape(
                    MSO_SHAPE.RECTANGLE,
                    Inches(left), Inches(bar_top), Inches(bar_width), Inches(h_eq))
                bar.fill.solid()
                bar.fill.fore_color.rgb = ACCENT_GREEN
                bar.line.fill.background()

            # Value label above/below bar
            label_val = f"${abs(val):,.0f}M" if abs(val) >= 1 else f"${abs(val) * 1000:,.0f}K"
            if item['type'] == 'total':
                label_top = chart_bottom - val_to_height(total_ev) - 0.25
            elif item['type'] == 'equity':
                label_top = chart_bottom - val_to_height(equity_value) - 0.25
            elif item['type'] == 'adjustment' and val < 0:
                label_top = running_top + 0.02
            else:
                label_top = (bar_top if 'bar_top' in dir() else running_top) - 0.25

            txBox_val = slide.shapes.add_textbox(
                Inches(left - 0.1), Inches(label_top), Inches(bar_width + 0.2), Inches(0.2))
            p_val = txBox_val.text_frame.paragraphs[0]
            p_val.text = label_val
            p_val.font.name = FONT_NAME
            p_val.font.size = Pt(8)
            p_val.font.bold = True
            p_val.font.color.rgb = DARK_GRAY
            p_val.alignment = PP_ALIGN.CENTER

            # Item name label below chart
            txBox_item = slide.shapes.add_textbox(
                Inches(left - 0.15), Inches(chart_bottom + 0.1),
                Inches(bar_width + 0.3), Inches(0.5))
            tf_item = txBox_item.text_frame
            tf_item.word_wrap = True
            p_item = tf_item.paragraphs[0]
            p_item.text = item['name']
            p_item.font.name = FONT_NAME
            p_item.font.size = Pt(7)
            p_item.font.bold = True
            p_item.font.color.rgb = DARK_GRAY
            p_item.alignment = PP_ALIGN.CENTER

        # Per-share summary box
        summary_box = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(0.5), Inches(6.3), Inches(9.0), Inches(0.5))
        summary_box.fill.solid()
        summary_box.fill.fore_color.rgb = RGBColor(248, 250, 252)
        summary_box.line.color.rgb = NAVY
        summary_box.line.width = Pt(1.5)

        per_share = equity_value / shares if shares else 0
        txBox_summary = slide.shapes.add_textbox(Inches(0.7), Inches(6.35), Inches(8.6), Inches(0.4))
        p_summary = txBox_summary.text_frame.paragraphs[0]
        p_summary.text = (f"Total EV: ${total_ev:,.0f}M  |  Equity Value: ${equity_value:,.0f}M  |  "
                          f"Shares: {shares:.1f}M  |  Per Share: ${per_share:,.2f}")
        p_summary.font.name = FONT_NAME
        p_summary.font.size = Pt(10)
        p_summary.font.bold = True
        p_summary.font.color.rgb = NAVY
        p_summary.alignment = PP_ALIGN.CENTER

        # Page number
        txBox_pg = slide.shapes.add_textbox(Inches(9.0), Inches(7.0), Inches(0.8), Inches(0.3))
        p_pg = txBox_pg.text_frame.paragraphs[0]
        p_pg.text = str(page)
        p_pg.font.name = FONT_NAME
        p_pg.font.size = Pt(10)
        p_pg.font.color.rgb = DARK_GRAY

        return self

    # ============================================================
    # OUTPUT
    # ============================================================

    def save(self, filepath: str) -> str:
        """Save the presentation to file."""
        self.prs.save(filepath)
        print(f"Presentation saved: {filepath}")
        print(f"Total slides: {len(self.prs.slides)}")
        return filepath

    def get_presentation(self) -> Presentation:
        """Return the raw presentation object for further manipulation."""
        return self.prs


# ============================================================
# QUICK ACCESS FUNCTIONS
# ============================================================

def generate_industry_analysis(company_name: str, output_path: str, data: Dict = None) -> str:
    """Quick function to generate just industry analysis slides."""
    gen = SlideGenerator(company_name)
    gen._add_cover_slide(f"{company_name} - Industry Analysis", "Market & Competitive Dynamics")
    gen.add_industry_analysis(data, include_cover=False)
    return gen.save(output_path)


def generate_competitive_analysis(company_name: str, output_path: str, data: Dict = None) -> str:
    """Quick function to generate just competitive analysis slides."""
    gen = SlideGenerator(company_name)
    gen._add_cover_slide(f"{company_name} - Competitive Analysis", "Market Positioning Assessment")
    gen.add_competitive_analysis(data, include_cover=False)
    return gen.save(output_path)


def generate_financial_analysis(company_name: str, output_path: str, data: Dict = None) -> str:
    """Quick function to generate just financial analysis slides."""
    gen = SlideGenerator(company_name)
    gen._add_cover_slide(f"{company_name} - Financial Analysis", "Historical & Projected Performance")
    gen.add_financial_analysis(data, include_cover=False)
    return gen.save(output_path)


def generate_debt_analysis(company_name: str, output_path: str, data: Dict = None) -> str:
    """Quick function to generate just debt/capital structure slides."""
    gen = SlideGenerator(company_name)
    gen._add_cover_slide(f"{company_name} - Capital Structure Analysis", "Debt Capacity & Coverage")
    gen.add_debt_analysis(data, include_cover=False)
    return gen.save(output_path)


def generate_lbo_analysis(company_name: str, output_path: str, data: Dict = None) -> str:
    """Quick function to generate just LBO/transaction slides."""
    gen = SlideGenerator(company_name)
    gen._add_cover_slide(f"{company_name} - Transaction Analysis", "LBO Returns & Sensitivity")
    gen.add_lbo_analysis(data, include_cover=False)
    return gen.save(output_path)


def generate_management_analysis(company_name: str, output_path: str, data: Dict = None) -> str:
    """Quick function to generate just management analysis slides."""
    gen = SlideGenerator(company_name)
    gen._add_cover_slide(f"{company_name} - Management Analysis", "Leadership & Team Assessment")
    gen.add_management_analysis(data, include_cover=False)
    return gen.save(output_path)


def generate_valuation_analysis(company_name: str, output_path: str, data: Dict = None) -> str:
    """Quick function to generate just valuation slides."""
    gen = SlideGenerator(company_name)
    gen._add_cover_slide(f"{company_name} - Valuation Analysis", "Comparable Companies & Transactions")
    gen.add_valuation(data, include_cover=False)
    return gen.save(output_path)


def generate_investment_thesis(company_name: str, output_path: str, data: Dict = None) -> str:
    """Quick function to generate just thesis and recommendation slides."""
    gen = SlideGenerator(company_name)
    gen._add_cover_slide(f"{company_name} - Investment Recommendation", "Thesis, Risks & Decision")
    gen.add_investment_thesis(data, include_cover=False)
    return gen.save(output_path)


def generate_full_deck(company_name: str, output_path: str, data: Dict = None) -> str:
    """Generate a complete investment memo deck."""
    gen = SlideGenerator(company_name)
    gen.add_full_deck(data)
    return gen.save(output_path)


# ============================================================
# INSTITUTIONAL MODULE QUICK ACCESS FUNCTIONS (7+ Day Case)
# ============================================================

def generate_scenario_analysis(company_name: str, output_path: str, data: Dict = None) -> str:
    """Quick function to generate scenario analysis slides (Bull/Bear/Base)."""
    gen = SlideGenerator(company_name)
    gen._add_cover_slide(f"{company_name} - Scenario Analysis", "Bull/Bear/Base Case Comparison")
    gen.add_scenario_analysis(data, include_cover=False)
    return gen.save(output_path)


def generate_management_vs_buyer(company_name: str, output_path: str, data: Dict = None) -> str:
    """Quick function to generate management vs buyer case comparison slides."""
    gen = SlideGenerator(company_name)
    gen._add_cover_slide(f"{company_name} - Case Comparison", "Management vs Buyer Underwriting")
    gen.add_management_vs_buyer(data, include_cover=False)
    return gen.save(output_path)


def generate_dcf_slides(company_name: str, output_path: str, data: Dict = None) -> str:
    """Quick function to generate DCF valuation slides."""
    gen = SlideGenerator(company_name)
    gen._add_cover_slide(f"{company_name} - DCF Analysis", "Discounted Cash Flow Valuation")
    gen.add_dcf_valuation(data, include_cover=False)
    return gen.save(output_path)


def generate_covenant_slides(company_name: str, output_path: str, data: Dict = None) -> str:
    """Quick function to generate covenant analysis slides."""
    gen = SlideGenerator(company_name)
    gen._add_cover_slide(f"{company_name} - Covenant Analysis", "Financial Covenant Compliance")
    gen.add_covenant_analysis(data, include_cover=False)
    return gen.save(output_path)


def generate_institutional_deck(company_name: str, output_path: str, data: Dict = None) -> str:
    """Generate all institutional-quality slides (7+ day case depth)."""
    data = data or {}
    gen = SlideGenerator(company_name)
    gen._add_cover_slide(f"{company_name} - Institutional Analysis", "Comprehensive Investment Analysis")

    # Core modules
    gen.add_company_overview(data.get('company'), include_cover=True)
    gen.add_industry_analysis(data.get('industry'), include_cover=True)
    gen.add_competitive_analysis(data.get('competitive'), include_cover=True)
    gen.add_financial_analysis(data.get('financial'), include_cover=True)
    gen.add_valuation(data.get('valuation'), include_cover=True)
    gen.add_lbo_analysis(data.get('lbo'), include_cover=True)

    # Institutional modules
    gen.add_scenario_analysis(data.get('scenarios'), include_cover=True)
    gen.add_management_vs_buyer(data.get('mgmt_vs_buyer'), include_cover=True)
    gen.add_dcf_valuation(data.get('dcf'), include_cover=True)
    gen.add_covenant_analysis(data.get('covenants'), include_cover=True)

    # Thesis and recommendation
    gen.add_investment_thesis(data.get('thesis'), include_cover=True)

    return gen.save(output_path)


# ============================================================
# DUE DILIGENCE & QUALITY QUICK ACCESS
# ============================================================

def generate_qoe_slides(company_name: str, output_path: str, data: Dict = None) -> str:
    """Quick function to generate Quality of Earnings slides."""
    gen = SlideGenerator(company_name)
    gen._add_cover_slide(f"{company_name} - Quality of Earnings", "EBITDA Normalization & Run-Rate Analysis")
    gen.add_quality_of_earnings(data, include_cover=False)
    return gen.save(output_path)


def generate_nwc_slides(company_name: str, output_path: str, data: Dict = None) -> str:
    """Quick function to generate Working Capital analysis slides."""
    gen = SlideGenerator(company_name)
    gen._add_cover_slide(f"{company_name} - Working Capital Analysis", "NWC Components & Peg Mechanism")
    gen.add_working_capital_analysis(data, include_cover=False)
    return gen.save(output_path)


def generate_customer_quality_slides(company_name: str, output_path: str, data: Dict = None) -> str:
    """Quick function to generate Customer/Revenue Quality slides."""
    gen = SlideGenerator(company_name)
    gen._add_cover_slide(f"{company_name} - Customer Quality", "Concentration, Retention & Unit Economics")
    gen.add_customer_quality_analysis(data, include_cover=False)
    return gen.save(output_path)


def generate_credit_slides(company_name: str, output_path: str, data: Dict = None) -> str:
    """Quick function to generate Credit Analysis slides."""
    gen = SlideGenerator(company_name)
    gen._add_cover_slide(f"{company_name} - Credit Analysis", "Debt Capacity & Stress Testing")
    gen.add_credit_analysis_slides(data, include_cover=False)
    return gen.save(output_path)


# ============================================================
# RETURNS & CAPITAL STRUCTURE QUICK ACCESS
# ============================================================

def generate_dividend_recap_slides(company_name: str, output_path: str, data: Dict = None) -> str:
    """Quick function to generate Dividend Recap slides."""
    gen = SlideGenerator(company_name)
    gen._add_cover_slide(f"{company_name} - Dividend Recapitalization", "Capital Return Analysis")
    gen.add_dividend_recap_slides(data, include_cover=False)
    return gen.save(output_path)


def generate_refinancing_slides(company_name: str, output_path: str, data: Dict = None) -> str:
    """Quick function to generate Refinancing Analysis slides."""
    gen = SlideGenerator(company_name)
    gen._add_cover_slide(f"{company_name} - Refinancing Analysis", "Rate Savings & Cost-Benefit")
    gen.add_refinancing_slides(data, include_cover=False)
    return gen.save(output_path)


def generate_waterfall_slides(company_name: str, output_path: str, data: Dict = None) -> str:
    """Quick function to generate Equity Waterfall slides."""
    gen = SlideGenerator(company_name)
    gen._add_cover_slide(f"{company_name} - Equity Waterfall", "Cap Table & Distribution Analysis")
    gen.add_waterfall_slides(data, include_cover=False)
    return gen.save(output_path)


def generate_sponsor_economics_slides(company_name: str, output_path: str, data: Dict = None) -> str:
    """Quick function to generate Sponsor Economics slides."""
    gen = SlideGenerator(company_name)
    gen._add_cover_slide(f"{company_name} - Sponsor Economics", "GP Carry & Fund Returns")
    gen.add_sponsor_economics_slides(data, include_cover=False)
    return gen.save(output_path)


# ============================================================
# TRANSACTION STRUCTURE QUICK ACCESS
# ============================================================

def generate_addon_slides(company_name: str, output_path: str, data: Dict = None) -> str:
    """Quick function to generate Add-on / Bolt-on analysis slides."""
    gen = SlideGenerator(company_name)
    gen._add_cover_slide(f"{company_name} - Add-on Analysis", "Platform + Bolt-on Combination")
    gen.add_addon_analysis_slides(data, include_cover=False)
    return gen.save(output_path)


def generate_synergy_slides(company_name: str, output_path: str, data: Dict = None) -> str:
    """Quick function to generate Synergy Analysis slides."""
    gen = SlideGenerator(company_name)
    gen._add_cover_slide(f"{company_name} - Synergy Analysis", "Revenue & Cost Synergies")
    gen.add_synergy_slides(data, include_cover=False)
    return gen.save(output_path)


def generate_carveout_slides(company_name: str, output_path: str, data: Dict = None) -> str:
    """Quick function to generate Carve-out Analysis slides."""
    gen = SlideGenerator(company_name)
    gen._add_cover_slide(f"{company_name} - Carve-out Analysis", "Standalone Economics & Separation")
    gen.add_carveout_slides(data, include_cover=False)
    return gen.save(output_path)


def generate_earnout_slides(company_name: str, output_path: str, data: Dict = None) -> str:
    """Quick function to generate Earnout / Contingent Consideration slides."""
    gen = SlideGenerator(company_name)
    gen._add_cover_slide(f"{company_name} - Earnout Analysis", "Contingent Consideration & Milestones")
    gen.add_earnout_slides(data, include_cover=False)
    return gen.save(output_path)


def generate_ppa_slides(company_name: str, output_path: str, data: Dict = None) -> str:
    """Quick function to generate Purchase Price Allocation slides."""
    gen = SlideGenerator(company_name)
    gen._add_cover_slide(f"{company_name} - Purchase Price Allocation", "Asset Valuation & Goodwill")
    gen.add_ppa_slides(data, include_cover=False)
    return gen.save(output_path)


# ============================================================
# VALUE CREATION QUICK ACCESS
# ============================================================

def generate_value_creation_slides(company_name: str, output_path: str, data: Dict = None) -> str:
    """Quick function to generate Value Creation Bridge slides."""
    gen = SlideGenerator(company_name)
    gen._add_cover_slide(f"{company_name} - Value Creation", "Equity Value Bridge Analysis")
    gen.add_value_creation_slides(data, include_cover=False)
    return gen.save(output_path)


def generate_hundred_day_slides(company_name: str, output_path: str, data: Dict = None) -> str:
    """Quick function to generate 100-Day Plan slides."""
    gen = SlideGenerator(company_name)
    gen._add_cover_slide(f"{company_name} - 100-Day Plan", "Post-Close Value Creation Priorities")
    gen.add_hundred_day_slides(data, include_cover=False)
    return gen.save(output_path)


def generate_exit_readiness_slides(company_name: str, output_path: str, data: Dict = None) -> str:
    """Quick function to generate Exit Readiness Assessment slides."""
    gen = SlideGenerator(company_name)
    gen._add_cover_slide(f"{company_name} - Exit Readiness", "Exit Options & Readiness Assessment")
    gen.add_exit_readiness_slides(data, include_cover=False)
    return gen.save(output_path)


def generate_mip_slides(company_name: str, output_path: str, data: Dict = None) -> str:
    """Quick function to generate Management Incentive Plan slides."""
    gen = SlideGenerator(company_name)
    gen._add_cover_slide(f"{company_name} - Management Incentive Plan", "MIP Structure & Payout Scenarios")
    gen.add_mip_slides(data, include_cover=False)
    return gen.save(output_path)


# ============================================================
# SPECIALIZED QUICK ACCESS
# ============================================================

def generate_rollup_slides(company_name: str, output_path: str, data: Dict = None) -> str:
    """Quick function to generate Rollup / Platform Build slides."""
    gen = SlideGenerator(company_name)
    gen._add_cover_slide(f"{company_name} - Rollup Strategy", "Platform Build & Acquisition Pipeline")
    gen.add_rollup_slides(data, include_cover=False)
    return gen.save(output_path)


def generate_tax_analysis_slides(company_name: str, output_path: str, data: Dict = None) -> str:
    """Quick function to generate Tax Analysis slides."""
    gen = SlideGenerator(company_name)
    gen._add_cover_slide(f"{company_name} - Tax Analysis", "Tax Structure & Shield Value")
    gen.add_tax_slides(data, include_cover=False)
    return gen.save(output_path)


def generate_control_premium_slides(company_name: str, output_path: str, data: Dict = None) -> str:
    """Quick function to generate Control Premium Analysis slides."""
    gen = SlideGenerator(company_name)
    gen._add_cover_slide(f"{company_name} - Control Premium", "Premium Analysis & Precedents")
    gen.add_control_premium_slides(data, include_cover=False)
    return gen.save(output_path)


# ============================================================
# CFA VISUAL SLIDES QUICK ACCESS
# ============================================================

def generate_wmbt_slide(company_name: str, output_path: str, data: Dict = None) -> str:
    """Quick function to generate What Must Be True slide."""
    gen = SlideGenerator(company_name)
    gen._add_cover_slide(f"{company_name} - Investment Thesis", "What Must Be True Analysis")
    gen.add_wmbt_slide(data, include_cover=False)
    return gen.save(output_path)


def generate_risk_mitigant_slide(company_name: str, output_path: str, data: Dict = None) -> str:
    """Quick function to generate Risk-Mitigant Pairing slide."""
    gen = SlideGenerator(company_name)
    gen._add_cover_slide(f"{company_name} - Risk Assessment", "Risk-Mitigant Pairing")
    gen.add_risk_mitigant_slide(data, include_cover=False)
    return gen.save(output_path)


def generate_valuation_blending_slide(company_name: str, output_path: str, data: Dict = None) -> str:
    """Quick function to generate Valuation Blending slide."""
    gen = SlideGenerator(company_name)
    gen._add_cover_slide(f"{company_name} - Valuation", "Blended Valuation Summary")
    gen.add_valuation_blending_slide(data, include_cover=False)
    return gen.save(output_path)


def generate_value_chain_slide(company_name: str, output_path: str, data: Dict = None) -> str:
    """Quick function to generate Value Chain Analysis slide."""
    gen = SlideGenerator(company_name)
    gen._add_cover_slide(f"{company_name} - Value Chain", "Industry Value Chain Analysis")
    gen.add_value_chain_slide(data, include_cover=False)
    return gen.save(output_path)


def generate_risk_matrix_slide(company_name: str, output_path: str, data: Dict = None) -> str:
    """Quick function to generate Risk Matrix slide."""
    gen = SlideGenerator(company_name)
    gen._add_cover_slide(f"{company_name} - Risk Matrix", "Probability-Impact Assessment")
    gen.add_risk_matrix_slide(data, include_cover=False)
    return gen.save(output_path)


def generate_tam_funnel_slide(company_name: str, output_path: str, data: Dict = None) -> str:
    """Quick function to generate TAM/SAM/SOM Funnel slide."""
    gen = SlideGenerator(company_name)
    gen._add_cover_slide(f"{company_name} - Market Sizing", "TAM / SAM / SOM Analysis")
    gen.add_tam_funnel_slide(data, include_cover=False)
    return gen.save(output_path)


def generate_pestel_slide(company_name: str, output_path: str, data: Dict = None) -> str:
    """Quick function to generate PESTEL Analysis slide."""
    gen = SlideGenerator(company_name)
    gen._add_cover_slide(f"{company_name} - PESTEL", "Macro Environment Analysis")
    gen.add_pestel_slide(data, include_cover=False)
    return gen.save(output_path)


def generate_porters_radar_slide(company_name: str, output_path: str, data: Dict = None) -> str:
    """Quick function to generate Porter's Five Forces Radar slide."""
    gen = SlideGenerator(company_name)
    gen._add_cover_slide(f"{company_name} - Porter's Forces", "Five Forces Quantified Radar")
    gen.add_porters_radar_slide(data, include_cover=False)
    return gen.save(output_path)


# ============================================================
# CFA EXCEL+SLIDE PAIRS QUICK ACCESS
# ============================================================

def generate_tornado_sensitivity_slide(company_name: str, output_path: str, data: Dict = None) -> str:
    """Quick function to generate Tornado Sensitivity Chart slide."""
    gen = SlideGenerator(company_name)
    gen._add_cover_slide(f"{company_name} - Sensitivity", "Tornado Sensitivity Analysis")
    gen.add_tornado_sensitivity_slide(data, include_cover=False)
    return gen.save(output_path)


def generate_football_field_slide(company_name: str, output_path: str, data: Dict = None) -> str:
    """Quick function to generate Football Field Valuation slide."""
    gen = SlideGenerator(company_name)
    gen._add_cover_slide(f"{company_name} - Valuation", "Football Field Analysis")
    gen.add_football_field_slide(data, include_cover=False)
    return gen.save(output_path)


def generate_sotp_waterfall_slide(company_name: str, output_path: str, data: Dict = None) -> str:
    """Quick function to generate SOTP Waterfall slide."""
    gen = SlideGenerator(company_name)
    gen._add_cover_slide(f"{company_name} - SOTP", "Sum-of-the-Parts Waterfall")
    gen.add_sotp_waterfall_slide(data, include_cover=False)
    return gen.save(output_path)


# ============================================================
# MODULE LISTING
# ============================================================

def list_available_modules() -> Dict:
    """Return the registry of available slide modules."""
    return SLIDE_MODULES


def get_module_info(module_name: str) -> Dict:
    """Get information about a specific module."""
    return SLIDE_MODULES.get(module_name, {})


if __name__ == "__main__":
    print("=" * 60)
    print("Modular Slide Generator - Available Modules")
    print("=" * 60)

    for key, info in SLIDE_MODULES.items():
        print(f"\n{info['name']} ({key})")
        print(f"  {info['description']}")
        print(f"  Slides: {', '.join(info['slides'])}")

    print("\n" + "=" * 60)
    print("Quick Generation Functions:")
    print("=" * 60)
    print("  generate_industry_analysis(company, path)")
    print("  generate_competitive_analysis(company, path)")
    print("  generate_financial_analysis(company, path)")
    print("  generate_debt_analysis(company, path)")
    print("  generate_lbo_analysis(company, path)")
    print("  generate_management_analysis(company, path)")
    print("  generate_valuation_analysis(company, path)")
    print("  generate_investment_thesis(company, path)")
    print("  generate_full_deck(company, path)")

    print("\n" + "=" * 60)
    print("Institutional Functions (7+ Day Case):")
    print("=" * 60)
    print("  generate_scenario_analysis(company, path)")
    print("  generate_management_vs_buyer(company, path)")
    print("  generate_dcf_slides(company, path)")
    print("  generate_covenant_slides(company, path)")
    print("  generate_institutional_deck(company, path)")

    print("\n" + "=" * 60)
    print("CFA Visual Slides:")
    print("=" * 60)
    print("  generate_wmbt_slide(company, path)")
    print("  generate_risk_mitigant_slide(company, path)")
    print("  generate_valuation_blending_slide(company, path)")
    print("  generate_value_chain_slide(company, path)")
    print("  generate_risk_matrix_slide(company, path)")
    print("  generate_tam_funnel_slide(company, path)")
    print("  generate_pestel_slide(company, path)")
    print("  generate_porters_radar_slide(company, path)")

    print("\n" + "=" * 60)
    print("CFA Excel+Slide Pairs:")
    print("=" * 60)
    print("  generate_tornado_sensitivity_slide(company, path)")
    print("  generate_football_field_slide(company, path)")
    print("  generate_sotp_waterfall_slide(company, path)")
