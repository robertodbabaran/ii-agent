#!/usr/bin/env python3
"""
CFA-Derived Excel Modules Mixin.

Contains modules based on CFA Research Challenge patterns:
Trading Comps, Transaction Comps, Three-Statement, Reverse DCF,
DuPont Analysis, ROIC Decomposition, DDM, Earnings Quality Scores,
Tornado Sensitivity, Football Field, SOTP Valuation, Enhanced WACC,
Geographic Terminal Growth, Multi-Stage DCF.

All modules write live Excel formulas and reference upstream modules
(Assumptions, Operating Model, WACC, etc.) via self.cell_map when
available.  Falls back to hardcoded inputs when upstream modules are
absent — no behaviour change for standalone usage.

Sprint 4 additions (2026-02-15): Tornado Sensitivity, Football Field,
SOTP Valuation, Enhanced WACC, Geographic Terminal Growth, Multi-Stage DCF.
Conditional formatting: data bars, RAG color scales.
"""

from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from .formula_builder import FormulaBuilder as FB
from ._base import (INPUT_FILL, INPUT_FONT, HEADER_FILL, HEADER_FONT,
                     BOTTOM_BORDER, DOUBLE_BORDER, ModelDepth)
from typing import Dict, List


class CFAModulesMixin:
    """Mixin providing CFA-derived analysis modules."""

    # ============================================================
    # MODULE: TRADING COMPS
    # ============================================================

    def add_trading_comps(self, data: Dict = None) -> 'CFAModulesMixin':
        """Add Public Comparable Companies (Trading Comps) analysis.

        Smart refs: Assumptions (LTM revenue, EBITDA for implied valuation).
        """
        ws = self.wb.create_sheet("Trading Comps")
        self.sheets_created.append("Trading Comps")
        data = data or {}

        self._add_title(ws, f"{self.company_name} - Public Comparable Companies", 1, 1)

        # --- Company metrics row (from Assumptions or data) ---
        ws.cell(row=3, column=1, value=f"{self.company_name} Metrics")
        ws.cell(row=3, column=1).font = Font(bold=True)

        amap = self.cell_map.get('assumptions', {})
        target_rev_row = 4
        ws.cell(row=target_rev_row, column=1, value="LTM Revenue ($M)")
        if amap.get('ltm_revenue'):
            ws.cell(row=target_rev_row, column=2,
                    value=f"='Assumptions'!{amap['ltm_revenue']}")
        else:
            ws.cell(row=target_rev_row, column=2, value=self.assumptions.ltm_revenue)
            self._format_input_cell(ws, target_rev_row, 2)
        ws.cell(row=target_rev_row, column=2).number_format = '#,##0.0'

        target_ebitda_row = 5
        ws.cell(row=target_ebitda_row, column=1, value="LTM EBITDA ($M)")
        if amap.get('ltm_ebitda'):
            ws.cell(row=target_ebitda_row, column=2,
                    value=f"='Assumptions'!{amap['ltm_ebitda']}")
        else:
            ws.cell(row=target_ebitda_row, column=2, value=self.assumptions.ltm_ebitda)
            self._format_input_cell(ws, target_ebitda_row, 2)
        ws.cell(row=target_ebitda_row, column=2).number_format = '#,##0.0'

        # --- Comparable Companies Table ---
        self._add_section_header(ws, "COMPARABLE COMPANIES", 7, 1)
        headers = ["Company", "EV ($M)", "Revenue ($M)", "EBITDA ($M)",
                    "EV/Revenue", "EV/EBITDA", "P/E", "Rev Growth %",
                    "EBITDA Margin %"]
        hdr_row = 8
        for i, h in enumerate(headers):
            ws.cell(row=hdr_row, column=1 + i, value=h)
        self._format_header_row(ws, hdr_row, 1, len(headers))

        comps = data.get('comps', [
            {'name': 'Comp A', 'ev': 500, 'revenue': 200, 'ebitda': 50, 'pe': 18.0, 'growth': 0.08},
            {'name': 'Comp B', 'ev': 800, 'revenue': 350, 'ebitda': 80, 'pe': 22.0, 'growth': 0.12},
            {'name': 'Comp C', 'ev': 300, 'revenue': 120, 'ebitda': 30, 'pe': 15.0, 'growth': 0.05},
            {'name': 'Comp D', 'ev': 600, 'revenue': 250, 'ebitda': 60, 'pe': 20.0, 'growth': 0.10},
            {'name': 'Comp E', 'ev': 450, 'revenue': 180, 'ebitda': 45, 'pe': 17.0, 'growth': 0.07},
        ])
        comp_start = hdr_row + 1
        for i, comp in enumerate(comps):
            row = comp_start + i
            ws.cell(row=row, column=1, value=comp['name'])
            ws.cell(row=row, column=2, value=comp['ev'])
            ws.cell(row=row, column=2).number_format = '#,##0.0'
            self._format_input_cell(ws, row, 2)
            ws.cell(row=row, column=3, value=comp['revenue'])
            ws.cell(row=row, column=3).number_format = '#,##0.0'
            self._format_input_cell(ws, row, 3)
            ws.cell(row=row, column=4, value=comp['ebitda'])
            ws.cell(row=row, column=4).number_format = '#,##0.0'
            self._format_input_cell(ws, row, 4)
            # EV/Revenue = B/C
            ws.cell(row=row, column=5, value=f"=IFERROR(B{row}/C{row},0)")
            ws.cell(row=row, column=5).number_format = '0.0x'
            # EV/EBITDA = B/D
            ws.cell(row=row, column=6, value=f"=IFERROR(B{row}/D{row},0)")
            ws.cell(row=row, column=6).number_format = '0.0x'
            # P/E
            ws.cell(row=row, column=7, value=comp['pe'])
            ws.cell(row=row, column=7).number_format = '0.0x'
            self._format_input_cell(ws, row, 7)
            # Revenue Growth
            ws.cell(row=row, column=8, value=comp['growth'])
            ws.cell(row=row, column=8).number_format = '0.0%'
            self._format_input_cell(ws, row, 8)
            # EBITDA Margin = D/C
            ws.cell(row=row, column=9, value=f"=IFERROR(D{row}/C{row},0)")
            ws.cell(row=row, column=9).number_format = '0.0%'
        comp_end = comp_start + len(comps) - 1

        # --- Statistics rows ---
        stats_row = comp_end + 2
        stats = [
            ("Mean", "AVERAGE"),
            ("Median", "MEDIAN"),
            ("High", "MAX"),
            ("Low", "MIN"),
        ]
        for i, (label, func) in enumerate(stats):
            row = stats_row + i
            ws.cell(row=row, column=1, value=label)
            ws.cell(row=row, column=1).font = Font(bold=True)
            for col in [5, 6, 7, 8, 9]:
                cl = get_column_letter(col)
                ws.cell(row=row, column=col,
                        value=f"={func}({cl}{comp_start}:{cl}{comp_end})")
                fmt = '0.0x' if col <= 7 else '0.0%'
                ws.cell(row=row, column=col).number_format = fmt
                ws.cell(row=row, column=col).font = Font(bold=True)

        mean_row = stats_row
        median_row = stats_row + 1

        # --- Implied Valuation ---
        impl_row = stats_row + len(stats) + 2
        self._add_section_header(ws, "IMPLIED VALUATION", impl_row, 1)
        impl_row += 1

        ws.cell(row=impl_row, column=1, value="Method")
        ws.cell(row=impl_row, column=2, value="Multiple")
        ws.cell(row=impl_row, column=3, value="Metric")
        ws.cell(row=impl_row, column=4, value="Implied EV")
        self._format_header_row(ws, impl_row, 1, 4)
        impl_row += 1

        # EV/Revenue (mean)
        ws.cell(row=impl_row, column=1, value="EV/Revenue (Mean)")
        ws.cell(row=impl_row, column=2, value=f"=E{mean_row}")
        ws.cell(row=impl_row, column=2).number_format = '0.0x'
        ws.cell(row=impl_row, column=3, value=f"=B{target_rev_row}")
        ws.cell(row=impl_row, column=3).number_format = '#,##0.0'
        ws.cell(row=impl_row, column=4, value=f"=B{impl_row}*C{impl_row}")
        ws.cell(row=impl_row, column=4).number_format = '#,##0.0'
        ev_rev_row = impl_row
        impl_row += 1

        # EV/Revenue (median)
        ws.cell(row=impl_row, column=1, value="EV/Revenue (Median)")
        ws.cell(row=impl_row, column=2, value=f"=E{median_row}")
        ws.cell(row=impl_row, column=2).number_format = '0.0x'
        ws.cell(row=impl_row, column=3, value=f"=B{target_rev_row}")
        ws.cell(row=impl_row, column=3).number_format = '#,##0.0'
        ws.cell(row=impl_row, column=4, value=f"=B{impl_row}*C{impl_row}")
        ws.cell(row=impl_row, column=4).number_format = '#,##0.0'
        impl_row += 1

        # EV/EBITDA (mean)
        ws.cell(row=impl_row, column=1, value="EV/EBITDA (Mean)")
        ws.cell(row=impl_row, column=2, value=f"=F{mean_row}")
        ws.cell(row=impl_row, column=2).number_format = '0.0x'
        ws.cell(row=impl_row, column=3, value=f"=B{target_ebitda_row}")
        ws.cell(row=impl_row, column=3).number_format = '#,##0.0'
        ws.cell(row=impl_row, column=4, value=f"=B{impl_row}*C{impl_row}")
        ws.cell(row=impl_row, column=4).number_format = '#,##0.0'
        ev_ebitda_mean_row = impl_row
        impl_row += 1

        # EV/EBITDA (median)
        ws.cell(row=impl_row, column=1, value="EV/EBITDA (Median)")
        ws.cell(row=impl_row, column=2, value=f"=F{median_row}")
        ws.cell(row=impl_row, column=2).number_format = '0.0x'
        ws.cell(row=impl_row, column=3, value=f"=B{target_ebitda_row}")
        ws.cell(row=impl_row, column=3).number_format = '#,##0.0'
        ws.cell(row=impl_row, column=4, value=f"=B{impl_row}*C{impl_row}")
        ws.cell(row=impl_row, column=4).number_format = '#,##0.0'
        ev_ebitda_median_row = impl_row

        # Range summary
        impl_row += 2
        ws.cell(row=impl_row, column=1, value="Implied EV Range")
        ws.cell(row=impl_row, column=1).font = Font(bold=True)
        ws.cell(row=impl_row, column=2, value="Low")
        ws.cell(row=impl_row, column=3,
                value=f"=MIN(D{ev_rev_row}:D{ev_ebitda_median_row})")
        ws.cell(row=impl_row, column=3).number_format = '#,##0.0'
        ws.cell(row=impl_row, column=3).font = Font(bold=True)
        impl_row += 1
        ws.cell(row=impl_row, column=2, value="High")
        ws.cell(row=impl_row, column=3,
                value=f"=MAX(D{ev_rev_row}:D{ev_ebitda_median_row})")
        ws.cell(row=impl_row, column=3).number_format = '#,##0.0'
        ws.cell(row=impl_row, column=3).font = Font(bold=True)
        ws.cell(row=impl_row, column=3).fill = PatternFill(
            start_color="90EE90", end_color="90EE90", fill_type="solid")

        # Column widths
        ws.column_dimensions['A'].width = 22
        for c in ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I']:
            ws.column_dimensions[c].width = 14

        self.cell_map['trading_comps'] = {
            'comp_start': comp_start,
            'comp_end': comp_end,
            'mean_row': mean_row,
            'median_row': median_row,
            'ev_rev_mean': f'E{mean_row}',
            'ev_ebitda_mean': f'F{mean_row}',
            'ev_rev_median': f'E{median_row}',
            'ev_ebitda_median': f'F{median_row}',
            'implied_ev_ebitda_mean': f'D{ev_ebitda_mean_row}',
        }

        return self

    # ============================================================
    # MODULE: TRANSACTION COMPS
    # ============================================================

    def add_transaction_comps(self, data: Dict = None) -> 'CFAModulesMixin':
        """Add Precedent Transactions analysis.

        Smart refs: Assumptions (LTM financials for implied valuation).
        """
        ws = self.wb.create_sheet("Transaction Comps")
        self.sheets_created.append("Transaction Comps")
        data = data or {}

        self._add_title(ws, f"{self.company_name} - Precedent Transactions", 1, 1)

        # --- Precedent Transactions Table ---
        self._add_section_header(ws, "PRECEDENT TRANSACTIONS", 3, 1)
        headers = ["Date", "Acquirer", "Target", "EV ($M)", "EBITDA ($M)",
                    "EV/EBITDA", "EV/Revenue", "Premium (%)"]
        hdr_row = 4
        for i, h in enumerate(headers):
            ws.cell(row=hdr_row, column=1 + i, value=h)
        self._format_header_row(ws, hdr_row, 1, len(headers))

        transactions = data.get('transactions', [
            {'date': '2024', 'acquirer': 'Buyer 1', 'target': 'Target A',
             'ev': 400, 'ebitda': 50, 'revenue': 180, 'premium': 0.30},
            {'date': '2024', 'acquirer': 'Buyer 2', 'target': 'Target B',
             'ev': 600, 'ebitda': 70, 'revenue': 250, 'premium': 0.35},
            {'date': '2023', 'acquirer': 'Buyer 3', 'target': 'Target C',
             'ev': 350, 'ebitda': 40, 'revenue': 150, 'premium': 0.25},
            {'date': '2023', 'acquirer': 'Buyer 4', 'target': 'Target D',
             'ev': 550, 'ebitda': 65, 'revenue': 220, 'premium': 0.28},
            {'date': '2022', 'acquirer': 'Buyer 5', 'target': 'Target E',
             'ev': 450, 'ebitda': 55, 'revenue': 190, 'premium': 0.32},
        ])
        txn_start = hdr_row + 1
        for i, txn in enumerate(transactions):
            row = txn_start + i
            ws.cell(row=row, column=1, value=txn['date'])
            ws.cell(row=row, column=2, value=txn['acquirer'])
            ws.cell(row=row, column=3, value=txn['target'])
            ws.cell(row=row, column=4, value=txn['ev'])
            ws.cell(row=row, column=4).number_format = '#,##0.0'
            self._format_input_cell(ws, row, 4)
            ws.cell(row=row, column=5, value=txn['ebitda'])
            ws.cell(row=row, column=5).number_format = '#,##0.0'
            self._format_input_cell(ws, row, 5)
            # EV/EBITDA = D/E
            ws.cell(row=row, column=6, value=f"=IFERROR(D{row}/E{row},0)")
            ws.cell(row=row, column=6).number_format = '0.0x'
            # EV/Revenue
            revenue = txn.get('revenue', 0)
            ws.cell(row=row, column=7,
                    value=f"=IFERROR(D{row}/{revenue},0)" if revenue else 0)
            ws.cell(row=row, column=7).number_format = '0.0x'
            # Premium
            ws.cell(row=row, column=8, value=txn['premium'])
            ws.cell(row=row, column=8).number_format = '0.0%'
            self._format_input_cell(ws, row, 8)
        txn_end = txn_start + len(transactions) - 1

        # --- Statistics ---
        stats_row = txn_end + 2
        stats = [("Mean", "AVERAGE"), ("Median", "MEDIAN"),
                 ("High", "MAX"), ("Low", "MIN")]
        for i, (label, func) in enumerate(stats):
            row = stats_row + i
            ws.cell(row=row, column=1, value=label)
            ws.cell(row=row, column=1).font = Font(bold=True)
            for col in [6, 7, 8]:
                cl = get_column_letter(col)
                ws.cell(row=row, column=col,
                        value=f"={func}({cl}{txn_start}:{cl}{txn_end})")
                fmt = '0.0%' if col == 8 else '0.0x'
                ws.cell(row=row, column=col).number_format = fmt
                ws.cell(row=row, column=col).font = Font(bold=True)
        mean_row = stats_row
        median_row = stats_row + 1

        # --- Implied Valuation ---
        impl_hdr = stats_row + len(stats) + 2
        self._add_section_header(ws, "IMPLIED VALUATION", impl_hdr, 1)
        impl_hdr += 1
        ws.cell(row=impl_hdr, column=1, value="Method")
        ws.cell(row=impl_hdr, column=2, value="Multiple")
        ws.cell(row=impl_hdr, column=3, value="Metric")
        ws.cell(row=impl_hdr, column=4, value="Implied EV")
        self._format_header_row(ws, impl_hdr, 1, 4)

        amap = self.cell_map.get('assumptions', {})
        ebitda_ref = f"='Assumptions'!{amap['ltm_ebitda']}" if amap.get('ltm_ebitda') else self.assumptions.ltm_ebitda

        r = impl_hdr + 1
        ws.cell(row=r, column=1, value="EV/EBITDA (Mean)")
        ws.cell(row=r, column=2, value=f"=F{mean_row}")
        ws.cell(row=r, column=2).number_format = '0.0x'
        ws.cell(row=r, column=3, value=ebitda_ref)
        ws.cell(row=r, column=3).number_format = '#,##0.0'
        ws.cell(row=r, column=4, value=f"=B{r}*C{r}")
        ws.cell(row=r, column=4).number_format = '#,##0.0'
        ev_mean_row = r

        r += 1
        ws.cell(row=r, column=1, value="EV/EBITDA (Median)")
        ws.cell(row=r, column=2, value=f"=F{median_row}")
        ws.cell(row=r, column=2).number_format = '0.0x'
        ws.cell(row=r, column=3, value=ebitda_ref)
        ws.cell(row=r, column=3).number_format = '#,##0.0'
        ws.cell(row=r, column=4, value=f"=B{r}*C{r}")
        ws.cell(row=r, column=4).number_format = '#,##0.0'
        ev_median_row = r

        r += 2
        ws.cell(row=r, column=1, value="Implied EV Range")
        ws.cell(row=r, column=1).font = Font(bold=True)
        ws.cell(row=r, column=3, value=f"=MIN(D{ev_mean_row}:D{ev_median_row})")
        ws.cell(row=r, column=3).number_format = '#,##0.0'
        ws.cell(row=r, column=3).font = Font(bold=True)
        ws.cell(row=r, column=4, value=f"=MAX(D{ev_mean_row}:D{ev_median_row})")
        ws.cell(row=r, column=4).number_format = '#,##0.0'
        ws.cell(row=r, column=4).font = Font(bold=True)
        ws.cell(row=r, column=4).fill = PatternFill(
            start_color="90EE90", end_color="90EE90", fill_type="solid")

        ws.column_dimensions['A'].width = 12
        ws.column_dimensions['B'].width = 16
        ws.column_dimensions['C'].width = 16
        for c in ['D', 'E', 'F', 'G', 'H']:
            ws.column_dimensions[c].width = 14

        self.cell_map['transaction_comps'] = {
            'txn_start': txn_start,
            'txn_end': txn_end,
            'mean_row': mean_row,
            'median_row': median_row,
            'ev_ebitda_mean': f'F{mean_row}',
            'ev_ebitda_median': f'F{median_row}',
            'premium_mean': f'H{mean_row}',
            'implied_ev_mean': f'D{ev_mean_row}',
        }

        return self

    # ============================================================
    # MODULE: THREE-STATEMENT MODEL
    # ============================================================

    def add_three_statement(self, data: Dict = None) -> 'CFAModulesMixin':
        """Add linked Three-Statement Model (IS, BS, CF).

        Smart refs: Operating Model (revenue, EBITDA), Debt Schedule,
        Working Capital, Assumptions.
        """
        ws = self.wb.create_sheet("3-Statement")
        self.sheets_created.append("3-Statement")
        a = self.assumptions
        data = data or {}
        n = self.projection_years

        self._add_title(ws, f"{self.company_name} - Three-Statement Model", 1, 1)

        # Year headers
        yr_hdr_row = 3
        ws.cell(row=yr_hdr_row, column=1, value="($M)")
        ws.cell(row=yr_hdr_row, column=2, value="LTM")
        for i in range(n):
            ws.cell(row=yr_hdr_row, column=3 + i, value=f"Year {i + 1}")
        self._format_header_row(ws, yr_hdr_row, 1, 2 + n)

        # Column refs
        ltm_col = 2
        yr1_col = 3

        # ============ INCOME STATEMENT ============
        is_start = yr_hdr_row + 1
        self._add_section_header(ws, "INCOME STATEMENT", is_start, 1)

        # Revenue (row is_start+1)
        rev_row = is_start + 1
        ws.cell(row=rev_row, column=1, value="Revenue")
        om = self.cell_map.get('operating_model', {})
        if om.get('revenue_row') and om.get('ltm_col'):
            # Pull LTM from Operating Model
            om_ltm_cl = get_column_letter(om['ltm_col'])
            ws.cell(row=rev_row, column=ltm_col,
                    value=f"='Operating Model'!{om_ltm_cl}{om['revenue_row']}")
        else:
            ws.cell(row=rev_row, column=ltm_col, value=a.ltm_revenue)
            self._format_input_cell(ws, rev_row, ltm_col)
        ws.cell(row=rev_row, column=ltm_col).number_format = '#,##0.0'
        # Projected revenue
        for i in range(n):
            col = yr1_col + i
            prev_cl = get_column_letter(col - 1)
            growth_cl = get_column_letter(3 + i)  # Assumptions growth row
            amap = self.cell_map.get('assumptions', {})
            if amap.get('rev_growth_row'):
                ws.cell(row=rev_row, column=col,
                        value=f"={prev_cl}{rev_row}*(1+'Assumptions'!{growth_cl}{amap['rev_growth_row']})")
            else:
                g = a.revenue_growth[i] if i < len(a.revenue_growth) else 0.05
                ws.cell(row=rev_row, column=col,
                        value=f"={prev_cl}{rev_row}*(1+{g})")
            ws.cell(row=rev_row, column=col).number_format = '#,##0.0'

        # COGS
        cogs_row = rev_row + 1
        ws.cell(row=cogs_row, column=1, value="COGS")
        cogs_pct = data.get('cogs_pct', 0.60)
        ws.cell(row=cogs_row, column=ltm_col, value=a.ltm_revenue * cogs_pct * -1)
        ws.cell(row=cogs_row, column=ltm_col).number_format = '#,##0.0'
        self._format_input_cell(ws, cogs_row, ltm_col)
        for i in range(n):
            col = yr1_col + i
            cl = get_column_letter(col)
            ws.cell(row=cogs_row, column=col,
                    value=f"=-{cl}{rev_row}*{cogs_pct}")
            ws.cell(row=cogs_row, column=col).number_format = '#,##0.0'

        # Gross Profit
        gp_row = cogs_row + 1
        ws.cell(row=gp_row, column=1, value="Gross Profit")
        ws.cell(row=gp_row, column=1).font = Font(bold=True)
        for col in range(ltm_col, yr1_col + n):
            cl = get_column_letter(col)
            ws.cell(row=gp_row, column=col,
                    value=f"={cl}{rev_row}+{cl}{cogs_row}")
            ws.cell(row=gp_row, column=col).number_format = '#,##0.0'
            ws.cell(row=gp_row, column=col).font = Font(bold=True)
            ws.cell(row=gp_row, column=col).border = BOTTOM_BORDER

        # SG&A
        sga_row = gp_row + 1
        ws.cell(row=sga_row, column=1, value="SG&A")
        sga_pct = data.get('sga_pct', 0.20)
        for col in range(ltm_col, yr1_col + n):
            cl = get_column_letter(col)
            ws.cell(row=sga_row, column=col,
                    value=f"=-{cl}{rev_row}*{sga_pct}")
            ws.cell(row=sga_row, column=col).number_format = '#,##0.0'

        # EBITDA
        ebitda_row = sga_row + 1
        ws.cell(row=ebitda_row, column=1, value="EBITDA")
        ws.cell(row=ebitda_row, column=1).font = Font(bold=True)
        for col in range(ltm_col, yr1_col + n):
            cl = get_column_letter(col)
            ws.cell(row=ebitda_row, column=col,
                    value=f"={cl}{gp_row}+{cl}{sga_row}")
            ws.cell(row=ebitda_row, column=col).number_format = '#,##0.0'
            ws.cell(row=ebitda_row, column=col).font = Font(bold=True)
            ws.cell(row=ebitda_row, column=col).border = BOTTOM_BORDER

        # D&A
        da_row = ebitda_row + 1
        ws.cell(row=da_row, column=1, value="Depreciation & Amort")
        da_pct = data.get('da_pct', 0.03)
        for col in range(ltm_col, yr1_col + n):
            cl = get_column_letter(col)
            ws.cell(row=da_row, column=col,
                    value=f"=-{cl}{rev_row}*{da_pct}")
            ws.cell(row=da_row, column=col).number_format = '#,##0.0'

        # EBIT
        ebit_row = da_row + 1
        ws.cell(row=ebit_row, column=1, value="EBIT")
        ws.cell(row=ebit_row, column=1).font = Font(bold=True)
        for col in range(ltm_col, yr1_col + n):
            cl = get_column_letter(col)
            ws.cell(row=ebit_row, column=col,
                    value=f"={cl}{ebitda_row}+{cl}{da_row}")
            ws.cell(row=ebit_row, column=col).number_format = '#,##0.0'
            ws.cell(row=ebit_row, column=col).font = Font(bold=True)

        # Interest Expense
        int_row = ebit_row + 1
        ws.cell(row=int_row, column=1, value="Interest Expense")
        ds = self.cell_map.get('debt_schedule', {})
        for col in range(ltm_col, yr1_col + n):
            if ds.get('total_interest_row') and col >= yr1_col:
                ds_cl = get_column_letter(col)
                ws.cell(row=int_row, column=col,
                        value=f"=-'Debt Schedule'!{ds_cl}{ds['total_interest_row']}")
            else:
                cl = get_column_letter(col)
                ws.cell(row=int_row, column=col,
                        value=f"=-{cl}{rev_row}*{a.senior_interest_rate * (a.senior_debt_multiple + a.sub_debt_multiple) / a.entry_multiple:.4f}")
            ws.cell(row=int_row, column=col).number_format = '#,##0.0'

        # EBT
        ebt_row = int_row + 1
        ws.cell(row=ebt_row, column=1, value="Earnings Before Tax")
        for col in range(ltm_col, yr1_col + n):
            cl = get_column_letter(col)
            ws.cell(row=ebt_row, column=col,
                    value=f"={cl}{ebit_row}+{cl}{int_row}")
            ws.cell(row=ebt_row, column=col).number_format = '#,##0.0'

        # Tax
        tax_row = ebt_row + 1
        ws.cell(row=tax_row, column=1, value="Income Tax")
        for col in range(ltm_col, yr1_col + n):
            cl = get_column_letter(col)
            ws.cell(row=tax_row, column=col,
                    value=f"=-MAX(0,{cl}{ebt_row})*'Assumptions'!B24")
            ws.cell(row=tax_row, column=col).number_format = '#,##0.0'

        # Net Income
        ni_row = tax_row + 1
        ws.cell(row=ni_row, column=1, value="Net Income")
        ws.cell(row=ni_row, column=1).font = Font(bold=True)
        for col in range(ltm_col, yr1_col + n):
            cl = get_column_letter(col)
            ws.cell(row=ni_row, column=col,
                    value=f"={cl}{ebt_row}+{cl}{tax_row}")
            ws.cell(row=ni_row, column=col).number_format = '#,##0.0'
            ws.cell(row=ni_row, column=col).font = Font(bold=True)
            ws.cell(row=ni_row, column=col).border = DOUBLE_BORDER

        # ============ BALANCE SHEET ============
        bs_start = ni_row + 2
        self._add_section_header(ws, "BALANCE SHEET", bs_start, 1)

        # Cash
        cash_row = bs_start + 1
        ws.cell(row=cash_row, column=1, value="Cash & Equivalents")
        ws.cell(row=cash_row, column=ltm_col, value=data.get('cash', 10.0))
        ws.cell(row=cash_row, column=ltm_col).number_format = '#,##0.0'
        self._format_input_cell(ws, cash_row, ltm_col)
        for i in range(n):
            col = yr1_col + i
            # Cash = prior cash + net income + D&A - capex - NWC change (simplified)
            prev_cl = get_column_letter(col - 1)
            cl = get_column_letter(col)
            ws.cell(row=cash_row, column=col,
                    value=f"={prev_cl}{cash_row}+{cl}{ni_row}-{cl}{da_row}+{cl}{rev_row}*'Assumptions'!B22-{cl}{rev_row}*'Assumptions'!B23")
            ws.cell(row=cash_row, column=col).number_format = '#,##0.0'

        # AR
        ar_row = cash_row + 1
        ws.cell(row=ar_row, column=1, value="Accounts Receivable")
        ar_days = data.get('ar_days', 45)
        for col in range(ltm_col, yr1_col + n):
            cl = get_column_letter(col)
            ws.cell(row=ar_row, column=col,
                    value=f"={cl}{rev_row}*{ar_days}/365")
            ws.cell(row=ar_row, column=col).number_format = '#,##0.0'

        # Inventory
        inv_row = ar_row + 1
        ws.cell(row=inv_row, column=1, value="Inventory")
        inv_days = data.get('inv_days', 30)
        for col in range(ltm_col, yr1_col + n):
            cl = get_column_letter(col)
            ws.cell(row=inv_row, column=col,
                    value=f"=-{cl}{cogs_row}*{inv_days}/365")
            ws.cell(row=inv_row, column=col).number_format = '#,##0.0'

        # PP&E
        ppe_row = inv_row + 1
        ws.cell(row=ppe_row, column=1, value="PP&E (net)")
        ws.cell(row=ppe_row, column=ltm_col,
                value=data.get('ppe', a.ltm_revenue * 0.30))
        ws.cell(row=ppe_row, column=ltm_col).number_format = '#,##0.0'
        self._format_input_cell(ws, ppe_row, ltm_col)
        for i in range(n):
            col = yr1_col + i
            prev_cl = get_column_letter(col - 1)
            cl = get_column_letter(col)
            ws.cell(row=ppe_row, column=col,
                    value=f"={prev_cl}{ppe_row}+{cl}{rev_row}*'Assumptions'!B22+{cl}{da_row}")
            ws.cell(row=ppe_row, column=col).number_format = '#,##0.0'

        # Total Assets
        ta_row = ppe_row + 1
        ws.cell(row=ta_row, column=1, value="Total Assets")
        ws.cell(row=ta_row, column=1).font = Font(bold=True)
        for col in range(ltm_col, yr1_col + n):
            cl = get_column_letter(col)
            ws.cell(row=ta_row, column=col,
                    value=f"=SUM({cl}{cash_row}:{cl}{ppe_row})")
            ws.cell(row=ta_row, column=col).number_format = '#,##0.0'
            ws.cell(row=ta_row, column=col).font = Font(bold=True)
            ws.cell(row=ta_row, column=col).border = DOUBLE_BORDER

        # AP
        ap_row = ta_row + 2
        ws.cell(row=ap_row, column=1, value="Accounts Payable")
        ap_days = data.get('ap_days', 40)
        for col in range(ltm_col, yr1_col + n):
            cl = get_column_letter(col)
            ws.cell(row=ap_row, column=col,
                    value=f"=-{cl}{cogs_row}*{ap_days}/365")
            ws.cell(row=ap_row, column=col).number_format = '#,##0.0'

        # Total Debt
        debt_row = ap_row + 1
        ws.cell(row=debt_row, column=1, value="Total Debt")
        if ds.get('total_debt_row'):
            for col in range(ltm_col, yr1_col + n):
                ds_cl = get_column_letter(col)
                ws.cell(row=debt_row, column=col,
                        value=f"='Debt Schedule'!{ds_cl}{ds['total_debt_row']}")
                ws.cell(row=debt_row, column=col).number_format = '#,##0.0'
        else:
            total_debt_val = a.ltm_ebitda * (a.senior_debt_multiple + a.sub_debt_multiple)
            ws.cell(row=debt_row, column=ltm_col, value=total_debt_val)
            ws.cell(row=debt_row, column=ltm_col).number_format = '#,##0.0'
            self._format_input_cell(ws, debt_row, ltm_col)
            for i in range(n):
                col = yr1_col + i
                prev_cl = get_column_letter(col - 1)
                ws.cell(row=debt_row, column=col,
                        value=f"={prev_cl}{debt_row}*(1-{a.senior_amortization})")
                ws.cell(row=debt_row, column=col).number_format = '#,##0.0'

        # Equity (plug)
        eq_row = debt_row + 1
        ws.cell(row=eq_row, column=1, value="Total Equity")
        for col in range(ltm_col, yr1_col + n):
            cl = get_column_letter(col)
            ws.cell(row=eq_row, column=col,
                    value=f"={cl}{ta_row}-{cl}{ap_row}-{cl}{debt_row}")
            ws.cell(row=eq_row, column=col).number_format = '#,##0.0'

        # Total L&E
        tle_row = eq_row + 1
        ws.cell(row=tle_row, column=1, value="Total Liabilities & Equity")
        ws.cell(row=tle_row, column=1).font = Font(bold=True)
        for col in range(ltm_col, yr1_col + n):
            cl = get_column_letter(col)
            ws.cell(row=tle_row, column=col,
                    value=f"={cl}{ap_row}+{cl}{debt_row}+{cl}{eq_row}")
            ws.cell(row=tle_row, column=col).number_format = '#,##0.0'
            ws.cell(row=tle_row, column=col).font = Font(bold=True)
            ws.cell(row=tle_row, column=col).border = DOUBLE_BORDER

        # Balance check
        check_row = tle_row + 1
        ws.cell(row=check_row, column=1, value="Balance Check (should = 0)")
        ws.cell(row=check_row, column=1).font = Font(italic=True, color="999999")
        for col in range(ltm_col, yr1_col + n):
            cl = get_column_letter(col)
            ws.cell(row=check_row, column=col,
                    value=f"={cl}{ta_row}-{cl}{tle_row}")
            ws.cell(row=check_row, column=col).number_format = '#,##0.0'

        # ============ CASH FLOW STATEMENT ============
        cf_start = check_row + 2
        self._add_section_header(ws, "CASH FLOW STATEMENT", cf_start, 1)

        cfo_ni_row = cf_start + 1
        ws.cell(row=cfo_ni_row, column=1, value="Net Income")
        for col in range(ltm_col, yr1_col + n):
            cl = get_column_letter(col)
            ws.cell(row=cfo_ni_row, column=col, value=f"={cl}{ni_row}")
            ws.cell(row=cfo_ni_row, column=col).number_format = '#,##0.0'

        # Add back D&A
        cfo_da_row = cfo_ni_row + 1
        ws.cell(row=cfo_da_row, column=1, value="Add: Depreciation & Amort")
        for col in range(ltm_col, yr1_col + n):
            cl = get_column_letter(col)
            ws.cell(row=cfo_da_row, column=col, value=f"=-{cl}{da_row}")
            ws.cell(row=cfo_da_row, column=col).number_format = '#,##0.0'

        # Change in NWC
        cfo_nwc_row = cfo_da_row + 1
        ws.cell(row=cfo_nwc_row, column=1, value="Change in NWC")
        ws.cell(row=cfo_nwc_row, column=ltm_col, value=0)
        for i in range(n):
            col = yr1_col + i
            prev_cl = get_column_letter(col - 1)
            cl = get_column_letter(col)
            # NWC = AR + Inv - AP; change = prior NWC - current NWC (negative = cash outflow)
            ws.cell(row=cfo_nwc_row, column=col,
                    value=f"=({prev_cl}{ar_row}+{prev_cl}{inv_row}-{prev_cl}{ap_row})-({cl}{ar_row}+{cl}{inv_row}-{cl}{ap_row})")
            ws.cell(row=cfo_nwc_row, column=col).number_format = '#,##0.0'

        # CFO Total
        cfo_total_row = cfo_nwc_row + 1
        ws.cell(row=cfo_total_row, column=1, value="Cash from Operations")
        ws.cell(row=cfo_total_row, column=1).font = Font(bold=True)
        for col in range(ltm_col, yr1_col + n):
            cl = get_column_letter(col)
            ws.cell(row=cfo_total_row, column=col,
                    value=f"={cl}{cfo_ni_row}+{cl}{cfo_da_row}+{cl}{cfo_nwc_row}")
            ws.cell(row=cfo_total_row, column=col).number_format = '#,##0.0'
            ws.cell(row=cfo_total_row, column=col).font = Font(bold=True)
            ws.cell(row=cfo_total_row, column=col).border = BOTTOM_BORDER

        # Capex
        cfi_capex_row = cfo_total_row + 1
        ws.cell(row=cfi_capex_row, column=1, value="Capital Expenditures")
        for col in range(ltm_col, yr1_col + n):
            cl = get_column_letter(col)
            ws.cell(row=cfi_capex_row, column=col,
                    value=f"=-{cl}{rev_row}*'Assumptions'!B22")
            ws.cell(row=cfi_capex_row, column=col).number_format = '#,##0.0'

        # FCF
        fcf_row = cfi_capex_row + 1
        ws.cell(row=fcf_row, column=1, value="Free Cash Flow")
        ws.cell(row=fcf_row, column=1).font = Font(bold=True)
        for col in range(ltm_col, yr1_col + n):
            cl = get_column_letter(col)
            ws.cell(row=fcf_row, column=col,
                    value=f"={cl}{cfo_total_row}+{cl}{cfi_capex_row}")
            ws.cell(row=fcf_row, column=col).number_format = '#,##0.0'
            ws.cell(row=fcf_row, column=col).font = Font(bold=True)
            ws.cell(row=fcf_row, column=col).border = DOUBLE_BORDER
            ws.cell(row=fcf_row, column=col).fill = PatternFill(
                start_color="90EE90", end_color="90EE90", fill_type="solid")

        # Column widths
        ws.column_dimensions['A'].width = 28
        for col in range(ltm_col, yr1_col + n):
            ws.column_dimensions[get_column_letter(col)].width = 14

        self.cell_map['three_statement'] = {
            'revenue_row': rev_row,
            'ebitda_row': ebitda_row,
            'ebit_row': ebit_row,
            'ni_row': ni_row,
            'ta_row': ta_row,
            'debt_row': debt_row,
            'equity_row': eq_row,
            'fcf_row': fcf_row,
            'ltm_col': ltm_col,
            'yr1_col': yr1_col,
        }

        return self

    # ============================================================
    # MODULE: REVERSE DCF
    # ============================================================

    def add_reverse_dcf(self, data: Dict = None) -> 'CFAModulesMixin':
        """Add Reverse DCF / Market-Implied Growth Rate analysis.

        Smart refs: Assumptions, WACC.
        """
        ws = self.wb.create_sheet("Reverse DCF")
        self.sheets_created.append("Reverse DCF")
        a = self.assumptions
        data = data or {}

        self._add_title(ws, f"{self.company_name} - Reverse DCF Analysis", 1, 1)

        # --- Inputs ---
        self._add_section_header(ws, "MARKET DATA & ASSUMPTIONS", 3, 1)
        self._format_header_row(ws, 3, 1, 2)

        share_price = data.get('share_price', 50.0)
        shares_out = data.get('shares_outstanding', 100.0)
        net_debt = data.get('net_debt', a.ltm_ebitda * (a.senior_debt_multiple + a.sub_debt_multiple))
        terminal_growth = data.get('terminal_growth', 0.025)

        inputs = [
            (5, "Current Share Price", share_price, '$#,##0.00'),
            (6, "Shares Outstanding (M)", shares_out, '#,##0.0'),
            (7, "Market Cap", None, '#,##0.0'),  # formula
            (8, "Net Debt", net_debt, '#,##0.0'),
            (9, "Implied Enterprise Value", None, '#,##0.0'),  # formula
            (11, "LTM EBITDA", None, '#,##0.0'),  # from Assumptions
            (12, "LTM FCF Margin", data.get('fcf_margin', 0.10), '0.0%'),
            (13, "Projection Period (years)", data.get('proj_years', 10), '0'),
            (14, "Terminal Growth Rate", terminal_growth, '0.0%'),
        ]

        amap = self.cell_map.get('assumptions', {})
        wacc_map = self.cell_map.get('wacc', {})

        for row, label, value, fmt in inputs:
            ws.cell(row=row, column=1, value=label)
            if row == 7:
                ws.cell(row=row, column=2, value="=B5*B6")
            elif row == 9:
                ws.cell(row=row, column=2, value="=B7+B8")
                ws.cell(row=row, column=2).font = Font(bold=True)
            elif row == 11:
                if amap.get('ltm_ebitda'):
                    ws.cell(row=row, column=2,
                            value=f"='Assumptions'!{amap['ltm_ebitda']}")
                else:
                    ws.cell(row=row, column=2, value=a.ltm_ebitda)
                    self._format_input_cell(ws, row, 2)
            else:
                ws.cell(row=row, column=2, value=value)
                self._format_input_cell(ws, row, 2)
            ws.cell(row=row, column=2).number_format = fmt

        # WACC
        wacc_row = 15
        ws.cell(row=wacc_row, column=1, value="WACC / Discount Rate")
        if wacc_map.get('wacc'):
            ws.cell(row=wacc_row, column=2,
                    value=f"='WACC'!{wacc_map['wacc']}")
        else:
            ws.cell(row=wacc_row, column=2, value=data.get('wacc', 0.10))
            self._format_input_cell(ws, wacc_row, 2)
        ws.cell(row=wacc_row, column=2).number_format = '0.0%'

        # --- Implied Growth Calculation ---
        self._add_section_header(ws, "IMPLIED GROWTH RATE", 17, 1)

        ws.cell(row=18, column=1, value="Implied EV/EBITDA")
        ws.cell(row=18, column=2, value="=IFERROR(B9/B11,0)")
        ws.cell(row=18, column=2).number_format = '0.0x'
        ws.cell(row=18, column=2).font = Font(bold=True)

        # Simplified reverse DCF: g = WACC - FCF/EV (Gordon Growth approximation)
        ws.cell(row=19, column=1, value="LTM FCF Proxy")
        ws.cell(row=19, column=2, value="=B11*B12")
        ws.cell(row=19, column=2).number_format = '#,##0.0'

        ws.cell(row=20, column=1, value="Implied Perpetuity Growth")
        ws.cell(row=20, column=2, value="=IFERROR(B15-B19/B9,0)")
        ws.cell(row=20, column=2).number_format = '0.0%'
        ws.cell(row=20, column=2).font = Font(bold=True)
        ws.cell(row=20, column=2).fill = PatternFill(
            start_color="90EE90", end_color="90EE90", fill_type="solid")

        ws.cell(row=21, column=1, value="Implied Revenue CAGR (5-yr est.)")
        ws.cell(row=21, column=2, value="=IFERROR(B20+B15*0.3,0)")
        ws.cell(row=21, column=2).number_format = '0.0%'

        # --- 2-Way Sensitivity: WACC x Terminal Growth ---
        sens_start = 23
        self._add_section_header(ws, "SENSITIVITY: IMPLIED EV/EBITDA", sens_start, 1)
        sens_start += 1

        wacc_values = data.get('wacc_range', [0.08, 0.09, 0.10, 0.11, 0.12])
        tg_values = data.get('tg_range', [0.015, 0.020, 0.025, 0.030, 0.035])

        # Column headers (terminal growth)
        for j, tg in enumerate(tg_values):
            ws.cell(row=sens_start, column=3 + j, value=tg)
            ws.cell(row=sens_start, column=3 + j).number_format = '0.0%'
            ws.cell(row=sens_start, column=3 + j).font = Font(bold=True)
            ws.cell(row=sens_start, column=3 + j).alignment = Alignment(horizontal='center')
        ws.cell(row=sens_start, column=2, value="WACC \\ Tg")
        ws.cell(row=sens_start, column=2).font = Font(bold=True)
        self._format_header_row(ws, sens_start, 2, 2 + len(tg_values))

        # Row headers (WACC) and formulas
        for i, w in enumerate(wacc_values):
            row = sens_start + 1 + i
            ws.cell(row=row, column=2, value=w)
            ws.cell(row=row, column=2).number_format = '0.0%'
            ws.cell(row=row, column=2).font = Font(bold=True)
            for j, tg in enumerate(tg_values):
                col = 3 + j
                # Implied EV/EBITDA via Gordon Growth: (FCF margin) / (WACC - g)
                ws.cell(row=row, column=col,
                        value=f"=IFERROR(B12/(B{row}2-{get_column_letter(col)}${sens_start}),0)")
                # Simplified: use actual formula with cell refs
                w_ref = f"B{row}"
                tg_ref = f"{get_column_letter(col)}${sens_start}"
                ws.cell(row=row, column=col,
                        value=f"=IFERROR($B$12/({w_ref}-{tg_ref}),0)")
                ws.cell(row=row, column=col).number_format = '0.0x'

        ws.column_dimensions['A'].width = 30
        ws.column_dimensions['B'].width = 16
        for j in range(len(tg_values)):
            ws.column_dimensions[get_column_letter(3 + j)].width = 12

        self.cell_map['reverse_dcf'] = {
            'share_price': 'B5',
            'market_cap': 'B7',
            'implied_ev': 'B9',
            'implied_ev_ebitda': 'B18',
            'implied_growth': 'B20',
            'wacc': f'B{wacc_row}',
        }

        return self

    # ============================================================
    # MODULE: DUPONT ANALYSIS
    # ============================================================

    def add_dupont_analysis(self, data: Dict = None) -> 'CFAModulesMixin':
        """Add DuPont ROE Decomposition (3-component and 5-component).

        Smart refs: Three-Statement or Operating Model for financials.
        """
        ws = self.wb.create_sheet("DuPont Analysis")
        self.sheets_created.append("DuPont Analysis")
        data = data or {}
        n = self.projection_years

        self._add_title(ws, f"{self.company_name} - DuPont ROE Decomposition", 1, 1)

        # Year headers
        yr_hdr = 3
        ws.cell(row=yr_hdr, column=1, value="Component")
        ws.cell(row=yr_hdr, column=2, value="LTM")
        for i in range(n):
            ws.cell(row=yr_hdr, column=3 + i, value=f"Year {i + 1}")
        self._format_header_row(ws, yr_hdr, 1, 2 + n)

        ltm_col = 2
        yr1_col = 3

        # --- Input rows (from data or defaults) ---
        rev_row = 4
        ws.cell(row=rev_row, column=1, value="Revenue")
        ni_row = 5
        ws.cell(row=ni_row, column=1, value="Net Income")
        ta_row = 6
        ws.cell(row=ta_row, column=1, value="Total Assets")
        eq_row = 7
        ws.cell(row=eq_row, column=1, value="Total Equity")
        ebt_row = 8
        ws.cell(row=ebt_row, column=1, value="Earnings Before Tax")
        ebit_row = 9
        ws.cell(row=ebit_row, column=1, value="EBIT")

        ts = self.cell_map.get('three_statement', {})

        # Populate financials from 3-Statement or data/defaults
        defaults = {
            'revenue': [self.assumptions.ltm_revenue] + [self.assumptions.ltm_revenue * (1.07 ** (i + 1)) for i in range(n)],
            'net_income': [self.assumptions.ltm_ebitda * 0.5] + [self.assumptions.ltm_ebitda * 0.5 * (1.05 ** (i + 1)) for i in range(n)],
            'total_assets': [self.assumptions.ltm_revenue * 1.5] * (n + 1),
            'equity': [self.assumptions.ltm_revenue * 0.6] * (n + 1),
            'ebt': [self.assumptions.ltm_ebitda * 0.65] * (n + 1),
            'ebit': [self.assumptions.ltm_ebitda * 0.85] * (n + 1),
        }

        row_map = {
            rev_row: ('revenue_row', 'revenue'),
            ni_row: ('ni_row', 'net_income'),
            ta_row: ('ta_row', 'total_assets'),
            eq_row: ('equity_row', 'equity'),
            ebt_row: ('ebt_row', 'ebt'),
            ebit_row: ('ebit_row', 'ebit'),
        }

        for r, (ts_key, default_key) in row_map.items():
            for col in range(ltm_col, yr1_col + n):
                if ts.get(ts_key) and ts.get('ltm_col'):
                    ts_cl = get_column_letter(col)
                    ws.cell(row=r, column=col,
                            value=f"='3-Statement'!{ts_cl}{ts[ts_key]}")
                else:
                    idx = col - ltm_col
                    vals = data.get(default_key, defaults[default_key])
                    ws.cell(row=r, column=col, value=vals[idx] if idx < len(vals) else vals[-1])
                    self._format_input_cell(ws, r, col)
                ws.cell(row=r, column=col).number_format = '#,##0.0'

        # --- 3-Component DuPont ---
        dp3_start = ebit_row + 2
        self._add_section_header(ws, "3-COMPONENT DUPONT", dp3_start, 1)

        # Net Profit Margin = NI / Revenue
        npm_row = dp3_start + 1
        ws.cell(row=npm_row, column=1, value="Net Profit Margin")
        for col in range(ltm_col, yr1_col + n):
            cl = get_column_letter(col)
            ws.cell(row=npm_row, column=col,
                    value=f"=IFERROR({cl}{ni_row}/{cl}{rev_row},0)")
            ws.cell(row=npm_row, column=col).number_format = '0.0%'

        # Asset Turnover = Revenue / Total Assets
        at_row = npm_row + 1
        ws.cell(row=at_row, column=1, value="Asset Turnover")
        for col in range(ltm_col, yr1_col + n):
            cl = get_column_letter(col)
            ws.cell(row=at_row, column=col,
                    value=f"=IFERROR({cl}{rev_row}/{cl}{ta_row},0)")
            ws.cell(row=at_row, column=col).number_format = '0.00x'

        # Financial Leverage = Total Assets / Equity
        fl_row = at_row + 1
        ws.cell(row=fl_row, column=1, value="Equity Multiplier")
        for col in range(ltm_col, yr1_col + n):
            cl = get_column_letter(col)
            ws.cell(row=fl_row, column=col,
                    value=f"=IFERROR({cl}{ta_row}/{cl}{eq_row},0)")
            ws.cell(row=fl_row, column=col).number_format = '0.00x'

        # ROE = NPM x AT x EM
        roe3_row = fl_row + 1
        ws.cell(row=roe3_row, column=1, value="ROE (3-Component)")
        ws.cell(row=roe3_row, column=1).font = Font(bold=True)
        for col in range(ltm_col, yr1_col + n):
            cl = get_column_letter(col)
            ws.cell(row=roe3_row, column=col,
                    value=f"={cl}{npm_row}*{cl}{at_row}*{cl}{fl_row}")
            ws.cell(row=roe3_row, column=col).number_format = '0.0%'
            ws.cell(row=roe3_row, column=col).font = Font(bold=True)
            ws.cell(row=roe3_row, column=col).border = DOUBLE_BORDER

        # --- 5-Component DuPont ---
        dp5_start = roe3_row + 2
        self._add_section_header(ws, "5-COMPONENT DUPONT", dp5_start, 1)

        # Tax Burden = NI / EBT
        tax_b_row = dp5_start + 1
        ws.cell(row=tax_b_row, column=1, value="Tax Burden (NI/EBT)")
        for col in range(ltm_col, yr1_col + n):
            cl = get_column_letter(col)
            ws.cell(row=tax_b_row, column=col,
                    value=f"=IFERROR({cl}{ni_row}/{cl}{ebt_row},0)")
            ws.cell(row=tax_b_row, column=col).number_format = '0.00'

        # Interest Burden = EBT / EBIT
        int_b_row = tax_b_row + 1
        ws.cell(row=int_b_row, column=1, value="Interest Burden (EBT/EBIT)")
        for col in range(ltm_col, yr1_col + n):
            cl = get_column_letter(col)
            ws.cell(row=int_b_row, column=col,
                    value=f"=IFERROR({cl}{ebt_row}/{cl}{ebit_row},0)")
            ws.cell(row=int_b_row, column=col).number_format = '0.00'

        # Operating Margin = EBIT / Revenue
        opm_row = int_b_row + 1
        ws.cell(row=opm_row, column=1, value="Operating Margin (EBIT/Rev)")
        for col in range(ltm_col, yr1_col + n):
            cl = get_column_letter(col)
            ws.cell(row=opm_row, column=col,
                    value=f"=IFERROR({cl}{ebit_row}/{cl}{rev_row},0)")
            ws.cell(row=opm_row, column=col).number_format = '0.0%'

        # Asset Turnover (same as above)
        at5_row = opm_row + 1
        ws.cell(row=at5_row, column=1, value="Asset Turnover")
        for col in range(ltm_col, yr1_col + n):
            cl = get_column_letter(col)
            ws.cell(row=at5_row, column=col, value=f"={cl}{at_row}")
            ws.cell(row=at5_row, column=col).number_format = '0.00x'

        # Equity Multiplier (same as above)
        em5_row = at5_row + 1
        ws.cell(row=em5_row, column=1, value="Equity Multiplier")
        for col in range(ltm_col, yr1_col + n):
            cl = get_column_letter(col)
            ws.cell(row=em5_row, column=col, value=f"={cl}{fl_row}")
            ws.cell(row=em5_row, column=col).number_format = '0.00x'

        # ROE = Tax Burden x Interest Burden x OPM x AT x EM
        roe5_row = em5_row + 1
        ws.cell(row=roe5_row, column=1, value="ROE (5-Component)")
        ws.cell(row=roe5_row, column=1).font = Font(bold=True)
        for col in range(ltm_col, yr1_col + n):
            cl = get_column_letter(col)
            ws.cell(row=roe5_row, column=col,
                    value=f"={cl}{tax_b_row}*{cl}{int_b_row}*{cl}{opm_row}*{cl}{at5_row}*{cl}{em5_row}")
            ws.cell(row=roe5_row, column=col).number_format = '0.0%'
            ws.cell(row=roe5_row, column=col).font = Font(bold=True)
            ws.cell(row=roe5_row, column=col).border = DOUBLE_BORDER
            ws.cell(row=roe5_row, column=col).fill = PatternFill(
                start_color="90EE90", end_color="90EE90", fill_type="solid")

        # Column widths
        ws.column_dimensions['A'].width = 30
        for col in range(ltm_col, yr1_col + n):
            ws.column_dimensions[get_column_letter(col)].width = 14

        self.cell_map['dupont'] = {
            'roe3_row': roe3_row,
            'roe5_row': roe5_row,
            'npm_row': npm_row,
            'asset_turnover_row': at_row,
            'equity_multiplier_row': fl_row,
            'ltm_col': ltm_col,
        }

        return self

    # ============================================================
    # MODULE: ROIC DECOMPOSITION
    # ============================================================

    def add_roic_decomposition(self, data: Dict = None) -> 'CFAModulesMixin':
        """Add ROIC vs WACC spread and economic profit analysis.

        Smart refs: Three-Statement, WACC, Assumptions.
        """
        ws = self.wb.create_sheet("ROIC Decomposition")
        self.sheets_created.append("ROIC Decomposition")
        a = self.assumptions
        data = data or {}
        n = self.projection_years

        self._add_title(ws, f"{self.company_name} - ROIC Decomposition", 1, 1)

        # Year headers
        yr_hdr = 3
        ws.cell(row=yr_hdr, column=1, value="($M)")
        ws.cell(row=yr_hdr, column=2, value="LTM")
        for i in range(n):
            ws.cell(row=yr_hdr, column=3 + i, value=f"Year {i + 1}")
        self._format_header_row(ws, yr_hdr, 1, 2 + n)

        ltm_col = 2
        yr1_col = 3

        # --- NOPAT Calculation ---
        self._add_section_header(ws, "NOPAT CALCULATION", 4, 1)

        ebit_input_row = 5
        ws.cell(row=ebit_input_row, column=1, value="EBIT")
        ts = self.cell_map.get('three_statement', {})
        for col in range(ltm_col, yr1_col + n):
            if ts.get('ebit_row') and ts.get('ltm_col'):
                ts_cl = get_column_letter(col)
                ws.cell(row=ebit_input_row, column=col,
                        value=f"='3-Statement'!{ts_cl}{ts['ebit_row']}")
            else:
                idx = col - ltm_col
                ebit_default = a.ltm_ebitda * 0.85 * (1.06 ** idx)
                ws.cell(row=ebit_input_row, column=col, value=ebit_default)
                self._format_input_cell(ws, ebit_input_row, col)
            ws.cell(row=ebit_input_row, column=col).number_format = '#,##0.0'

        tax_rate_row = 6
        ws.cell(row=tax_rate_row, column=1, value="Tax Rate")
        ws.cell(row=tax_rate_row, column=2, value=f"='Assumptions'!B24")
        ws.cell(row=tax_rate_row, column=2).number_format = '0.0%'

        nopat_row = 7
        ws.cell(row=nopat_row, column=1, value="NOPAT (EBIT x (1-t))")
        ws.cell(row=nopat_row, column=1).font = Font(bold=True)
        for col in range(ltm_col, yr1_col + n):
            cl = get_column_letter(col)
            ws.cell(row=nopat_row, column=col,
                    value=f"={cl}{ebit_input_row}*(1-$B${tax_rate_row})")
            ws.cell(row=nopat_row, column=col).number_format = '#,##0.0'
            ws.cell(row=nopat_row, column=col).font = Font(bold=True)

        # --- Invested Capital ---
        self._add_section_header(ws, "INVESTED CAPITAL", 9, 1)

        nwc_row = 10
        ws.cell(row=nwc_row, column=1, value="Net Working Capital")
        for col in range(ltm_col, yr1_col + n):
            idx = col - ltm_col
            rev_est = a.ltm_revenue * (1.07 ** idx)
            if ts.get('revenue_row'):
                ts_cl = get_column_letter(col)
                ws.cell(row=nwc_row, column=col,
                        value=f"='3-Statement'!{ts_cl}{ts['revenue_row']}*'Assumptions'!B23")
            else:
                ws.cell(row=nwc_row, column=col, value=rev_est * a.nwc_pct_revenue)
                self._format_input_cell(ws, nwc_row, col)
            ws.cell(row=nwc_row, column=col).number_format = '#,##0.0'

        fa_row = 11
        ws.cell(row=fa_row, column=1, value="Net Fixed Assets")
        for col in range(ltm_col, yr1_col + n):
            idx = col - ltm_col
            rev_est = a.ltm_revenue * (1.07 ** idx)
            ws.cell(row=fa_row, column=col, value=rev_est * 0.30)
            ws.cell(row=fa_row, column=col).number_format = '#,##0.0'
            self._format_input_cell(ws, fa_row, col)

        ic_row = 12
        ws.cell(row=ic_row, column=1, value="Invested Capital")
        ws.cell(row=ic_row, column=1).font = Font(bold=True)
        for col in range(ltm_col, yr1_col + n):
            cl = get_column_letter(col)
            ws.cell(row=ic_row, column=col,
                    value=f"={cl}{nwc_row}+{cl}{fa_row}")
            ws.cell(row=ic_row, column=col).number_format = '#,##0.0'
            ws.cell(row=ic_row, column=col).font = Font(bold=True)
            ws.cell(row=ic_row, column=col).border = BOTTOM_BORDER

        # --- ROIC ---
        self._add_section_header(ws, "RETURN ON INVESTED CAPITAL", 14, 1)

        roic_row = 15
        ws.cell(row=roic_row, column=1, value="ROIC (NOPAT / IC)")
        ws.cell(row=roic_row, column=1).font = Font(bold=True)
        for col in range(ltm_col, yr1_col + n):
            cl = get_column_letter(col)
            ws.cell(row=roic_row, column=col,
                    value=f"=IFERROR({cl}{nopat_row}/{cl}{ic_row},0)")
            ws.cell(row=roic_row, column=col).number_format = '0.0%'
            ws.cell(row=roic_row, column=col).font = Font(bold=True)

        # WACC
        wacc_row = 16
        ws.cell(row=wacc_row, column=1, value="WACC")
        wacc_map = self.cell_map.get('wacc', {})
        if wacc_map.get('wacc'):
            ws.cell(row=wacc_row, column=ltm_col,
                    value=f"='WACC'!{wacc_map['wacc']}")
        else:
            ws.cell(row=wacc_row, column=ltm_col, value=data.get('wacc', 0.10))
            self._format_input_cell(ws, wacc_row, ltm_col)
        ws.cell(row=wacc_row, column=ltm_col).number_format = '0.0%'
        # Copy WACC across years
        for i in range(n):
            col = yr1_col + i
            ws.cell(row=wacc_row, column=col, value=f"=$B${wacc_row}")
            ws.cell(row=wacc_row, column=col).number_format = '0.0%'

        # ROIC - WACC Spread
        spread_row = 17
        ws.cell(row=spread_row, column=1, value="ROIC - WACC Spread")
        ws.cell(row=spread_row, column=1).font = Font(bold=True)
        for col in range(ltm_col, yr1_col + n):
            cl = get_column_letter(col)
            ws.cell(row=spread_row, column=col,
                    value=f"={cl}{roic_row}-{cl}{wacc_row}")
            ws.cell(row=spread_row, column=col).number_format = '0.0%'
            ws.cell(row=spread_row, column=col).font = Font(bold=True)
            ws.cell(row=spread_row, column=col).border = DOUBLE_BORDER

        # --- Economic Profit ---
        self._add_section_header(ws, "ECONOMIC PROFIT (EVA)", 19, 1)

        ep_row = 20
        ws.cell(row=ep_row, column=1, value="Economic Profit ($M)")
        ws.cell(row=ep_row, column=1).font = Font(bold=True)
        for col in range(ltm_col, yr1_col + n):
            cl = get_column_letter(col)
            ws.cell(row=ep_row, column=col,
                    value=f"={cl}{spread_row}*{cl}{ic_row}")
            ws.cell(row=ep_row, column=col).number_format = '#,##0.0'
            ws.cell(row=ep_row, column=col).font = Font(bold=True)
            ws.cell(row=ep_row, column=col).fill = PatternFill(
                start_color="90EE90", end_color="90EE90", fill_type="solid")

        # IC Turnover
        ict_row = 21
        ws.cell(row=ict_row, column=1, value="IC Turnover (Revenue/IC)")
        for col in range(ltm_col, yr1_col + n):
            cl = get_column_letter(col)
            if ts.get('revenue_row'):
                ts_cl = get_column_letter(col)
                ws.cell(row=ict_row, column=col,
                        value=f"=IFERROR('3-Statement'!{ts_cl}{ts['revenue_row']}/{cl}{ic_row},0)")
            else:
                idx = col - ltm_col
                rev_est = a.ltm_revenue * (1.07 ** idx)
                ws.cell(row=ict_row, column=col,
                        value=f"=IFERROR({rev_est}/{cl}{ic_row},0)")
            ws.cell(row=ict_row, column=col).number_format = '0.00x'

        # NOPAT Margin
        npm_row_roic = 22
        ws.cell(row=npm_row_roic, column=1, value="NOPAT Margin")
        for col in range(ltm_col, yr1_col + n):
            cl = get_column_letter(col)
            if ts.get('revenue_row'):
                ts_cl = get_column_letter(col)
                ws.cell(row=npm_row_roic, column=col,
                        value=f"=IFERROR({cl}{nopat_row}/'3-Statement'!{ts_cl}{ts['revenue_row']},0)")
            else:
                idx = col - ltm_col
                rev_est = a.ltm_revenue * (1.07 ** idx)
                ws.cell(row=npm_row_roic, column=col,
                        value=f"=IFERROR({cl}{nopat_row}/{rev_est},0)")
            ws.cell(row=npm_row_roic, column=col).number_format = '0.0%'

        # Column widths
        ws.column_dimensions['A'].width = 30
        for col in range(ltm_col, yr1_col + n):
            ws.column_dimensions[get_column_letter(col)].width = 14

        self.cell_map['roic'] = {
            'nopat_row': nopat_row,
            'ic_row': ic_row,
            'roic_row': roic_row,
            'wacc_row': wacc_row,
            'spread_row': spread_row,
            'ep_row': ep_row,
            'ltm_col': ltm_col,
        }

        return self

    # ============================================================
    # MODULE: DDM VALUATION
    # ============================================================

    def add_ddm_valuation(self, data: Dict = None) -> 'CFAModulesMixin':
        """Add Dividend Discount Model (Gordon Growth + Two-Stage + H-Model).

        Smart refs: Assumptions, WACC.
        """
        ws = self.wb.create_sheet("DDM Valuation")
        self.sheets_created.append("DDM Valuation")
        a = self.assumptions
        data = data or {}

        self._add_title(ws, f"{self.company_name} - Dividend Discount Model", 1, 1)

        # --- Inputs ---
        self._add_section_header(ws, "DDM INPUTS", 3, 1)
        self._format_header_row(ws, 3, 1, 2)

        current_div = data.get('current_dividend', 2.00)
        high_growth = data.get('high_growth', 0.15)
        stable_growth = data.get('stable_growth', 0.03)
        high_growth_years = data.get('high_growth_years', 5)
        payout_ratio = data.get('payout_ratio', 0.40)

        inputs = [
            (5, "Current Dividend per Share", current_div, '$#,##0.00'),
            (6, "Current EPS", data.get('eps', 5.00), '$#,##0.00'),
            (7, "Payout Ratio", payout_ratio, '0.0%'),
            (8, "High-Growth Rate", high_growth, '0.0%'),
            (9, "High-Growth Period (years)", high_growth_years, '0'),
            (10, "Stable Growth Rate", stable_growth, '0.0%'),
        ]

        for row, label, value, fmt in inputs:
            ws.cell(row=row, column=1, value=label)
            ws.cell(row=row, column=2, value=value)
            ws.cell(row=row, column=2).number_format = fmt
            self._format_input_cell(ws, row, 2)

        # Cost of Equity
        ke_row = 11
        ws.cell(row=ke_row, column=1, value="Cost of Equity (Ke)")
        wacc_map = self.cell_map.get('wacc', {})
        if wacc_map.get('cost_of_equity'):
            ws.cell(row=ke_row, column=2,
                    value=f"='WACC'!{wacc_map['cost_of_equity']}")
        else:
            ke = a.risk_free_rate + a.beta * a.equity_risk_premium + a.size_premium
            ws.cell(row=ke_row, column=2, value=ke)
            self._format_input_cell(ws, ke_row, 2)
        ws.cell(row=ke_row, column=2).number_format = '0.0%'

        # ============ GORDON GROWTH MODEL ============
        ggm_start = 13
        self._add_section_header(ws, "1. GORDON GROWTH MODEL (Single-Stage)", ggm_start, 1)

        ws.cell(row=ggm_start + 1, column=1, value="D1 (Next Year Dividend)")
        ws.cell(row=ggm_start + 1, column=2, value=f"=B5*(1+B10)")
        ws.cell(row=ggm_start + 1, column=2).number_format = '$#,##0.00'

        ws.cell(row=ggm_start + 2, column=1, value="GGM Value = D1 / (Ke - g)")
        ws.cell(row=ggm_start + 2, column=2,
                value=f"=IFERROR(B{ggm_start + 1}/(B{ke_row}-B10),0)")
        ws.cell(row=ggm_start + 2, column=2).number_format = '$#,##0.00'
        ws.cell(row=ggm_start + 2, column=2).font = Font(bold=True)
        ws.cell(row=ggm_start + 2, column=2).fill = PatternFill(
            start_color="90EE90", end_color="90EE90", fill_type="solid")
        ggm_val_row = ggm_start + 2

        # ============ TWO-STAGE DDM ============
        ts_start = ggm_start + 4
        self._add_section_header(ws, "2. TWO-STAGE DDM", ts_start, 1)

        # High-growth phase dividends
        ts_start += 1
        ws.cell(row=ts_start, column=1, value="Year")
        ws.cell(row=ts_start, column=2, value="Dividend")
        ws.cell(row=ts_start, column=3, value="PV Factor")
        ws.cell(row=ts_start, column=4, value="PV of Dividend")
        self._format_header_row(ws, ts_start, 1, 4)

        div_rows_start = ts_start + 1
        max_hg = 7  # Max years to display
        for i in range(max_hg):
            row = div_rows_start + i
            yr = i + 1
            ws.cell(row=row, column=1, value=yr)
            # Dividend = D0 * (1+g)^yr
            ws.cell(row=row, column=2,
                    value=f"=B5*(1+B8)^{yr}")
            ws.cell(row=row, column=2).number_format = '$#,##0.00'
            # PV factor = 1/(1+Ke)^yr
            ws.cell(row=row, column=3,
                    value=f"=1/(1+$B${ke_row})^{yr}")
            ws.cell(row=row, column=3).number_format = '0.0000'
            # PV of dividend
            ws.cell(row=row, column=4, value=f"=B{row}*C{row}")
            ws.cell(row=row, column=4).number_format = '$#,##0.00'
        div_rows_end = div_rows_start + max_hg - 1

        # PV of high-growth dividends
        pv_hg_row = div_rows_end + 1
        ws.cell(row=pv_hg_row, column=1, value="PV of High-Growth Divs")
        ws.cell(row=pv_hg_row, column=4,
                value=f"=SUMPRODUCT((ROW(A{div_rows_start}:A{div_rows_end})-ROW(A{div_rows_start})+1<=B9)*1,D{div_rows_start}:D{div_rows_end})")
        ws.cell(row=pv_hg_row, column=4).number_format = '$#,##0.00'
        ws.cell(row=pv_hg_row, column=4).font = Font(bold=True)

        # Terminal value
        tv_row = pv_hg_row + 1
        ws.cell(row=tv_row, column=1, value="Terminal Div (D_n+1)")
        ws.cell(row=tv_row, column=2,
                value=f"=B5*(1+B8)^B9*(1+B10)")
        ws.cell(row=tv_row, column=2).number_format = '$#,##0.00'

        tv_val_row = tv_row + 1
        ws.cell(row=tv_val_row, column=1, value="Terminal Value at Year n")
        ws.cell(row=tv_val_row, column=2,
                value=f"=IFERROR(B{tv_row}/(B{ke_row}-B10),0)")
        ws.cell(row=tv_val_row, column=2).number_format = '$#,##0.00'

        pv_tv_row = tv_val_row + 1
        ws.cell(row=pv_tv_row, column=1, value="PV of Terminal Value")
        ws.cell(row=pv_tv_row, column=2,
                value=f"=B{tv_val_row}/(1+B{ke_row})^B9")
        ws.cell(row=pv_tv_row, column=2).number_format = '$#,##0.00'
        ws.cell(row=pv_tv_row, column=2).font = Font(bold=True)

        two_stage_val_row = pv_tv_row + 1
        ws.cell(row=two_stage_val_row, column=1, value="Two-Stage DDM Value")
        ws.cell(row=two_stage_val_row, column=2,
                value=f"=D{pv_hg_row}+B{pv_tv_row}")
        ws.cell(row=two_stage_val_row, column=2).number_format = '$#,##0.00'
        ws.cell(row=two_stage_val_row, column=2).font = Font(bold=True)
        ws.cell(row=two_stage_val_row, column=2).fill = PatternFill(
            start_color="90EE90", end_color="90EE90", fill_type="solid")

        # ============ H-MODEL ============
        h_start = two_stage_val_row + 2
        self._add_section_header(ws, "3. H-MODEL (Linear Decay)", h_start, 1)

        ws.cell(row=h_start + 1, column=1, value="H (half-life of decay, years)")
        ws.cell(row=h_start + 1, column=2, value=f"=B9/2")
        ws.cell(row=h_start + 1, column=2).number_format = '0.0'

        h_val_row = h_start + 2
        ws.cell(row=h_val_row, column=1, value="H-Model Value")
        # V = D0(1+gL)/(Ke-gL) + D0*H*(gS-gL)/(Ke-gL)
        ws.cell(row=h_val_row, column=2,
                value=f"=IFERROR(B5*(1+B10)/(B{ke_row}-B10)+B5*B{h_start + 1}*(B8-B10)/(B{ke_row}-B10),0)")
        ws.cell(row=h_val_row, column=2).number_format = '$#,##0.00'
        ws.cell(row=h_val_row, column=2).font = Font(bold=True)
        ws.cell(row=h_val_row, column=2).fill = PatternFill(
            start_color="90EE90", end_color="90EE90", fill_type="solid")

        # ============ SUMMARY ============
        sum_start = h_val_row + 2
        self._add_section_header(ws, "VALUATION SUMMARY", sum_start, 1)
        sum_start += 1
        ws.cell(row=sum_start, column=1, value="Model")
        ws.cell(row=sum_start, column=2, value="Intrinsic Value")
        self._format_header_row(ws, sum_start, 1, 2)

        ws.cell(row=sum_start + 1, column=1, value="Gordon Growth")
        ws.cell(row=sum_start + 1, column=2, value=f"=B{ggm_val_row}")
        ws.cell(row=sum_start + 1, column=2).number_format = '$#,##0.00'
        ws.cell(row=sum_start + 2, column=1, value="Two-Stage DDM")
        ws.cell(row=sum_start + 2, column=2, value=f"=B{two_stage_val_row}")
        ws.cell(row=sum_start + 2, column=2).number_format = '$#,##0.00'
        ws.cell(row=sum_start + 3, column=1, value="H-Model")
        ws.cell(row=sum_start + 3, column=2, value=f"=B{h_val_row}")
        ws.cell(row=sum_start + 3, column=2).number_format = '$#,##0.00'

        avg_row = sum_start + 4
        ws.cell(row=avg_row, column=1, value="Average")
        ws.cell(row=avg_row, column=1).font = Font(bold=True)
        ws.cell(row=avg_row, column=2,
                value=f"=AVERAGE(B{sum_start + 1}:B{sum_start + 3})")
        ws.cell(row=avg_row, column=2).number_format = '$#,##0.00'
        ws.cell(row=avg_row, column=2).font = Font(bold=True)
        ws.cell(row=avg_row, column=2).border = DOUBLE_BORDER

        ws.column_dimensions['A'].width = 32
        ws.column_dimensions['B'].width = 18
        ws.column_dimensions['C'].width = 14
        ws.column_dimensions['D'].width = 16

        self.cell_map['ddm'] = {
            'ggm_value': f'B{ggm_val_row}',
            'two_stage_value': f'B{two_stage_val_row}',
            'h_model_value': f'B{h_val_row}',
            'average_value': f'B{avg_row}',
            'cost_of_equity': f'B{ke_row}',
        }

        return self

    # ============================================================
    # MODULE: EARNINGS QUALITY SCORES
    # ============================================================

    def add_earnings_quality(self, data: Dict = None) -> 'CFAModulesMixin':
        """Add Earnings Quality dashboard: Beneish M-Score, Altman Z-Score,
        Piotroski F-Score.

        Smart refs: Three-Statement, Assumptions.
        """
        ws = self.wb.create_sheet("Earnings Quality")
        self.sheets_created.append("Earnings Quality")
        data = data or {}

        self._add_title(ws, f"{self.company_name} - Earnings Quality Scores", 1, 1)

        # ============ INPUT DATA ============
        self._add_section_header(ws, "FINANCIAL DATA", 3, 1)
        ws.cell(row=3, column=2, value="Current")
        ws.cell(row=3, column=3, value="Prior Year")
        self._format_header_row(ws, 3, 1, 3)

        # Default financial data (override via data dict)
        fin = data.get('financials', {})
        fin_rows = [
            (4, "Revenue", fin.get('revenue', 100.0), fin.get('revenue_py', 90.0)),
            (5, "COGS", fin.get('cogs', 60.0), fin.get('cogs_py', 55.0)),
            (6, "Gross Profit", None, None),  # formula
            (7, "SG&A", fin.get('sga', 15.0), fin.get('sga_py', 13.0)),
            (8, "Depreciation", fin.get('da', 5.0), fin.get('da_py', 4.5)),
            (9, "Net Income", fin.get('ni', 12.0), fin.get('ni_py', 10.0)),
            (10, "Cash from Ops", fin.get('cfo', 18.0), fin.get('cfo_py', 15.0)),
            (11, "Total Assets", fin.get('ta', 150.0), fin.get('ta_py', 140.0)),
            (12, "Current Assets", fin.get('ca', 50.0), fin.get('ca_py', 45.0)),
            (13, "Current Liabilities", fin.get('cl', 30.0), fin.get('cl_py', 28.0)),
            (14, "Total Debt", fin.get('debt', 40.0), fin.get('debt_py', 42.0)),
            (15, "Total Equity", fin.get('equity', 80.0), fin.get('equity_py', 70.0)),
            (16, "Accounts Receivable", fin.get('ar', 15.0), fin.get('ar_py', 12.0)),
            (17, "PP&E (net)", fin.get('ppe', 45.0), fin.get('ppe_py', 43.0)),
            (18, "EBIT", fin.get('ebit', 15.0), fin.get('ebit_py', 13.0)),
            (19, "Interest Expense", fin.get('interest', 3.5), fin.get('interest_py', 3.8)),
            (20, "Market Cap", fin.get('market_cap', 200.0), None),
            (21, "Working Capital", None, None),  # formula
            (22, "Retained Earnings", fin.get('re', 50.0), fin.get('re_py', 40.0)),
        ]

        for row, label, curr, prior in fin_rows:
            ws.cell(row=row, column=1, value=label)
            if row == 6:  # Gross Profit = Revenue - COGS
                ws.cell(row=row, column=2, value="=B4-B5")
                ws.cell(row=row, column=3, value="=C4-C5")
            elif row == 21:  # Working Capital = CA - CL
                ws.cell(row=row, column=2, value="=B12-B13")
                ws.cell(row=row, column=3, value="=C12-C13")
            else:
                if curr is not None:
                    ws.cell(row=row, column=2, value=curr)
                    self._format_input_cell(ws, row, 2)
                if prior is not None:
                    ws.cell(row=row, column=3, value=prior)
                    self._format_input_cell(ws, row, 3)
            ws.cell(row=row, column=2).number_format = '#,##0.0'
            if prior is not None or row in (6, 21):
                ws.cell(row=row, column=3).number_format = '#,##0.0'

        # ============ BENEISH M-SCORE ============
        ben_start = 24
        self._add_section_header(ws, "BENEISH M-SCORE (Earnings Manipulation)", ben_start, 1)

        components = [
            (ben_start + 1, "DSRI (Days Sales Receivable Index)",
             "=IFERROR((B16/B4)/(C16/C4),1)"),
            (ben_start + 2, "GMI (Gross Margin Index)",
             "=IFERROR((C6/C4)/(B6/B4),1)"),
            (ben_start + 3, "AQI (Asset Quality Index)",
             "=IFERROR((1-B12/B11-B17/B11)/(1-C12/C11-C17/C11),1)"),
            (ben_start + 4, "SGI (Sales Growth Index)",
             "=IFERROR(B4/C4,1)"),
            (ben_start + 5, "DEPI (Depreciation Index)",
             "=IFERROR((C8/(C8+C17))/(B8/(B8+B17)),1)"),
            (ben_start + 6, "SGAI (SG&A Index)",
             "=IFERROR((B7/B4)/(C7/C4),1)"),
            (ben_start + 7, "LVGI (Leverage Index)",
             "=IFERROR((B14+B13)/B11/((C14+C13)/C11),1)"),
            (ben_start + 8, "TATA (Total Accruals / TA)",
             "=IFERROR((B9-B10)/B11,0)"),
        ]

        for row, label, formula in components:
            ws.cell(row=row, column=1, value=label)
            ws.cell(row=row, column=2, value=formula)
            ws.cell(row=row, column=2).number_format = '0.000'

        # M-Score formula
        m_row = ben_start + 9
        ws.cell(row=m_row, column=1, value="M-Score")
        ws.cell(row=m_row, column=1).font = Font(bold=True)
        dsri = f"B{ben_start + 1}"
        gmi = f"B{ben_start + 2}"
        aqi = f"B{ben_start + 3}"
        sgi = f"B{ben_start + 4}"
        depi = f"B{ben_start + 5}"
        sgai = f"B{ben_start + 6}"
        lvgi = f"B{ben_start + 7}"
        tata = f"B{ben_start + 8}"
        ws.cell(row=m_row, column=2,
                value=f"=-4.84+0.92*{dsri}+0.528*{gmi}+0.404*{aqi}+0.892*{sgi}+0.115*{depi}-0.172*{sgai}+4.679*{tata}-0.327*{lvgi}")
        ws.cell(row=m_row, column=2).number_format = '0.00'
        ws.cell(row=m_row, column=2).font = Font(bold=True)

        # Interpretation
        ws.cell(row=m_row + 1, column=1, value="Assessment")
        ws.cell(row=m_row + 1, column=2,
                value=f'=IF(B{m_row}<-1.78,"Unlikely Manipulation","Possible Manipulation")')
        ws.cell(row=m_row + 1, column=2).font = Font(bold=True)

        # ============ ALTMAN Z-SCORE ============
        z_start = m_row + 3
        self._add_section_header(ws, "ALTMAN Z-SCORE (Bankruptcy Risk)", z_start, 1)

        z_components = [
            (z_start + 1, "X1 (WC / TA)", "=IFERROR(B21/B11,0)"),
            (z_start + 2, "X2 (RE / TA)", "=IFERROR(B22/B11,0)"),
            (z_start + 3, "X3 (EBIT / TA)", "=IFERROR(B18/B11,0)"),
            (z_start + 4, "X4 (Market Cap / TL)", f"=IFERROR(B20/(B11-B15),0)"),
            (z_start + 5, "X5 (Revenue / TA)", "=IFERROR(B4/B11,0)"),
        ]

        for row, label, formula in z_components:
            ws.cell(row=row, column=1, value=label)
            ws.cell(row=row, column=2, value=formula)
            ws.cell(row=row, column=2).number_format = '0.000'

        z_row = z_start + 6
        ws.cell(row=z_row, column=1, value="Z-Score")
        ws.cell(row=z_row, column=1).font = Font(bold=True)
        x1 = f"B{z_start + 1}"
        x2 = f"B{z_start + 2}"
        x3 = f"B{z_start + 3}"
        x4 = f"B{z_start + 4}"
        x5 = f"B{z_start + 5}"
        ws.cell(row=z_row, column=2,
                value=f"=1.2*{x1}+1.4*{x2}+3.3*{x3}+0.6*{x4}+{x5}")
        ws.cell(row=z_row, column=2).number_format = '0.00'
        ws.cell(row=z_row, column=2).font = Font(bold=True)

        ws.cell(row=z_row + 1, column=1, value="Assessment")
        ws.cell(row=z_row + 1, column=2,
                value=f'=IF(B{z_row}>2.99,"Safe Zone",IF(B{z_row}>1.81,"Grey Zone","Distress Zone"))')
        ws.cell(row=z_row + 1, column=2).font = Font(bold=True)

        # ============ PIOTROSKI F-SCORE ============
        f_start = z_row + 3
        self._add_section_header(ws, "PIOTROSKI F-SCORE (Financial Strength)", f_start, 1)

        f_tests = [
            (f_start + 1, "1. ROA > 0", "=IF(B9/B11>0,1,0)"),
            (f_start + 2, "2. CFO > 0", "=IF(B10>0,1,0)"),
            (f_start + 3, "3. ROA Improving", "=IF(B9/B11>C9/C11,1,0)"),
            (f_start + 4, "4. CFO > NI (Accrual Quality)", "=IF(B10>B9,1,0)"),
            (f_start + 5, "5. Leverage Declining", "=IF(B14/B11<C14/C11,1,0)"),
            (f_start + 6, "6. Current Ratio Improving", "=IF(B12/B13>C12/C13,1,0)"),
            (f_start + 7, "7. No Dilution (shares constant)", "=1"),
            (f_start + 8, "8. Gross Margin Improving", "=IF(B6/B4>C6/C4,1,0)"),
            (f_start + 9, "9. Asset Turnover Improving", "=IF(B4/B11>C4/C11,1,0)"),
        ]

        for row, label, formula in f_tests:
            ws.cell(row=row, column=1, value=label)
            ws.cell(row=row, column=2, value=formula)
            ws.cell(row=row, column=2).number_format = '0'
            ws.cell(row=row, column=2).alignment = Alignment(horizontal='center')

        f_row = f_start + 10
        ws.cell(row=f_row, column=1, value="F-Score (0-9)")
        ws.cell(row=f_row, column=1).font = Font(bold=True)
        ws.cell(row=f_row, column=2,
                value=f"=SUM(B{f_start + 1}:B{f_start + 9})")
        ws.cell(row=f_row, column=2).number_format = '0'
        ws.cell(row=f_row, column=2).font = Font(bold=True, size=14)

        ws.cell(row=f_row + 1, column=1, value="Assessment")
        ws.cell(row=f_row + 1, column=2,
                value=f'=IF(B{f_row}>=7,"Strong",IF(B{f_row}>=4,"Neutral","Weak"))')
        ws.cell(row=f_row + 1, column=2).font = Font(bold=True)

        # ============ COMPOSITE DASHBOARD ============
        dash_start = f_row + 3
        self._add_section_header(ws, "COMPOSITE DASHBOARD", dash_start, 1)
        dash_start += 1

        ws.cell(row=dash_start, column=1, value="Score")
        ws.cell(row=dash_start, column=2, value="Value")
        ws.cell(row=dash_start, column=3, value="Assessment")
        self._format_header_row(ws, dash_start, 1, 3)

        ws.cell(row=dash_start + 1, column=1, value="Beneish M-Score")
        ws.cell(row=dash_start + 1, column=2, value=f"=B{m_row}")
        ws.cell(row=dash_start + 1, column=2).number_format = '0.00'
        ws.cell(row=dash_start + 1, column=3, value=f"=B{m_row + 1}")

        ws.cell(row=dash_start + 2, column=1, value="Altman Z-Score")
        ws.cell(row=dash_start + 2, column=2, value=f"=B{z_row}")
        ws.cell(row=dash_start + 2, column=2).number_format = '0.00'
        ws.cell(row=dash_start + 2, column=3, value=f"=B{z_row + 1}")

        ws.cell(row=dash_start + 3, column=1, value="Piotroski F-Score")
        ws.cell(row=dash_start + 3, column=2, value=f"=B{f_row}")
        ws.cell(row=dash_start + 3, column=2).number_format = '0'
        ws.cell(row=dash_start + 3, column=3, value=f"=B{f_row + 1}")

        for r in range(dash_start + 1, dash_start + 4):
            for c in range(1, 4):
                ws.cell(row=r, column=c).font = Font(bold=True)

        # --- Conditional Formatting (RAG) ---
        # M-Score: below -1.78 is good (green), above is red — lower is better
        from openpyxl.formatting.rule import FormulaRule
        ws.conditional_formatting.add(
            f'C{dash_start + 1}',
            FormulaRule(formula=[f'C{dash_start + 1}="Unlikely Manipulation"'],
                        fill=PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')))
        ws.conditional_formatting.add(
            f'C{dash_start + 1}',
            FormulaRule(formula=[f'C{dash_start + 1}="Possible Manipulation"'],
                        fill=PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')))
        # Z-Score: Safe Zone=green, Grey=amber, Distress=red
        ws.conditional_formatting.add(
            f'C{dash_start + 2}',
            FormulaRule(formula=[f'C{dash_start + 2}="Safe Zone"'],
                        fill=PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')))
        ws.conditional_formatting.add(
            f'C{dash_start + 2}',
            FormulaRule(formula=[f'C{dash_start + 2}="Grey Zone"'],
                        fill=PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')))
        ws.conditional_formatting.add(
            f'C{dash_start + 2}',
            FormulaRule(formula=[f'C{dash_start + 2}="Distress Zone"'],
                        fill=PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')))
        # F-Score: Strong=green, Neutral=amber, Weak=red
        ws.conditional_formatting.add(
            f'C{dash_start + 3}',
            FormulaRule(formula=[f'C{dash_start + 3}="Strong"'],
                        fill=PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')))
        ws.conditional_formatting.add(
            f'C{dash_start + 3}',
            FormulaRule(formula=[f'C{dash_start + 3}="Neutral"'],
                        fill=PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')))
        ws.conditional_formatting.add(
            f'C{dash_start + 3}',
            FormulaRule(formula=[f'C{dash_start + 3}="Weak"'],
                        fill=PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')))

        ws.column_dimensions['A'].width = 35
        ws.column_dimensions['B'].width = 16
        ws.column_dimensions['C'].width = 22

        self.cell_map['earnings_quality'] = {
            'm_score': f'B{m_row}',
            'z_score': f'B{z_row}',
            'f_score': f'B{f_row}',
            'm_assessment': f'B{m_row + 1}',
            'z_assessment': f'B{z_row + 1}',
            'f_assessment': f'B{f_row + 1}',
        }

        return self

    # ============================================================
    # MODULE: TORNADO SENSITIVITY
    # ============================================================

    def add_tornado_sensitivity(self, data: Dict = None) -> 'CFAModulesMixin':
        """Add single-variable tornado sensitivity analysis.

        Tests each assumption ±delta around base case, shows IRR impact
        with data bars for visual ranking.
        Smart refs: Assumptions, Returns Analysis.
        """
        ws = self.wb.create_sheet("Tornado Sensitivity")
        self.sheets_created.append("Tornado Sensitivity")
        a = self.assumptions
        data = data or {}

        self._add_title(ws, f"{self.company_name} - Tornado Sensitivity", 1, 1)

        # --- Base Case ---
        self._add_section_header(ws, "BASE CASE RETURNS", 3, 1)
        self._format_header_row(ws, 3, 1, 2)

        ret_map = self.cell_map.get('returns', {})
        base_irr = data.get('base_irr', 0.20)
        base_moic = data.get('base_moic', 2.5)

        ws.cell(row=4, column=1, value="Base Case IRR")
        if ret_map.get('irr'):
            ws.cell(row=4, column=2,
                    value=f"='Returns Analysis'!{ret_map['irr']}")
        else:
            ws.cell(row=4, column=2, value=base_irr)
            self._format_input_cell(ws, 4, 2)
        ws.cell(row=4, column=2).number_format = '0.0%'

        ws.cell(row=5, column=1, value="Base Case MOIC")
        if ret_map.get('moic'):
            ws.cell(row=5, column=2,
                    value=f"='Returns Analysis'!{ret_map['moic']}")
        else:
            ws.cell(row=5, column=2, value=base_moic)
            self._format_input_cell(ws, 5, 2)
        ws.cell(row=5, column=2).number_format = '0.00"x"'

        # --- Variable Sensitivity Table ---
        self._add_section_header(ws, "IRR SENSITIVITY BY VARIABLE", 7, 1)

        headers = ["Variable", "Base", "Low (-10%)", "High (+10%)",
                   "Low IRR", "Base IRR", "High IRR", "Range (bps)"]
        hdr_row = 8
        for i, h in enumerate(headers):
            ws.cell(row=hdr_row, column=1 + i, value=h)
        self._format_header_row(ws, hdr_row, 1, len(headers))

        delta = data.get('delta', 0.10)
        variables = data.get('variables', [
            {'name': 'Entry Multiple', 'base': a.entry_multiple, 'fmt': '0.0"x"',
             'lo': base_irr + 0.04, 'hi': base_irr - 0.03},
            {'name': 'Exit Multiple', 'base': a.exit_multiple, 'fmt': '0.0"x"',
             'lo': base_irr - 0.03, 'hi': base_irr + 0.04},
            {'name': 'Revenue Growth', 'base': a.revenue_growth[0], 'fmt': '0.0%',
             'lo': base_irr - 0.02, 'hi': base_irr + 0.02},
            {'name': 'EBITDA Margin', 'base': a.ebitda_margin[0], 'fmt': '0.0%',
             'lo': base_irr - 0.025, 'hi': base_irr + 0.025},
            {'name': 'Senior Leverage', 'base': a.senior_debt_multiple, 'fmt': '0.0"x"',
             'lo': base_irr - 0.015, 'hi': base_irr + 0.02},
            {'name': 'Interest Rate', 'base': a.senior_interest_rate, 'fmt': '0.0%',
             'lo': base_irr + 0.01, 'hi': base_irr - 0.015},
            {'name': 'CapEx % Revenue', 'base': a.capex_pct_revenue, 'fmt': '0.0%',
             'lo': base_irr + 0.008, 'hi': base_irr - 0.008},
        ])

        var_start = hdr_row + 1
        for i, v in enumerate(variables):
            row = var_start + i
            ws.cell(row=row, column=1, value=v['name'])
            ws.cell(row=row, column=1).font = Font(bold=True)
            ws.cell(row=row, column=2, value=v['base'])
            ws.cell(row=row, column=2).number_format = v['fmt']
            ws.cell(row=row, column=3, value=v['base'] * (1 - delta))
            ws.cell(row=row, column=3).number_format = v['fmt']
            self._format_input_cell(ws, row, 3)
            ws.cell(row=row, column=4, value=v['base'] * (1 + delta))
            ws.cell(row=row, column=4).number_format = v['fmt']
            self._format_input_cell(ws, row, 4)
            ws.cell(row=row, column=5, value=v['lo'])
            ws.cell(row=row, column=5).number_format = '0.0%'
            self._format_input_cell(ws, row, 5)
            ws.cell(row=row, column=6, value="=$B$4")
            ws.cell(row=row, column=6).number_format = '0.0%'
            ws.cell(row=row, column=7, value=v['hi'])
            ws.cell(row=row, column=7).number_format = '0.0%'
            self._format_input_cell(ws, row, 7)
            # Range in basis points
            ws.cell(row=row, column=8, value=f"=ABS(G{row}-E{row})*10000")
            ws.cell(row=row, column=8).number_format = '#,##0'
            ws.cell(row=row, column=8).font = Font(bold=True)
        var_end = var_start + len(variables) - 1

        # Data bars on range column
        from openpyxl.formatting.rule import DataBarRule
        ws.conditional_formatting.add(
            f"H{var_start}:H{var_end}",
            DataBarRule(start_type='min', end_type='max',
                       color='638EC6', showValue=True))

        # --- Key Takeaways ---
        tk_row = var_end + 2
        self._add_section_header(ws, "KEY TAKEAWAYS", tk_row, 1)
        ws.cell(row=tk_row + 1, column=1, value="Most Sensitive Variable")
        ws.cell(row=tk_row + 1, column=2,
                value=f"=INDEX(A{var_start}:A{var_end},MATCH(MAX(H{var_start}:H{var_end}),H{var_start}:H{var_end},0))")
        ws.cell(row=tk_row + 1, column=2).font = Font(bold=True)
        ws.cell(row=tk_row + 2, column=1, value="Least Sensitive Variable")
        ws.cell(row=tk_row + 2, column=2,
                value=f"=INDEX(A{var_start}:A{var_end},MATCH(MIN(H{var_start}:H{var_end}),H{var_start}:H{var_end},0))")
        ws.cell(row=tk_row + 3, column=1, value="Avg IRR Range (bps)")
        ws.cell(row=tk_row + 3, column=2,
                value=f"=AVERAGE(H{var_start}:H{var_end})")
        ws.cell(row=tk_row + 3, column=2).number_format = '#,##0'

        ws.column_dimensions['A'].width = 22
        for c in ['B', 'C', 'D', 'E', 'F', 'G', 'H']:
            ws.column_dimensions[c].width = 14

        self.cell_map['tornado_sensitivity'] = {
            'var_start': var_start,
            'var_end': var_end,
            'base_irr': 'B4',
            'base_moic': 'B5',
            'range_col': 8,
        }

        return self

    # ============================================================
    # MODULE: FOOTBALL FIELD
    # ============================================================

    def add_football_field(self, data: Dict = None) -> 'CFAModulesMixin':
        """Add Football Field valuation range comparison.

        Shows low / mid / high for each valuation methodology side by side.
        Smart refs: Trading Comps, Transaction Comps, DCF, Reverse DCF, DDM.
        """
        ws = self.wb.create_sheet("Football Field")
        self.sheets_created.append("Football Field")
        a = self.assumptions
        data = data or {}

        self._add_title(ws, f"{self.company_name} - Valuation Football Field", 1, 1)

        # --- Valuation Methods Table ---
        self._add_section_header(ws, "VALUATION RANGE BY METHODOLOGY", 3, 1)
        headers = ["Methodology", "Low", "Mid", "High", "Weight"]
        hdr_row = 4
        for i, h in enumerate(headers):
            ws.cell(row=hdr_row, column=1 + i, value=h)
        self._format_header_row(ws, hdr_row, 1, len(headers))

        tc_map = self.cell_map.get('trading_comps', {})
        txn_map = self.cell_map.get('transaction_comps', {})
        dcf_map = self.cell_map.get('dcf', {})
        rdcf_map = self.cell_map.get('reverse_dcf', {})
        ddm_map = self.cell_map.get('ddm', {})

        # Default EV estimate
        base_ev = a.ltm_ebitda * a.entry_multiple

        methods = data.get('methods', [
            {'name': 'Trading Comps (EV/EBITDA)', 'low': base_ev * 0.85,
             'mid': base_ev, 'high': base_ev * 1.15, 'weight': 0.25},
            {'name': 'Precedent Transactions', 'low': base_ev * 0.90,
             'mid': base_ev * 1.05, 'high': base_ev * 1.20, 'weight': 0.25},
            {'name': 'DCF Valuation', 'low': base_ev * 0.80,
             'mid': base_ev * 0.95, 'high': base_ev * 1.10, 'weight': 0.30},
            {'name': '52-Week Trading Range', 'low': base_ev * 0.75,
             'mid': base_ev * 0.90, 'high': base_ev * 1.05, 'weight': 0.10},
            {'name': 'LBO Analysis (Floor)', 'low': base_ev * 0.70,
             'mid': base_ev * 0.85, 'high': base_ev * 1.00, 'weight': 0.10},
        ])

        method_start = hdr_row + 1
        for i, m in enumerate(methods):
            row = method_start + i
            ws.cell(row=row, column=1, value=m['name'])
            ws.cell(row=row, column=1).font = Font(bold=True)
            ws.cell(row=row, column=2, value=m['low'])
            ws.cell(row=row, column=2).number_format = '#,##0.0'
            self._format_input_cell(ws, row, 2)
            ws.cell(row=row, column=3, value=m['mid'])
            ws.cell(row=row, column=3).number_format = '#,##0.0'
            self._format_input_cell(ws, row, 3)
            ws.cell(row=row, column=4, value=m['high'])
            ws.cell(row=row, column=4).number_format = '#,##0.0'
            self._format_input_cell(ws, row, 4)
            ws.cell(row=row, column=5, value=m['weight'])
            ws.cell(row=row, column=5).number_format = '0%'
            self._format_input_cell(ws, row, 5)
        method_end = method_start + len(methods) - 1

        # --- Range bars (data bars on Mid column) ---
        from openpyxl.formatting.rule import DataBarRule
        ws.conditional_formatting.add(
            f"C{method_start}:C{method_end}",
            DataBarRule(start_type='min', end_type='max',
                       color='4472C4', showValue=True))

        # --- Weighted Valuation ---
        wv_row = method_end + 2
        self._add_section_header(ws, "WEIGHTED VALUATION", wv_row, 1)
        wv_row += 1

        ws.cell(row=wv_row, column=1, value="Weight Check")
        ws.cell(row=wv_row, column=2,
                value=f"=SUM(E{method_start}:E{method_end})")
        ws.cell(row=wv_row, column=2).number_format = '0%'
        wv_row += 1

        ws.cell(row=wv_row, column=1, value="Weighted Low")
        ws.cell(row=wv_row, column=1).font = Font(bold=True)
        ws.cell(row=wv_row, column=2,
                value=f"=SUMPRODUCT(B{method_start}:B{method_end},E{method_start}:E{method_end})")
        ws.cell(row=wv_row, column=2).number_format = '#,##0.0'
        ws.cell(row=wv_row, column=2).font = Font(bold=True)
        wt_low_row = wv_row

        wv_row += 1
        ws.cell(row=wv_row, column=1, value="Weighted Mid")
        ws.cell(row=wv_row, column=1).font = Font(bold=True)
        ws.cell(row=wv_row, column=2,
                value=f"=SUMPRODUCT(C{method_start}:C{method_end},E{method_start}:E{method_end})")
        ws.cell(row=wv_row, column=2).number_format = '#,##0.0'
        ws.cell(row=wv_row, column=2).font = Font(bold=True)
        ws.cell(row=wv_row, column=2).fill = PatternFill(
            start_color="90EE90", end_color="90EE90", fill_type="solid")
        wt_mid_row = wv_row

        wv_row += 1
        ws.cell(row=wv_row, column=1, value="Weighted High")
        ws.cell(row=wv_row, column=1).font = Font(bold=True)
        ws.cell(row=wv_row, column=2,
                value=f"=SUMPRODUCT(D{method_start}:D{method_end},E{method_start}:E{method_end})")
        ws.cell(row=wv_row, column=2).number_format = '#,##0.0'
        ws.cell(row=wv_row, column=2).font = Font(bold=True)
        wt_high_row = wv_row

        # --- Implied Metrics ---
        im_row = wv_row + 2
        self._add_section_header(ws, "IMPLIED METRICS", im_row, 1)
        im_row += 1
        ws.cell(row=im_row, column=1, value="Implied EV/EBITDA (Mid)")
        ws.cell(row=im_row, column=2,
                value=f"=IFERROR(B{wt_mid_row}/{a.ltm_ebitda},0)")
        ws.cell(row=im_row, column=2).number_format = '0.0"x"'
        ws.cell(row=im_row, column=2).font = Font(bold=True)

        im_row += 1
        net_debt = a.ltm_ebitda * (a.senior_debt_multiple + a.sub_debt_multiple)
        ws.cell(row=im_row, column=1, value="Net Debt")
        ws.cell(row=im_row, column=2, value=net_debt)
        ws.cell(row=im_row, column=2).number_format = '#,##0.0'
        self._format_input_cell(ws, im_row, 2)
        net_debt_row = im_row

        im_row += 1
        ws.cell(row=im_row, column=1, value="Implied Equity Value (Mid)")
        ws.cell(row=im_row, column=2,
                value=f"=B{wt_mid_row}-B{net_debt_row}")
        ws.cell(row=im_row, column=2).number_format = '#,##0.0'
        ws.cell(row=im_row, column=2).font = Font(bold=True)
        ws.cell(row=im_row, column=2).border = DOUBLE_BORDER

        ws.column_dimensions['A'].width = 30
        for c in ['B', 'C', 'D', 'E']:
            ws.column_dimensions[c].width = 16

        self.cell_map['football_field'] = {
            'method_start': method_start,
            'method_end': method_end,
            'weighted_low': f'B{wt_low_row}',
            'weighted_mid': f'B{wt_mid_row}',
            'weighted_high': f'B{wt_high_row}',
        }

        return self

    # ============================================================
    # MODULE: SOTP VALUATION
    # ============================================================

    def add_sotp_valuation(self, data: Dict = None) -> 'CFAModulesMixin':
        """Add Sum-of-the-Parts valuation by business segment.

        Individual segment EV → sum → less net debt → equity value.
        Smart refs: Assumptions, Trading Comps.
        """
        ws = self.wb.create_sheet("SOTP Valuation")
        self.sheets_created.append("SOTP Valuation")
        a = self.assumptions
        data = data or {}

        self._add_title(ws, f"{self.company_name} - Sum-of-the-Parts Valuation", 1, 1)

        # --- Segment Breakdown ---
        self._add_section_header(ws, "SEGMENT ANALYSIS", 3, 1)
        headers = ["Segment", "Revenue ($M)", "EBITDA ($M)", "EBITDA Margin",
                   "Comp Multiple", "Implied EV ($M)", "% of Total"]
        hdr_row = 4
        for i, h in enumerate(headers):
            ws.cell(row=hdr_row, column=1 + i, value=h)
        self._format_header_row(ws, hdr_row, 1, len(headers))

        segments = data.get('segments', [
            {'name': 'Core Business', 'revenue': a.ltm_revenue * 0.55,
             'ebitda': a.ltm_ebitda * 0.60, 'multiple': a.entry_multiple + 0.5},
            {'name': 'Growth Segment', 'revenue': a.ltm_revenue * 0.25,
             'ebitda': a.ltm_ebitda * 0.25, 'multiple': a.entry_multiple + 2.0},
            {'name': 'Mature / Cash Cow', 'revenue': a.ltm_revenue * 0.15,
             'ebitda': a.ltm_ebitda * 0.12, 'multiple': a.entry_multiple - 1.5},
            {'name': 'Other / Corporate', 'revenue': a.ltm_revenue * 0.05,
             'ebitda': a.ltm_ebitda * 0.03, 'multiple': a.entry_multiple - 2.0},
        ])

        seg_start = hdr_row + 1
        for i, seg in enumerate(segments):
            row = seg_start + i
            ws.cell(row=row, column=1, value=seg['name'])
            ws.cell(row=row, column=1).font = Font(bold=True)
            ws.cell(row=row, column=2, value=seg['revenue'])
            ws.cell(row=row, column=2).number_format = '#,##0.0'
            self._format_input_cell(ws, row, 2)
            ws.cell(row=row, column=3, value=seg['ebitda'])
            ws.cell(row=row, column=3).number_format = '#,##0.0'
            self._format_input_cell(ws, row, 3)
            # EBITDA Margin = D/C
            ws.cell(row=row, column=4, value=f"=IFERROR(C{row}/B{row},0)")
            ws.cell(row=row, column=4).number_format = '0.0%'
            ws.cell(row=row, column=5, value=seg['multiple'])
            ws.cell(row=row, column=5).number_format = '0.0"x"'
            self._format_input_cell(ws, row, 5)
            # Implied EV = EBITDA × Multiple
            ws.cell(row=row, column=6, value=f"=C{row}*E{row}")
            ws.cell(row=row, column=6).number_format = '#,##0.0'
            ws.cell(row=row, column=6).font = Font(bold=True)
        seg_end = seg_start + len(segments) - 1

        # % of Total (formula after totals computed)
        total_ev_row = seg_end + 1
        for i in range(len(segments)):
            row = seg_start + i
            ws.cell(row=row, column=7,
                    value=f"=IFERROR(F{row}/F{total_ev_row},0)")
            ws.cell(row=row, column=7).number_format = '0.0%'

        # --- Totals ---
        ws.cell(row=total_ev_row, column=1, value="Total")
        ws.cell(row=total_ev_row, column=1).font = Font(bold=True)
        ws.cell(row=total_ev_row, column=2,
                value=f"=SUM(B{seg_start}:B{seg_end})")
        ws.cell(row=total_ev_row, column=2).number_format = '#,##0.0'
        ws.cell(row=total_ev_row, column=2).font = Font(bold=True)
        ws.cell(row=total_ev_row, column=3,
                value=f"=SUM(C{seg_start}:C{seg_end})")
        ws.cell(row=total_ev_row, column=3).number_format = '#,##0.0'
        ws.cell(row=total_ev_row, column=3).font = Font(bold=True)
        ws.cell(row=total_ev_row, column=4,
                value=f"=IFERROR(C{total_ev_row}/B{total_ev_row},0)")
        ws.cell(row=total_ev_row, column=4).number_format = '0.0%'
        ws.cell(row=total_ev_row, column=4).font = Font(bold=True)
        # Blended multiple
        ws.cell(row=total_ev_row, column=5,
                value=f"=IFERROR(F{total_ev_row}/C{total_ev_row},0)")
        ws.cell(row=total_ev_row, column=5).number_format = '0.0"x"'
        ws.cell(row=total_ev_row, column=5).font = Font(bold=True)
        ws.cell(row=total_ev_row, column=6,
                value=f"=SUM(F{seg_start}:F{seg_end})")
        ws.cell(row=total_ev_row, column=6).number_format = '#,##0.0'
        ws.cell(row=total_ev_row, column=6).font = Font(bold=True)
        ws.cell(row=total_ev_row, column=6).border = DOUBLE_BORDER
        ws.cell(row=total_ev_row, column=7, value=1.0)
        ws.cell(row=total_ev_row, column=7).number_format = '0.0%'
        ws.cell(row=total_ev_row, column=7).font = Font(bold=True)

        # --- EV to Equity Bridge ---
        br_row = total_ev_row + 2
        self._add_section_header(ws, "EV TO EQUITY BRIDGE", br_row, 1)
        br_row += 1

        ws.cell(row=br_row, column=1, value="Total SOTP Enterprise Value")
        ws.cell(row=br_row, column=2, value=f"=F{total_ev_row}")
        ws.cell(row=br_row, column=2).number_format = '#,##0.0'
        ws.cell(row=br_row, column=2).font = Font(bold=True)
        sotp_ev_row = br_row

        br_row += 1
        ws.cell(row=br_row, column=1, value="Less: Total Debt")
        net_debt_val = a.ltm_ebitda * (a.senior_debt_multiple + a.sub_debt_multiple)
        ws.cell(row=br_row, column=2, value=-net_debt_val)
        ws.cell(row=br_row, column=2).number_format = '#,##0.0'
        self._format_input_cell(ws, br_row, 2)
        debt_bridge_row = br_row

        br_row += 1
        ws.cell(row=br_row, column=1, value="Plus: Cash & Equivalents")
        cash_val = data.get('cash', 10.0)
        ws.cell(row=br_row, column=2, value=cash_val)
        ws.cell(row=br_row, column=2).number_format = '#,##0.0'
        self._format_input_cell(ws, br_row, 2)
        cash_bridge_row = br_row

        br_row += 1
        ws.cell(row=br_row, column=1, value="Equity Value")
        ws.cell(row=br_row, column=1).font = Font(bold=True)
        ws.cell(row=br_row, column=2,
                value=f"=B{sotp_ev_row}+B{debt_bridge_row}+B{cash_bridge_row}")
        ws.cell(row=br_row, column=2).number_format = '#,##0.0'
        ws.cell(row=br_row, column=2).font = Font(bold=True)
        ws.cell(row=br_row, column=2).border = BOTTOM_BORDER
        equity_row = br_row

        br_row += 1
        shares = data.get('shares_outstanding', 100.0)
        ws.cell(row=br_row, column=1, value="Shares Outstanding (M)")
        ws.cell(row=br_row, column=2, value=shares)
        ws.cell(row=br_row, column=2).number_format = '#,##0.0'
        self._format_input_cell(ws, br_row, 2)
        shares_row = br_row

        br_row += 1
        ws.cell(row=br_row, column=1, value="Implied Price per Share")
        ws.cell(row=br_row, column=1).font = Font(bold=True)
        ws.cell(row=br_row, column=2,
                value=f"=IFERROR(B{equity_row}/B{shares_row},0)")
        ws.cell(row=br_row, column=2).number_format = '$#,##0.00'
        ws.cell(row=br_row, column=2).font = Font(bold=True, size=14)
        ws.cell(row=br_row, column=2).fill = PatternFill(
            start_color="90EE90", end_color="90EE90", fill_type="solid")
        ws.cell(row=br_row, column=2).border = DOUBLE_BORDER

        # --- Sensitivity: Multiple ± ---
        sens_row = br_row + 2
        self._add_section_header(ws, "SENSITIVITY: BLENDED MULTIPLE", sens_row, 1)
        sens_row += 1
        ws.cell(row=sens_row, column=1, value="Multiple Adj")
        adjustments = [-2.0, -1.0, 0, 1.0, 2.0]
        for j, adj in enumerate(adjustments):
            ws.cell(row=sens_row, column=2 + j, value=adj)
            ws.cell(row=sens_row, column=2 + j).number_format = '+0.0;-0.0;0.0'
        self._format_header_row(ws, sens_row, 1, 1 + len(adjustments))
        sens_row += 1
        ws.cell(row=sens_row, column=1, value="Implied EV")
        ws.cell(row=sens_row, column=1).font = Font(bold=True)
        for j, adj in enumerate(adjustments):
            # Total EBITDA × (blended multiple + adj)
            ws.cell(row=sens_row, column=2 + j,
                    value=f"=C{total_ev_row}*(E{total_ev_row}+{adj})")
            ws.cell(row=sens_row, column=2 + j).number_format = '#,##0.0'
            ws.cell(row=sens_row, column=2 + j).font = Font(bold=True)

        ws.column_dimensions['A'].width = 28
        for c in ['B', 'C', 'D', 'E', 'F', 'G']:
            ws.column_dimensions[c].width = 16

        self.cell_map['sotp'] = {
            'seg_start': seg_start,
            'seg_end': seg_end,
            'total_ev': f'F{total_ev_row}',
            'equity_value': f'B{equity_row}',
            'sotp_ev_row': sotp_ev_row,
        }

        return self

    # ============================================================
    # MODULE: ENHANCED WACC
    # ============================================================

    def add_enhanced_wacc(self, data: Dict = None) -> 'CFAModulesMixin':
        """Add Enhanced WACC with multi-method beta and cost of debt.

        Features: 3 beta methods (Raw, Blume Adjusted, Hamada),
        3 cost of debt methods (Synthetic, YTM, Book Rate),
        market-based capital structure weights.
        Smart refs: Assumptions.
        """
        ws = self.wb.create_sheet("Enhanced WACC")
        self.sheets_created.append("Enhanced WACC")
        a = self.assumptions
        data = data or {}

        self._add_title(ws, f"{self.company_name} - Enhanced WACC Analysis", 1, 1)

        # --- Market Data Inputs ---
        self._add_section_header(ws, "MARKET DATA", 3, 1)
        self._format_header_row(ws, 3, 1, 2)

        rf = data.get('risk_free_rate', a.risk_free_rate)
        erp = data.get('equity_risk_premium', a.equity_risk_premium)
        raw_beta = data.get('raw_beta', a.beta)
        size_prem = data.get('size_premium', a.size_premium)
        country_prem = data.get('country_premium', 0.0)
        market_cap = data.get('market_cap', 500.0)
        total_debt = data.get('total_debt',
                              a.ltm_ebitda * (a.senior_debt_multiple + a.sub_debt_multiple))
        tax_rate = a.tax_rate

        inputs = [
            (5, "Risk-Free Rate", rf, '0.00%'),
            (6, "Equity Risk Premium", erp, '0.00%'),
            (7, "Raw Beta", raw_beta, '0.00'),
            (8, "Size Premium", size_prem, '0.00%'),
            (9, "Country Risk Premium", country_prem, '0.00%'),
            (10, "Tax Rate", tax_rate, '0.0%'),
            (12, "Market Cap ($M)", market_cap, '#,##0.0'),
            (13, "Total Debt ($M)", total_debt, '#,##0.0'),
            (14, "Cash ($M)", data.get('cash', 10.0), '#,##0.0'),
        ]

        for row, label, value, fmt in inputs:
            ws.cell(row=row, column=1, value=label)
            ws.cell(row=row, column=2, value=value)
            ws.cell(row=row, column=2).number_format = fmt
            self._format_input_cell(ws, row, 2)

        # Net Debt
        ws.cell(row=15, column=1, value="Net Debt ($M)")
        ws.cell(row=15, column=2, value="=B13-B14")
        ws.cell(row=15, column=2).number_format = '#,##0.0'
        ws.cell(row=15, column=2).font = Font(bold=True)

        # ============ BETA METHODS ============
        beta_start = 17
        self._add_section_header(ws, "BETA ESTIMATION (3 Methods)", beta_start, 1)

        # Method 1: Raw Beta
        ws.cell(row=beta_start + 1, column=1, value="1. Raw (Regression) Beta")
        ws.cell(row=beta_start + 1, column=2, value="=B7")
        ws.cell(row=beta_start + 1, column=2).number_format = '0.00'
        raw_beta_row = beta_start + 1

        # Method 2: Blume Adjusted
        ws.cell(row=beta_start + 2, column=1, value="2. Blume Adjusted Beta")
        ws.cell(row=beta_start + 2, column=2, value="=2/3*B7+1/3*1")
        ws.cell(row=beta_start + 2, column=2).number_format = '0.00'
        blume_row = beta_start + 2

        # Method 3: Hamada (unlever/relever)
        ws.cell(row=beta_start + 3, column=1, value="3a. Unlevered Beta")
        ws.cell(row=beta_start + 3, column=2,
                value="=IFERROR(B7/(1+(1-B10)*B13/B12),B7)")
        ws.cell(row=beta_start + 3, column=2).number_format = '0.00'
        unlev_row = beta_start + 3

        ws.cell(row=beta_start + 4, column=1, value="   Target D/E Ratio")
        target_de = data.get('target_de', 0.40)
        ws.cell(row=beta_start + 4, column=2, value=target_de)
        ws.cell(row=beta_start + 4, column=2).number_format = '0.00'
        self._format_input_cell(ws, beta_start + 4, 2)

        ws.cell(row=beta_start + 5, column=1, value="3b. Relevered Beta")
        ws.cell(row=beta_start + 5, column=2,
                value=f"=B{unlev_row}*(1+(1-B10)*B{beta_start + 4})")
        ws.cell(row=beta_start + 5, column=2).number_format = '0.00'
        relev_row = beta_start + 5

        # Selected Beta
        ws.cell(row=beta_start + 6, column=1, value="Selected Beta")
        ws.cell(row=beta_start + 6, column=1).font = Font(bold=True)
        ws.cell(row=beta_start + 6, column=2,
                value=f"=AVERAGE(B{raw_beta_row},B{blume_row},B{relev_row})")
        ws.cell(row=beta_start + 6, column=2).number_format = '0.00'
        ws.cell(row=beta_start + 6, column=2).font = Font(bold=True)
        ws.cell(row=beta_start + 6, column=2).fill = PatternFill(
            start_color="90EE90", end_color="90EE90", fill_type="solid")
        sel_beta_row = beta_start + 6

        # ============ COST OF EQUITY ============
        ke_start = sel_beta_row + 2
        self._add_section_header(ws, "COST OF EQUITY (CAPM)", ke_start, 1)

        ws.cell(row=ke_start + 1, column=1, value="Ke = Rf + Beta × ERP + SP + CRP")
        ws.cell(row=ke_start + 1, column=2,
                value=f"=B5+B{sel_beta_row}*B6+B8+B9")
        ws.cell(row=ke_start + 1, column=2).number_format = '0.00%'
        ws.cell(row=ke_start + 1, column=2).font = Font(bold=True)
        ws.cell(row=ke_start + 1, column=2).fill = PatternFill(
            start_color="90EE90", end_color="90EE90", fill_type="solid")
        ke_row = ke_start + 1

        # ============ COST OF DEBT (3 Methods) ============
        kd_start = ke_row + 2
        self._add_section_header(ws, "COST OF DEBT (3 Methods)", kd_start, 1)

        # Method 1: Synthetic Rating
        ws.cell(row=kd_start + 1, column=1, value="1. Synthetic Rating Spread")
        syn_spread = data.get('synthetic_spread', 0.025)
        ws.cell(row=kd_start + 1, column=2, value=syn_spread)
        ws.cell(row=kd_start + 1, column=2).number_format = '0.00%'
        self._format_input_cell(ws, kd_start + 1, 2)

        ws.cell(row=kd_start + 2, column=1, value="   Kd (Synthetic) = Rf + Spread")
        ws.cell(row=kd_start + 2, column=2,
                value=f"=B5+B{kd_start + 1}")
        ws.cell(row=kd_start + 2, column=2).number_format = '0.00%'
        kd_syn_row = kd_start + 2

        # Method 2: YTM
        ws.cell(row=kd_start + 3, column=1, value="2. Yield-to-Maturity")
        ytm = data.get('ytm', a.senior_interest_rate + 0.005)
        ws.cell(row=kd_start + 3, column=2, value=ytm)
        ws.cell(row=kd_start + 3, column=2).number_format = '0.00%'
        self._format_input_cell(ws, kd_start + 3, 2)
        kd_ytm_row = kd_start + 3

        # Method 3: Book Rate
        ws.cell(row=kd_start + 4, column=1, value="3. Book (Effective) Rate")
        ws.cell(row=kd_start + 4, column=2, value=a.senior_interest_rate)
        ws.cell(row=kd_start + 4, column=2).number_format = '0.00%'
        self._format_input_cell(ws, kd_start + 4, 2)
        kd_book_row = kd_start + 4

        # Selected Kd
        ws.cell(row=kd_start + 5, column=1, value="Selected Pre-Tax Kd")
        ws.cell(row=kd_start + 5, column=1).font = Font(bold=True)
        ws.cell(row=kd_start + 5, column=2,
                value=f"=AVERAGE(B{kd_syn_row},B{kd_ytm_row},B{kd_book_row})")
        ws.cell(row=kd_start + 5, column=2).number_format = '0.00%'
        ws.cell(row=kd_start + 5, column=2).font = Font(bold=True)
        sel_kd_row = kd_start + 5

        ws.cell(row=kd_start + 6, column=1, value="After-Tax Kd")
        ws.cell(row=kd_start + 6, column=2,
                value=f"=B{sel_kd_row}*(1-B10)")
        ws.cell(row=kd_start + 6, column=2).number_format = '0.00%'
        ws.cell(row=kd_start + 6, column=2).font = Font(bold=True)
        ws.cell(row=kd_start + 6, column=2).fill = PatternFill(
            start_color="90EE90", end_color="90EE90", fill_type="solid")
        at_kd_row = kd_start + 6

        # ============ CAPITAL STRUCTURE WEIGHTS ============
        cs_start = at_kd_row + 2
        self._add_section_header(ws, "CAPITAL STRUCTURE (Market Weights)", cs_start, 1)

        ws.cell(row=cs_start + 1, column=1, value="Total Capital")
        ws.cell(row=cs_start + 1, column=2, value="=B12+B13")
        ws.cell(row=cs_start + 1, column=2).number_format = '#,##0.0'
        tot_cap_row = cs_start + 1

        ws.cell(row=cs_start + 2, column=1, value="Equity Weight (We)")
        ws.cell(row=cs_start + 2, column=2,
                value=f"=IFERROR(B12/B{tot_cap_row},0.5)")
        ws.cell(row=cs_start + 2, column=2).number_format = '0.0%'
        ws.cell(row=cs_start + 2, column=2).font = Font(bold=True)
        we_row = cs_start + 2

        ws.cell(row=cs_start + 3, column=1, value="Debt Weight (Wd)")
        ws.cell(row=cs_start + 3, column=2,
                value=f"=IFERROR(B13/B{tot_cap_row},0.5)")
        ws.cell(row=cs_start + 3, column=2).number_format = '0.0%'
        ws.cell(row=cs_start + 3, column=2).font = Font(bold=True)
        wd_row = cs_start + 3

        # ============ WACC CALCULATION ============
        wacc_start = wd_row + 2
        self._add_section_header(ws, "WACC CALCULATION", wacc_start, 1)

        ws.cell(row=wacc_start + 1, column=1, value="WACC = We × Ke + Wd × Kd(1-t)")
        ws.cell(row=wacc_start + 1, column=2,
                value=f"=B{we_row}*B{ke_row}+B{wd_row}*B{at_kd_row}")
        ws.cell(row=wacc_start + 1, column=2).number_format = '0.00%'
        ws.cell(row=wacc_start + 1, column=2).font = Font(bold=True, size=14)
        ws.cell(row=wacc_start + 1, column=2).fill = PatternFill(
            start_color="90EE90", end_color="90EE90", fill_type="solid")
        ws.cell(row=wacc_start + 1, column=2).border = DOUBLE_BORDER
        wacc_result_row = wacc_start + 1

        # --- WACC Range ---
        wr_start = wacc_result_row + 2
        self._add_section_header(ws, "WACC SENSITIVITY RANGE", wr_start, 1)
        wr_start += 1
        ws.cell(row=wr_start, column=1, value="Scenario")
        ws.cell(row=wr_start, column=2, value="WACC")
        self._format_header_row(ws, wr_start, 1, 2)

        ws.cell(row=wr_start + 1, column=1, value="Low (All Low Inputs)")
        ws.cell(row=wr_start + 1, column=2,
                value=f"=B{wacc_result_row}-0.015")
        ws.cell(row=wr_start + 1, column=2).number_format = '0.00%'
        ws.cell(row=wr_start + 2, column=1, value="Base Case")
        ws.cell(row=wr_start + 2, column=2, value=f"=B{wacc_result_row}")
        ws.cell(row=wr_start + 2, column=2).number_format = '0.00%'
        ws.cell(row=wr_start + 2, column=2).font = Font(bold=True)
        ws.cell(row=wr_start + 3, column=1, value="High (All High Inputs)")
        ws.cell(row=wr_start + 3, column=2,
                value=f"=B{wacc_result_row}+0.015")
        ws.cell(row=wr_start + 3, column=2).number_format = '0.00%'

        ws.column_dimensions['A'].width = 35
        ws.column_dimensions['B'].width = 18

        self.cell_map['enhanced_wacc'] = {
            'wacc': f'B{wacc_result_row}',
            'cost_of_equity': f'B{ke_row}',
            'after_tax_kd': f'B{at_kd_row}',
            'equity_weight': f'B{we_row}',
            'debt_weight': f'B{wd_row}',
            'selected_beta': f'B{sel_beta_row}',
        }

        return self

    # ============================================================
    # MODULE: GEOGRAPHIC TERMINAL GROWTH
    # ============================================================

    def add_geographic_terminal_growth(self, data: Dict = None) -> 'CFAModulesMixin':
        """Add GDP-weighted terminal growth rate by geography.

        Blends regional GDP growth rates by revenue exposure to derive
        a defensible terminal growth rate.
        Smart refs: Assumptions.
        """
        ws = self.wb.create_sheet("Geographic Term Growth")
        self.sheets_created.append("Geographic Term Growth")
        a = self.assumptions
        data = data or {}

        self._add_title(ws, f"{self.company_name} - Geographic Terminal Growth", 1, 1)

        # --- Revenue by Geography ---
        self._add_section_header(ws, "REVENUE EXPOSURE BY REGION", 3, 1)
        headers = ["Region", "Revenue ($M)", "Revenue %", "Real GDP Growth",
                   "Inflation", "Nominal GDP Growth", "Weighted Growth"]
        hdr_row = 4
        for i, h in enumerate(headers):
            ws.cell(row=hdr_row, column=1 + i, value=h)
        self._format_header_row(ws, hdr_row, 1, len(headers))

        regions = data.get('regions', [
            {'name': 'North America', 'revenue': a.ltm_revenue * 0.50,
             'gdp_real': 0.020, 'inflation': 0.025},
            {'name': 'Europe', 'revenue': a.ltm_revenue * 0.25,
             'gdp_real': 0.015, 'inflation': 0.020},
            {'name': 'Asia-Pacific', 'revenue': a.ltm_revenue * 0.15,
             'gdp_real': 0.040, 'inflation': 0.030},
            {'name': 'Latin America', 'revenue': a.ltm_revenue * 0.05,
             'gdp_real': 0.025, 'inflation': 0.050},
            {'name': 'Rest of World', 'revenue': a.ltm_revenue * 0.05,
             'gdp_real': 0.030, 'inflation': 0.035},
        ])

        reg_start = hdr_row + 1
        for i, r in enumerate(regions):
            row = reg_start + i
            ws.cell(row=row, column=1, value=r['name'])
            ws.cell(row=row, column=1).font = Font(bold=True)
            ws.cell(row=row, column=2, value=r['revenue'])
            ws.cell(row=row, column=2).number_format = '#,##0.0'
            self._format_input_cell(ws, row, 2)
            # Revenue % (computed after total row)
            ws.cell(row=row, column=4, value=r['gdp_real'])
            ws.cell(row=row, column=4).number_format = '0.0%'
            self._format_input_cell(ws, row, 4)
            ws.cell(row=row, column=5, value=r['inflation'])
            ws.cell(row=row, column=5).number_format = '0.0%'
            self._format_input_cell(ws, row, 5)
            # Nominal GDP = (1+real)*(1+inflation)-1
            ws.cell(row=row, column=6,
                    value=f"=(1+D{row})*(1+E{row})-1")
            ws.cell(row=row, column=6).number_format = '0.0%'
            # Weighted Growth = revenue% × nominal GDP
            ws.cell(row=row, column=7,
                    value=f"=C{row}*F{row}")
            ws.cell(row=row, column=7).number_format = '0.00%'
        reg_end = reg_start + len(regions) - 1

        # Total row
        total_row = reg_end + 1
        ws.cell(row=total_row, column=1, value="Total")
        ws.cell(row=total_row, column=1).font = Font(bold=True)
        ws.cell(row=total_row, column=2,
                value=f"=SUM(B{reg_start}:B{reg_end})")
        ws.cell(row=total_row, column=2).number_format = '#,##0.0'
        ws.cell(row=total_row, column=2).font = Font(bold=True)
        ws.cell(row=total_row, column=3, value=1.0)
        ws.cell(row=total_row, column=3).number_format = '0.0%'
        ws.cell(row=total_row, column=3).font = Font(bold=True)

        # Fill in Revenue % (needs total)
        for i in range(len(regions)):
            row = reg_start + i
            ws.cell(row=row, column=3,
                    value=f"=IFERROR(B{row}/B{total_row},0)")
            ws.cell(row=row, column=3).number_format = '0.0%'

        # --- GDP-Weighted Terminal Growth ---
        tg_row = total_row + 2
        self._add_section_header(ws, "TERMINAL GROWTH RATE DERIVATION", tg_row, 1)
        tg_row += 1

        ws.cell(row=tg_row, column=1, value="GDP-Weighted Nominal Growth")
        ws.cell(row=tg_row, column=2,
                value=f"=SUM(G{reg_start}:G{reg_end})")
        ws.cell(row=tg_row, column=2).number_format = '0.00%'
        ws.cell(row=tg_row, column=2).font = Font(bold=True)
        gdp_wt_row = tg_row

        tg_row += 1
        ws.cell(row=tg_row, column=1, value="Market Share Adj (conservative)")
        ms_adj = data.get('market_share_adj', -0.005)
        ws.cell(row=tg_row, column=2, value=ms_adj)
        ws.cell(row=tg_row, column=2).number_format = '0.00%'
        self._format_input_cell(ws, tg_row, 2)
        ms_adj_row = tg_row

        tg_row += 1
        ws.cell(row=tg_row, column=1, value="Company-Specific Adj")
        co_adj = data.get('company_adj', 0.0)
        ws.cell(row=tg_row, column=2, value=co_adj)
        ws.cell(row=tg_row, column=2).number_format = '0.00%'
        self._format_input_cell(ws, tg_row, 2)
        co_adj_row = tg_row

        tg_row += 1
        ws.cell(row=tg_row, column=1, value="Terminal Growth Rate")
        ws.cell(row=tg_row, column=1).font = Font(bold=True)
        ws.cell(row=tg_row, column=2,
                value=f"=B{gdp_wt_row}+B{ms_adj_row}+B{co_adj_row}")
        ws.cell(row=tg_row, column=2).number_format = '0.00%'
        ws.cell(row=tg_row, column=2).font = Font(bold=True, size=14)
        ws.cell(row=tg_row, column=2).fill = PatternFill(
            start_color="90EE90", end_color="90EE90", fill_type="solid")
        ws.cell(row=tg_row, column=2).border = DOUBLE_BORDER
        term_growth_row = tg_row

        # --- Reasonability Check ---
        rc_row = tg_row + 2
        self._add_section_header(ws, "REASONABILITY CHECK", rc_row, 1)
        rc_row += 1
        ws.cell(row=rc_row, column=1, value="US 10-Year Treasury (proxy)")
        ws.cell(row=rc_row, column=2, value=a.risk_free_rate)
        ws.cell(row=rc_row, column=2).number_format = '0.00%'
        self._format_input_cell(ws, rc_row, 2)
        rf_check_row = rc_row

        rc_row += 1
        ws.cell(row=rc_row, column=1, value="Terminal Growth vs Risk-Free")
        ws.cell(row=rc_row, column=2,
                value=f"=B{term_growth_row}-B{rf_check_row}")
        ws.cell(row=rc_row, column=2).number_format = '0.00%'
        ws.cell(row=rc_row, column=2).font = Font(bold=True)

        rc_row += 1
        ws.cell(row=rc_row, column=1, value="Reasonability")
        ws.cell(row=rc_row, column=2,
                value=f'=IF(B{term_growth_row}<=B{rf_check_row},"PASS - Below Rf","CAUTION - Above Rf")')
        ws.cell(row=rc_row, column=2).font = Font(bold=True)

        ws.column_dimensions['A'].width = 32
        for c in ['B', 'C', 'D', 'E', 'F', 'G']:
            ws.column_dimensions[c].width = 16

        self.cell_map['geographic_terminal_growth'] = {
            'gdp_weighted_growth': f'B{gdp_wt_row}',
            'terminal_growth': f'B{term_growth_row}',
            'reg_start': reg_start,
            'reg_end': reg_end,
        }

        return self

    # ============================================================
    # MODULE: MULTI-STAGE DCF
    # ============================================================

    def add_multi_stage_dcf(self, data: Dict = None) -> 'CFAModulesMixin':
        """Add 3-stage DCF: high-growth → transition → terminal value.

        Stage 1: Explicit FCF projection (high growth).
        Stage 2: Linear growth decay to terminal rate.
        Stage 3: Gordon Growth terminal value.
        Smart refs: Assumptions, WACC, Operating Model, Geographic Terminal Growth.
        """
        ws = self.wb.create_sheet("Multi-Stage DCF")
        self.sheets_created.append("Multi-Stage DCF")
        a = self.assumptions
        data = data or {}

        self._add_title(ws, f"{self.company_name} - Multi-Stage DCF Valuation", 1, 1)

        # --- Inputs ---
        self._add_section_header(ws, "DCF ASSUMPTIONS", 3, 1)
        self._format_header_row(ws, 3, 1, 2)

        high_growth = data.get('high_growth', a.revenue_growth[0])
        stage1_years = data.get('stage1_years', 5)
        stage2_years = data.get('stage2_years', 5)
        fcf_margin = data.get('fcf_margin', 0.10)

        amap = self.cell_map.get('assumptions', {})
        wacc_map = self.cell_map.get('wacc', {})
        ewacc_map = self.cell_map.get('enhanced_wacc', {})
        geo_map = self.cell_map.get('geographic_terminal_growth', {})

        inputs = [
            (5, "LTM Revenue ($M)", None, '#,##0.0'),   # from Assumptions
            (6, "LTM FCF Margin", fcf_margin, '0.0%'),
            (7, "High-Growth Rate (Stage 1)", high_growth, '0.0%'),
            (8, "Stage 1 Period (years)", stage1_years, '0'),
            (9, "Stage 2 Transition (years)", stage2_years, '0'),
            (10, "Terminal Growth Rate", None, '0.0%'),  # from Geo or input
        ]

        for row, label, value, fmt in inputs:
            ws.cell(row=row, column=1, value=label)
            if row == 5:
                if amap.get('ltm_revenue'):
                    ws.cell(row=row, column=2,
                            value=f"='Assumptions'!{amap['ltm_revenue']}")
                else:
                    ws.cell(row=row, column=2, value=a.ltm_revenue)
                    self._format_input_cell(ws, row, 2)
            elif row == 10:
                if geo_map.get('terminal_growth'):
                    ws.cell(row=row, column=2,
                            value=f"='Geographic Term Growth'!{geo_map['terminal_growth']}")
                else:
                    ws.cell(row=row, column=2,
                            value=data.get('terminal_growth', 0.025))
                    self._format_input_cell(ws, row, 2)
            else:
                ws.cell(row=row, column=2, value=value)
                self._format_input_cell(ws, row, 2)
            ws.cell(row=row, column=2).number_format = fmt

        # WACC
        wacc_row = 11
        ws.cell(row=wacc_row, column=1, value="WACC / Discount Rate")
        if ewacc_map.get('wacc'):
            ws.cell(row=wacc_row, column=2,
                    value=f"='Enhanced WACC'!{ewacc_map['wacc']}")
        elif wacc_map.get('wacc'):
            ws.cell(row=wacc_row, column=2,
                    value=f"='WACC'!{wacc_map['wacc']}")
        else:
            ws.cell(row=wacc_row, column=2, value=data.get('wacc', 0.10))
            self._format_input_cell(ws, wacc_row, 2)
        ws.cell(row=wacc_row, column=2).number_format = '0.00%'

        # ============ STAGE 1: HIGH GROWTH ============
        s1_start = 13
        self._add_section_header(ws, "STAGE 1: HIGH GROWTH", s1_start, 1)

        # Year headers
        s1_hdr = s1_start + 1
        ws.cell(row=s1_hdr, column=1, value="")
        for i in range(stage1_years):
            ws.cell(row=s1_hdr, column=2 + i, value=f"Year {i + 1}")
        self._format_header_row(ws, s1_hdr, 1, 1 + stage1_years)

        # Growth Rate
        gr_row = s1_hdr + 1
        ws.cell(row=gr_row, column=1, value="Growth Rate")
        for i in range(stage1_years):
            ws.cell(row=gr_row, column=2 + i, value="=$B$7")
            ws.cell(row=gr_row, column=2 + i).number_format = '0.0%'

        # Revenue
        rev_row = gr_row + 1
        ws.cell(row=rev_row, column=1, value="Revenue")
        for i in range(stage1_years):
            if i == 0:
                ws.cell(row=rev_row, column=2, value=f"=B5*(1+B{gr_row})")
            else:
                prev = get_column_letter(1 + i)
                ws.cell(row=rev_row, column=2 + i,
                        value=f"={prev}{rev_row}*(1+{get_column_letter(2 + i)}{gr_row})")
            ws.cell(row=rev_row, column=2 + i).number_format = '#,##0.0'

        # FCF
        fcf_row = rev_row + 1
        ws.cell(row=fcf_row, column=1, value="Free Cash Flow")
        ws.cell(row=fcf_row, column=1).font = Font(bold=True)
        for i in range(stage1_years):
            cl = get_column_letter(2 + i)
            ws.cell(row=fcf_row, column=2 + i,
                    value=f"={cl}{rev_row}*$B$6")
            ws.cell(row=fcf_row, column=2 + i).number_format = '#,##0.0'
            ws.cell(row=fcf_row, column=2 + i).font = Font(bold=True)

        # Discount Factor
        df_row = fcf_row + 1
        ws.cell(row=df_row, column=1, value="Discount Factor")
        for i in range(stage1_years):
            yr = i + 1
            ws.cell(row=df_row, column=2 + i,
                    value=f"=1/(1+$B${wacc_row})^{yr}")
            ws.cell(row=df_row, column=2 + i).number_format = '0.0000'

        # PV of FCF
        pv_row = df_row + 1
        ws.cell(row=pv_row, column=1, value="PV of FCF")
        for i in range(stage1_years):
            cl = get_column_letter(2 + i)
            ws.cell(row=pv_row, column=2 + i,
                    value=f"={cl}{fcf_row}*{cl}{df_row}")
            ws.cell(row=pv_row, column=2 + i).number_format = '#,##0.0'

        # Stage 1 PV
        s1_pv_row = pv_row + 1
        ws.cell(row=s1_pv_row, column=1, value="PV of Stage 1 FCFs")
        ws.cell(row=s1_pv_row, column=1).font = Font(bold=True)
        last_s1_col = get_column_letter(1 + stage1_years)
        ws.cell(row=s1_pv_row, column=2,
                value=f"=SUM(B{pv_row}:{last_s1_col}{pv_row})")
        ws.cell(row=s1_pv_row, column=2).number_format = '#,##0.0'
        ws.cell(row=s1_pv_row, column=2).font = Font(bold=True)
        ws.cell(row=s1_pv_row, column=2).border = BOTTOM_BORDER

        # ============ STAGE 2: TRANSITION ============
        s2_start = s1_pv_row + 2
        self._add_section_header(ws, "STAGE 2: TRANSITION (Linear Decay)", s2_start, 1)

        s2_hdr = s2_start + 1
        ws.cell(row=s2_hdr, column=1, value="")
        for i in range(stage2_years):
            ws.cell(row=s2_hdr, column=2 + i,
                    value=f"Year {stage1_years + i + 1}")
        self._format_header_row(ws, s2_hdr, 1, 1 + stage2_years)

        # Growth (linear decay from high to terminal)
        s2_gr_row = s2_hdr + 1
        ws.cell(row=s2_gr_row, column=1, value="Growth Rate")
        for i in range(stage2_years):
            # g_t = g_high - (g_high - g_terminal) × (i+1) / stage2_years
            frac = (i + 1) / stage2_years
            ws.cell(row=s2_gr_row, column=2 + i,
                    value=f"=$B$7-($B$7-$B$10)*{frac:.4f}")
            ws.cell(row=s2_gr_row, column=2 + i).number_format = '0.0%'

        # Revenue
        s2_rev_row = s2_gr_row + 1
        ws.cell(row=s2_rev_row, column=1, value="Revenue")
        for i in range(stage2_years):
            cl = get_column_letter(2 + i)
            if i == 0:
                # Continues from last Stage 1 revenue
                ws.cell(row=s2_rev_row, column=2,
                        value=f"={last_s1_col}{rev_row}*(1+B{s2_gr_row})")
            else:
                prev = get_column_letter(1 + i)
                ws.cell(row=s2_rev_row, column=2 + i,
                        value=f"={prev}{s2_rev_row}*(1+{cl}{s2_gr_row})")
            ws.cell(row=s2_rev_row, column=2 + i).number_format = '#,##0.0'

        # FCF
        s2_fcf_row = s2_rev_row + 1
        ws.cell(row=s2_fcf_row, column=1, value="Free Cash Flow")
        ws.cell(row=s2_fcf_row, column=1).font = Font(bold=True)
        for i in range(stage2_years):
            cl = get_column_letter(2 + i)
            ws.cell(row=s2_fcf_row, column=2 + i,
                    value=f"={cl}{s2_rev_row}*$B$6")
            ws.cell(row=s2_fcf_row, column=2 + i).number_format = '#,##0.0'
            ws.cell(row=s2_fcf_row, column=2 + i).font = Font(bold=True)

        # Discount Factor (continuing from Stage 1)
        s2_df_row = s2_fcf_row + 1
        ws.cell(row=s2_df_row, column=1, value="Discount Factor")
        for i in range(stage2_years):
            yr = stage1_years + i + 1
            ws.cell(row=s2_df_row, column=2 + i,
                    value=f"=1/(1+$B${wacc_row})^{yr}")
            ws.cell(row=s2_df_row, column=2 + i).number_format = '0.0000'

        # PV
        s2_pv_row = s2_df_row + 1
        ws.cell(row=s2_pv_row, column=1, value="PV of FCF")
        for i in range(stage2_years):
            cl = get_column_letter(2 + i)
            ws.cell(row=s2_pv_row, column=2 + i,
                    value=f"={cl}{s2_fcf_row}*{cl}{s2_df_row}")
            ws.cell(row=s2_pv_row, column=2 + i).number_format = '#,##0.0'

        # Stage 2 PV
        s2_pv_total = s2_pv_row + 1
        ws.cell(row=s2_pv_total, column=1, value="PV of Stage 2 FCFs")
        ws.cell(row=s2_pv_total, column=1).font = Font(bold=True)
        last_s2_col = get_column_letter(1 + stage2_years)
        ws.cell(row=s2_pv_total, column=2,
                value=f"=SUM(B{s2_pv_row}:{last_s2_col}{s2_pv_row})")
        ws.cell(row=s2_pv_total, column=2).number_format = '#,##0.0'
        ws.cell(row=s2_pv_total, column=2).font = Font(bold=True)
        ws.cell(row=s2_pv_total, column=2).border = BOTTOM_BORDER

        # ============ STAGE 3: TERMINAL VALUE ============
        s3_start = s2_pv_total + 2
        self._add_section_header(ws, "STAGE 3: TERMINAL VALUE (Gordon Growth)", s3_start, 1)

        s3_start += 1
        ws.cell(row=s3_start, column=1, value="Terminal Year FCF")
        ws.cell(row=s3_start, column=2,
                value=f"={last_s2_col}{s2_fcf_row}*(1+$B$10)")
        ws.cell(row=s3_start, column=2).number_format = '#,##0.0'
        term_fcf_row = s3_start

        s3_start += 1
        ws.cell(row=s3_start, column=1, value="Terminal Value")
        ws.cell(row=s3_start, column=2,
                value=f"=IFERROR(B{term_fcf_row}/($B${wacc_row}-$B$10),0)")
        ws.cell(row=s3_start, column=2).number_format = '#,##0.0'
        ws.cell(row=s3_start, column=2).font = Font(bold=True)
        tv_row = s3_start

        s3_start += 1
        total_years = stage1_years + stage2_years
        ws.cell(row=s3_start, column=1, value="PV of Terminal Value")
        ws.cell(row=s3_start, column=2,
                value=f"=B{tv_row}/(1+$B${wacc_row})^{total_years}")
        ws.cell(row=s3_start, column=2).number_format = '#,##0.0'
        ws.cell(row=s3_start, column=2).font = Font(bold=True)
        ws.cell(row=s3_start, column=2).border = BOTTOM_BORDER
        pv_tv_row = s3_start

        # ============ VALUATION SUMMARY ============
        val_start = pv_tv_row + 2
        self._add_section_header(ws, "VALUATION SUMMARY", val_start, 1)
        val_start += 1

        ws.cell(row=val_start, column=1, value="Component")
        ws.cell(row=val_start, column=2, value="Value ($M)")
        ws.cell(row=val_start, column=3, value="% of EV")
        self._format_header_row(ws, val_start, 1, 3)

        vs_s1 = val_start + 1
        ws.cell(row=vs_s1, column=1, value="PV Stage 1 (High Growth)")
        ws.cell(row=vs_s1, column=2, value=f"=B{s1_pv_row}")
        ws.cell(row=vs_s1, column=2).number_format = '#,##0.0'

        vs_s2 = val_start + 2
        ws.cell(row=vs_s2, column=1, value="PV Stage 2 (Transition)")
        ws.cell(row=vs_s2, column=2, value=f"=B{s2_pv_total}")
        ws.cell(row=vs_s2, column=2).number_format = '#,##0.0'

        vs_s3 = val_start + 3
        ws.cell(row=vs_s3, column=1, value="PV Stage 3 (Terminal)")
        ws.cell(row=vs_s3, column=2, value=f"=B{pv_tv_row}")
        ws.cell(row=vs_s3, column=2).number_format = '#,##0.0'

        ev_row = val_start + 4
        ws.cell(row=ev_row, column=1, value="Enterprise Value")
        ws.cell(row=ev_row, column=1).font = Font(bold=True)
        ws.cell(row=ev_row, column=2,
                value=f"=B{vs_s1}+B{vs_s2}+B{vs_s3}")
        ws.cell(row=ev_row, column=2).number_format = '#,##0.0'
        ws.cell(row=ev_row, column=2).font = Font(bold=True)
        ws.cell(row=ev_row, column=2).border = DOUBLE_BORDER
        ws.cell(row=ev_row, column=2).fill = PatternFill(
            start_color="90EE90", end_color="90EE90", fill_type="solid")

        # % of EV
        for r in [vs_s1, vs_s2, vs_s3]:
            ws.cell(row=r, column=3,
                    value=f"=IFERROR(B{r}/B{ev_row},0)")
            ws.cell(row=r, column=3).number_format = '0.0%'
        ws.cell(row=ev_row, column=3, value=1.0)
        ws.cell(row=ev_row, column=3).number_format = '0.0%'
        ws.cell(row=ev_row, column=3).font = Font(bold=True)

        # Data bars on % of EV
        from openpyxl.formatting.rule import DataBarRule
        ws.conditional_formatting.add(
            f"C{vs_s1}:C{vs_s3}",
            DataBarRule(start_type='min', end_type='max',
                       color='4472C4', showValue=True))

        # --- EV Bridge ---
        br_row = ev_row + 2
        net_debt_val = a.ltm_ebitda * (a.senior_debt_multiple + a.sub_debt_multiple)
        ws.cell(row=br_row, column=1, value="Less: Net Debt")
        ws.cell(row=br_row, column=2, value=-net_debt_val)
        ws.cell(row=br_row, column=2).number_format = '#,##0.0'
        self._format_input_cell(ws, br_row, 2)
        nd_row = br_row

        br_row += 1
        ws.cell(row=br_row, column=1, value="Equity Value")
        ws.cell(row=br_row, column=1).font = Font(bold=True)
        ws.cell(row=br_row, column=2,
                value=f"=B{ev_row}+B{nd_row}")
        ws.cell(row=br_row, column=2).number_format = '#,##0.0'
        ws.cell(row=br_row, column=2).font = Font(bold=True)
        ws.cell(row=br_row, column=2).border = BOTTOM_BORDER
        eq_val_row = br_row

        br_row += 1
        shares = data.get('shares_outstanding', 100.0)
        ws.cell(row=br_row, column=1, value="Shares Outstanding (M)")
        ws.cell(row=br_row, column=2, value=shares)
        ws.cell(row=br_row, column=2).number_format = '#,##0.0'
        self._format_input_cell(ws, br_row, 2)
        shares_row = br_row

        br_row += 1
        ws.cell(row=br_row, column=1, value="Implied Price per Share")
        ws.cell(row=br_row, column=1).font = Font(bold=True)
        ws.cell(row=br_row, column=2,
                value=f"=IFERROR(B{eq_val_row}/B{shares_row},0)")
        ws.cell(row=br_row, column=2).number_format = '$#,##0.00'
        ws.cell(row=br_row, column=2).font = Font(bold=True, size=14)
        ws.cell(row=br_row, column=2).fill = PatternFill(
            start_color="90EE90", end_color="90EE90", fill_type="solid")
        ws.cell(row=br_row, column=2).border = DOUBLE_BORDER

        # --- Implied EV/EBITDA ---
        br_row += 1
        ws.cell(row=br_row, column=1, value="Implied EV/EBITDA")
        ws.cell(row=br_row, column=2,
                value=f"=IFERROR(B{ev_row}/B5*B6/{a.ltm_ebitda if a.ltm_ebitda else 1},0)")
        # Simplified: just use LTM EBITDA
        ws.cell(row=br_row, column=2,
                value=f"=IFERROR(B{ev_row}/{a.ltm_ebitda},0)")
        ws.cell(row=br_row, column=2).number_format = '0.0"x"'
        ws.cell(row=br_row, column=2).font = Font(bold=True)

        ws.column_dimensions['A'].width = 35
        max_cols = max(stage1_years, stage2_years)
        for i in range(max_cols + 1):
            ws.column_dimensions[get_column_letter(2 + i)].width = 14

        self.cell_map['multi_stage_dcf'] = {
            'pv_stage1': f'B{s1_pv_row}',
            'pv_stage2': f'B{s2_pv_total}',
            'pv_terminal': f'B{pv_tv_row}',
            'enterprise_value': f'B{ev_row}',
            'equity_value': f'B{eq_val_row}',
            'wacc': f'B{wacc_row}',
        }

        return self
