#!/usr/bin/env python3
"""
Capital Structure Excel Modules Mixin.

Contains: Dividend Recap, Refinancing Analysis,
Cap Table Waterfall, Sponsor Economics.

Smart formula upgrade: Cross-sheet references to Assumptions, Sources & Uses,
Operating Model, Debt Schedule, and Returns Analysis when those modules have
been built (detected via self.cell_map).  Falls back to hardcoded inputs when
upstream modules are absent — no behaviour change for standalone usage.
"""

from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from .formula_builder import FormulaBuilder as FB
from ._base import (INPUT_FILL, INPUT_FONT, HEADER_FILL, HEADER_FONT,
                     BOTTOM_BORDER, DOUBLE_BORDER, ModelDepth)
from typing import Dict


class CapitalStructureMixin:
    """Mixin providing capital structure analysis modules."""

    # ============================================================
    # MODULE: DIVIDEND RECAPITALIZATION
    # ============================================================

    def add_dividend_recap(self, data: Dict = None) -> 'CapitalStructureMixin':
        """Add Dividend Recapitalization analysis with live Excel formulas.

        Cross-references (when available):
        - Assumptions: LTM EBITDA, entry multiple, growth, hold period
        - Sources & Uses: total debt, sponsor equity
        - Debt Schedule: total debt row for DS cross-check
        """
        ws = self.wb.create_sheet("Dividend Recap")
        self.sheets_created.append("Dividend Recap")

        a = self.assumptions
        data = data or {}

        # Upstream module lookups
        ds = self.cell_map.get('debt_schedule', {})

        # Title
        self._add_title(ws, f"{self.company_name} - Dividend Recapitalization Analysis", 1, 1)

        # ── CURRENT CAPITAL STRUCTURE (rows 3-7) ──
        self._add_section_header(ws, "CURRENT CAPITAL STRUCTURE ($M)", 3, 1)

        ws.cell(row=4, column=1, value="Enterprise Value")
        ws.cell(row=4, column=2, value="='Assumptions'!B6*'Assumptions'!B7")
        ws.cell(row=4, column=2).number_format = '#,##0.0'

        su = self.cell_map.get('sources_uses', {})
        su_debt = su.get('total_debt', None)
        ws.cell(row=5, column=1, value="Less: Existing Debt")
        if su_debt:
            ws.cell(row=5, column=2, value=f"=-'Sources & Uses'!{su_debt}")
        else:
            ws.cell(row=5, column=2, value=f"=-'Assumptions'!B6*('Assumptions'!B12+'Assumptions'!B15)")
        ws.cell(row=5, column=2).number_format = '(#,##0.0)'

        ws.cell(row=6, column=1, value="Equity Value")
        ws.cell(row=6, column=2, value="=B4+B5")
        ws.cell(row=6, column=2).number_format = '#,##0.0'
        ws.cell(row=6, column=2).font = Font(bold=True)

        ws.cell(row=7, column=1, value="Current Leverage")
        ws.cell(row=7, column=2, value="=IFERROR(-B5/'Assumptions'!B6,0)")
        ws.cell(row=7, column=2).number_format = '0.0x'

        # ── RECAP PARAMETERS (rows 9-17) ──
        self._add_section_header(ws, "DIVIDEND RECAP PARAMETERS", 9, 1)

        recap_year = data.get('recap_year', 2)
        new_debt_amount = data.get('new_debt_amount', a.ltm_ebitda * 1.5)
        recap_rate = data.get('recap_debt_rate', 0.09)

        ws.cell(row=10, column=1, value="Recap Timing (Year)")
        ws.cell(row=10, column=2, value=recap_year)
        self._format_input_cell(ws, 10, 2)

        ws.cell(row=11, column=1, value="EBITDA at Recap")
        # Grow LTM EBITDA by avg growth for recap_year periods
        ws.cell(row=11, column=2, value=f"='Assumptions'!B6*(1+'Assumptions'!C{self.cell_map['assumptions']['rev_growth_row']})^B10")
        ws.cell(row=11, column=2).number_format = '#,##0.0'

        ws.cell(row=12, column=1, value="New Debt Amount")
        ws.cell(row=12, column=2, value=new_debt_amount)
        ws.cell(row=12, column=2).number_format = '#,##0.0'
        self._format_input_cell(ws, 12, 2)

        ws.cell(row=13, column=1, value="New Debt Rate")
        ws.cell(row=13, column=2, value=recap_rate)
        ws.cell(row=13, column=2).number_format = '0.0%'
        self._format_input_cell(ws, 13, 2)

        # Existing debt at recap (20% paydown/year approximation)
        ws.cell(row=14, column=1, value="Existing Debt at Recap")
        ws.cell(row=14, column=2, value="=-B5*0.8^B10")
        ws.cell(row=14, column=2).number_format = '#,##0.0'

        # DS Cross-check row: when Debt Schedule exists, show total debt at
        # the recap year column for manual validation of the 0.8^year approx.
        ds_cross_check_row = 15
        if ds and ds.get('total_debt_row'):
            # Default recap year = 2 → column offset from LTM.
            # The Debt Schedule's LTM column is typically the first data column;
            # year columns follow sequentially.  Use Y2 col as a reasonable default.
            ds_total_debt_row = ds['total_debt_row']
            # Recap year 2 data column: typically LTM col + recap_year
            om = self.cell_map.get('operating_model', {})
            ltm_col = om.get('ltm_col', 2)
            recap_col_num = ltm_col + recap_year
            recap_col_letter = get_column_letter(recap_col_num)
            ws.cell(row=ds_cross_check_row, column=1, value=f"DS Cross-check (Year {recap_year})")
            ws.cell(row=ds_cross_check_row, column=2,
                    value=f"='Debt Schedule'!{recap_col_letter}{ds_total_debt_row}")
            ws.cell(row=ds_cross_check_row, column=2).number_format = '#,##0.0'
            ws.cell(row=ds_cross_check_row, column=2).font = Font(italic=True, color="808080")
            total_post_row = 16
            leverage_post_row = 17
        else:
            # No DS available — skip cross-check, keep original row numbering
            total_post_row = 15
            leverage_post_row = 16

        ws.cell(row=total_post_row, column=1, value="Total Debt Post-Recap")
        ws.cell(row=total_post_row, column=2, value="=B14+B12")
        ws.cell(row=total_post_row, column=2).number_format = '#,##0.0'
        ws.cell(row=total_post_row, column=2).font = Font(bold=True)

        ws.cell(row=leverage_post_row, column=1, value="Post-Recap Leverage")
        ws.cell(row=leverage_post_row, column=2, value=f"=IFERROR(B{total_post_row}/B11,0)")
        ws.cell(row=leverage_post_row, column=2).number_format = '0.0x'
        ws.cell(row=leverage_post_row, column=2).font = Font(bold=True)

        # ── DIVIDEND DISTRIBUTION ──
        dist_header_row = leverage_post_row + 2
        self._add_section_header(ws, "DIVIDEND DISTRIBUTION", dist_header_row, 1)

        r = dist_header_row + 1
        ws.cell(row=r, column=1, value="Gross Proceeds")
        ws.cell(row=r, column=2, value="=B12")
        ws.cell(row=r, column=2).number_format = '#,##0.0'
        gross_row = r
        r += 1

        ws.cell(row=r, column=1, value="Less: Financing Fees (2%)")
        ws.cell(row=r, column=2, value="=-B12*0.02")
        ws.cell(row=r, column=2).number_format = '(#,##0.0)'
        fee_row = r
        r += 1

        net_div_row = r
        ws.cell(row=r, column=1, value="Net Dividend to Equity")
        ws.cell(row=r, column=2, value=f"=B{gross_row}+B{fee_row}")
        ws.cell(row=r, column=2).number_format = '#,##0.0'
        ws.cell(row=r, column=2).font = Font(bold=True)
        ws.cell(row=r, column=2).fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
        r += 1

        # ── RETURNS IMPACT ──
        impact_header_row = r + 1
        self._add_section_header(ws, "RETURNS IMPACT ANALYSIS", impact_header_row, 1)
        r = impact_header_row + 1

        ws.cell(row=r, column=1, value="Metric")
        ws.cell(row=r, column=2, value="Without Recap")
        ws.cell(row=r, column=3, value="With Recap")
        ws.cell(row=r, column=4, value="Impact")
        self._format_header_row(ws, r, 1, 4)
        r += 1

        # Initial Equity = from S&U or calculated
        su_equity = su.get('sponsor_equity', None)
        ws.cell(row=r, column=1, value="Initial Equity Investment")
        if su_equity:
            ws.cell(row=r, column=2, value=f"='Sources & Uses'!{su_equity}")
        else:
            ws.cell(row=r, column=2, value="=B4*(1+'Assumptions'!B19)-(-B5)")
        ws.cell(row=r, column=3, value=f"=B{r}")
        ws.cell(row=r, column=4, value=f"=C{r}-B{r}")
        for col in range(2, 5):
            ws.cell(row=r, column=col).number_format = '#,##0.0'
        equity_row = r
        r += 1

        # Dividend Proceeds
        ws.cell(row=r, column=1, value="Dividend Proceeds")
        ws.cell(row=r, column=2, value=0)
        ws.cell(row=r, column=3, value=f"=B{net_div_row}")
        ws.cell(row=r, column=4, value=f"=C{r}-B{r}")
        for col in range(2, 5):
            ws.cell(row=r, column=col).number_format = '#,##0.0'
        div_proceeds_row = r
        r += 1

        # Exit EV = exit EBITDA x exit multiple
        ws.cell(row=r, column=1, value="Exit EV")
        ws.cell(row=r, column=2, value=f"='Assumptions'!B6*(1+'Assumptions'!C{self.cell_map['assumptions']['rev_growth_row']})^'Assumptions'!B9*'Assumptions'!B8")
        ws.cell(row=r, column=3, value=f"=B{r}")
        ws.cell(row=r, column=4, value=0)
        for col in range(2, 5):
            ws.cell(row=r, column=col).number_format = '#,##0.0'
        exit_ev_row = r
        r += 1

        # Exit Equity (without recap: 50% debt paydown; with recap: 60% paydown)
        ws.cell(row=r, column=1, value="Exit Equity Proceeds")
        ws.cell(row=r, column=2, value=f"=B{exit_ev_row}-(-B5)*0.5")
        ws.cell(row=r, column=3, value=f"=C{exit_ev_row}-B{total_post_row}*0.6")
        ws.cell(row=r, column=4, value=f"=C{r}-B{r}")
        for col in range(2, 5):
            ws.cell(row=r, column=col).number_format = '#,##0.0'
        exit_equity_row = r
        r += 1

        # Total Proceeds
        ws.cell(row=r, column=1, value="Total Proceeds")
        ws.cell(row=r, column=2, value=f"=B{exit_equity_row}")
        ws.cell(row=r, column=3, value=f"=C{div_proceeds_row}+C{exit_equity_row}")
        ws.cell(row=r, column=4, value=f"=C{r}-B{r}")
        for col in range(2, 5):
            ws.cell(row=r, column=col).number_format = '#,##0.0'
        total_proceeds_row = r
        r += 1

        # MOIC
        ws.cell(row=r, column=1, value="MOIC")
        ws.cell(row=r, column=2, value=f"=IFERROR(B{total_proceeds_row}/B{equity_row},0)")
        ws.cell(row=r, column=3, value=f"=IFERROR(C{total_proceeds_row}/C{equity_row},0)")
        ws.cell(row=r, column=4, value=f"=C{r}-B{r}")
        for col in range(2, 5):
            ws.cell(row=r, column=col).number_format = '0.00x'
            ws.cell(row=r, column=col).font = Font(bold=True)
        moic_row = r
        r += 1

        # IRR
        ws.cell(row=r, column=1, value="IRR (approx)")
        ws.cell(row=r, column=2, value=f"=IFERROR(B{moic_row}^(1/'Assumptions'!B9)-1,0)")
        ws.cell(row=r, column=3, value=f"=IFERROR(C{moic_row}^(1/'Assumptions'!B9)-1,0)")
        ws.cell(row=r, column=4, value=f"=C{r}-B{r}")
        for col in range(2, 5):
            ws.cell(row=r, column=col).number_format = '0.0%'
            ws.cell(row=r, column=col).font = Font(bold=True)
        irr_row = r

        # Column widths
        ws.column_dimensions['A'].width = 30
        for col in ['B', 'C', 'D']:
            ws.column_dimensions[col].width = 15

        self.cell_map['dividend_recap'] = {
            'ev': 'B4',
            'existing_debt': 'B5',
            'equity_value': 'B6',
            'recap_year': 'B10',
            'new_debt': 'B12',
            'total_debt_post': f'B{total_post_row}',
            'net_dividend': f'B{net_div_row}',
            'moic_without': f'B{moic_row}',
            'moic_with': f'C{moic_row}',
            'irr_without': f'B{irr_row}',
            'irr_with': f'C{irr_row}',
        }

        return self

    # ============================================================
    # MODULE: REFINANCING ANALYSIS
    # ============================================================

    def add_refinancing_analysis(self, data: Dict = None) -> 'CapitalStructureMixin':
        """
        Add Refinancing Analysis sheet.
        Analyzes optimal timing to refinance at lower rates or extend maturities.

        Cross-references (when available):
        - Debt Schedule: senior/sub ending balances for cross-check notes
        - Assumptions: LTM EBITDA for post-refi leverage
        """
        ws = self.wb.create_sheet("Refinancing")
        self.sheets_created.append("Refinancing")

        a = self.assumptions
        data = data or {}

        # Upstream module lookups
        ds = self.cell_map.get('debt_schedule', {})
        om = self.cell_map.get('operating_model', {})
        assumptions_map = self.cell_map.get('assumptions', {})

        # Title
        self._add_title(ws, f"{self.company_name} - Refinancing Analysis", 1, 1)

        # Current Debt Structure
        self._add_section_header(ws, "CURRENT DEBT STRUCTURE", 3, 1)

        current_debt = data.get('current_debt', [
            {'tranche': 'Term Loan A', 'amount': 150.0, 'rate': 0.085, 'maturity': 2027, 'callable': 'Par'},
            {'tranche': 'Term Loan B', 'amount': 100.0, 'rate': 0.095, 'maturity': 2028, 'callable': '101'},
            {'tranche': 'Senior Notes', 'amount': 50.0, 'rate': 0.0875, 'maturity': 2029, 'callable': 'NC-2'},
        ])

        headers = ["Tranche", "Amount ($M)", "Rate", "Maturity", "Call Protection"]
        row = 4
        for i, h in enumerate(headers):
            ws.cell(row=row, column=1 + i, value=h)
        self._format_header_row(ws, row, 1, len(headers))
        row += 1

        old_start = row
        for debt in current_debt:
            ws.cell(row=row, column=1, value=debt['tranche'])
            ws.cell(row=row, column=2, value=debt['amount'])
            ws.cell(row=row, column=2).number_format = '#,##0.0'
            self._format_input_cell(ws, row, 2)
            ws.cell(row=row, column=3, value=debt['rate'])
            ws.cell(row=row, column=3).number_format = '0.00%'
            self._format_input_cell(ws, row, 3)
            ws.cell(row=row, column=4, value=debt['maturity'])
            ws.cell(row=row, column=5, value=debt['callable'])
            row += 1
        old_end = row - 1

        row += 1
        old_total_row = row
        ws.cell(row=row, column=1, value="Total / Weighted Average")
        ws.cell(row=row, column=2, value=f"=SUM(B{old_start}:B{old_end})")
        ws.cell(row=row, column=2).number_format = '#,##0.0'
        ws.cell(row=row, column=2).font = Font(bold=True)
        ws.cell(row=row, column=3, value=f"=IFERROR(SUMPRODUCT(B{old_start}:B{old_end},C{old_start}:C{old_end})/B{row},0)")
        ws.cell(row=row, column=3).number_format = '0.00%'
        ws.cell(row=row, column=3).font = Font(bold=True)
        total_debt = sum(d['amount'] for d in current_debt)  # keep for default new_structure
        row += 1

        # ── DS Cross-reference notes (when Debt Schedule exists) ──
        if ds and ds.get('senior_ending_row'):
            # Determine the entry/LTM column in the Debt Schedule
            entry_col_num = om.get('ltm_col', 2)
            entry_col_letter = get_column_letter(entry_col_num)

            ws.cell(row=row, column=1, value="Senior Debt per DS")
            ws.cell(row=row, column=2,
                    value=f"='Debt Schedule'!{entry_col_letter}{ds['senior_ending_row']}")
            ws.cell(row=row, column=2).number_format = '#,##0.0'
            ws.cell(row=row, column=2).font = Font(italic=True, color="808080")
            row += 1

            if ds.get('sub_ending_row'):
                ws.cell(row=row, column=1, value="Sub Debt per DS")
                ws.cell(row=row, column=2,
                        value=f"='Debt Schedule'!{entry_col_letter}{ds['sub_ending_row']}")
                ws.cell(row=row, column=2).number_format = '#,##0.0'
                ws.cell(row=row, column=2).font = Font(italic=True, color="808080")
                row += 1

        row += 2

        # Market Conditions
        self._add_section_header(ws, "CURRENT MARKET CONDITIONS", row, 1)
        row += 1

        market = data.get('market_conditions', {
            'tla_rate': 0.070,
            'tlb_rate': 0.080,
            'notes_rate': 0.075,
            'sofr': 0.045,
        })

        ws.cell(row=row, column=1, value="New TLA Rate (est.)")
        ws.cell(row=row, column=2, value=market['tla_rate'])
        ws.cell(row=row, column=2).number_format = '0.00%'
        self._format_input_cell(ws, row, 2)
        row += 1

        ws.cell(row=row, column=1, value="New TLB Rate (est.)")
        ws.cell(row=row, column=2, value=market['tlb_rate'])
        ws.cell(row=row, column=2).number_format = '0.00%'
        self._format_input_cell(ws, row, 2)
        row += 1

        ws.cell(row=row, column=1, value="New Notes Rate (est.)")
        ws.cell(row=row, column=2, value=market['notes_rate'])
        ws.cell(row=row, column=2).number_format = '0.00%'
        self._format_input_cell(ws, row, 2)
        row += 1

        ws.cell(row=row, column=1, value="Current SOFR")
        ws.cell(row=row, column=2, value=market['sofr'])
        ws.cell(row=row, column=2).number_format = '0.00%'
        row += 3

        # Refinancing Scenario
        self._add_section_header(ws, "REFINANCING SCENARIO", row, 1)
        row += 1

        new_structure = data.get('new_structure', [
            {'tranche': 'New Term Loan', 'amount': total_debt, 'rate': market['tlb_rate'], 'maturity': 2031},
        ])

        headers = ["New Tranche", "Amount ($M)", "Rate", "Maturity"]
        for i, h in enumerate(headers):
            ws.cell(row=row, column=1 + i, value=h)
        self._format_header_row(ws, row, 1, 4)
        row += 1

        new_start = row
        for debt in new_structure:
            ws.cell(row=row, column=1, value=debt['tranche'])
            ws.cell(row=row, column=2, value=debt['amount'])
            ws.cell(row=row, column=2).number_format = '#,##0.0'
            self._format_input_cell(ws, row, 2)
            ws.cell(row=row, column=3, value=debt['rate'])
            ws.cell(row=row, column=3).number_format = '0.00%'
            self._format_input_cell(ws, row, 3)
            ws.cell(row=row, column=4, value=debt['maturity'])
            row += 1
        new_end = row - 1

        row += 1
        new_total_row = row
        ws.cell(row=row, column=1, value="Total / Weighted Average")
        ws.cell(row=row, column=2, value=f"=SUM(B{new_start}:B{new_end})")
        ws.cell(row=row, column=2).number_format = '#,##0.0'
        ws.cell(row=row, column=2).font = Font(bold=True)
        ws.cell(row=row, column=3, value=f"=IFERROR(SUMPRODUCT(B{new_start}:B{new_end},C{new_start}:C{new_end})/B{row},0)")
        ws.cell(row=row, column=3).number_format = '0.00%'
        ws.cell(row=row, column=3).font = Font(bold=True)
        row += 3

        # Savings Analysis
        self._add_section_header(ws, "ANNUAL SAVINGS ANALYSIS", row, 1)
        row += 1

        old_int_row = row
        ws.cell(row=row, column=1, value="Current Interest Expense")
        ws.cell(row=row, column=2, value=f"=SUMPRODUCT(B{old_start}:B{old_end},C{old_start}:C{old_end})")
        ws.cell(row=row, column=2).number_format = '#,##0.0'
        row += 1

        new_int_row = row
        ws.cell(row=row, column=1, value="New Interest Expense")
        ws.cell(row=row, column=2, value=f"=SUMPRODUCT(B{new_start}:B{new_end},C{new_start}:C{new_end})")
        ws.cell(row=row, column=2).number_format = '#,##0.0'
        row += 1

        savings_row = row
        ws.cell(row=row, column=1, value="Annual Interest Savings")
        ws.cell(row=row, column=2, value=f"=B{old_int_row}-B{new_int_row}")
        ws.cell(row=row, column=2).number_format = '#,##0.0'
        ws.cell(row=row, column=2).font = Font(bold=True)
        ws.cell(row=row, column=2).fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
        row += 1

        ws.cell(row=row, column=1, value="Rate Reduction")
        ws.cell(row=row, column=2, value=f"=(C{old_total_row}-C{new_total_row})*10000")
        ws.cell(row=row, column=2).number_format = '0" bps"'
        row += 3

        # Transaction Costs
        self._add_section_header(ws, "TRANSACTION COSTS", row, 1)
        row += 1

        call_row = row
        ws.cell(row=row, column=1, value="Call Premium / Make-Whole")
        ws.cell(row=row, column=2, value=f"=B{old_total_row}*0.01")
        ws.cell(row=row, column=2).number_format = '#,##0.0'
        self._format_input_cell(ws, row, 2)
        row += 1

        arr_row = row
        ws.cell(row=row, column=1, value="Arrangement Fee (1%)")
        ws.cell(row=row, column=2, value=f"=B{new_total_row}*0.01")
        ws.cell(row=row, column=2).number_format = '#,##0.0'
        row += 1

        legal_row = row
        ws.cell(row=row, column=1, value="Legal & Advisory")
        ws.cell(row=row, column=2, value=0.5)
        ws.cell(row=row, column=2).number_format = '#,##0.0'
        self._format_input_cell(ws, row, 2)
        row += 1

        total_costs_row = row
        ws.cell(row=row, column=1, value="Total Transaction Costs")
        ws.cell(row=row, column=2, value=f"=B{call_row}+B{arr_row}+B{legal_row}")
        ws.cell(row=row, column=2).number_format = '#,##0.0'
        ws.cell(row=row, column=2).font = Font(bold=True)
        row += 2

        # Payback
        payback_row = row
        ws.cell(row=row, column=1, value="Payback Period (years)")
        ws.cell(row=row, column=2, value=f"=IFERROR(B{total_costs_row}/B{savings_row},999)")
        ws.cell(row=row, column=2).number_format = '0.0'
        ws.cell(row=row, column=2).font = Font(bold=True)
        ws.cell(row=row, column=3, value=f'=IF(B{row}<=2,"Attractive",IF(B{row}<=3,"Marginal","Not recommended"))')
        ws.cell(row=row, column=3).font = Font(italic=True)
        row += 2

        # ── Post-Refi Leverage cross-ref (when Assumptions exists) ──
        if assumptions_map:
            ws.cell(row=row, column=1, value="Post-Refi Leverage")
            ws.cell(row=row, column=2,
                    value=f"=IFERROR(B{new_total_row}/'Assumptions'!B6,0)")
            ws.cell(row=row, column=2).number_format = '0.0x'
            ws.cell(row=row, column=2).font = Font(bold=True)

        # Column widths
        ws.column_dimensions['A'].width = 30
        for col in ['B', 'C', 'D', 'E']:
            ws.column_dimensions[col].width = 15

        return self

    # ============================================================
    # MODULE: CAP TABLE / WATERFALL
    # ============================================================

    def add_cap_table_waterfall(self, data: Dict = None) -> 'CapitalStructureMixin':
        """
        Add Cap Table / Waterfall analysis sheet.
        Detailed equity waterfall with preferred returns, participation, catch-ups.

        Cross-references (when available):
        - Sources & Uses: equity cross-check
        - Assumptions: hold period
        - Operating Model + Debt Schedule: base case exit equity computation
        """
        ws = self.wb.create_sheet("Cap Table Waterfall")
        self.sheets_created.append("Cap Table Waterfall")

        a = self.assumptions
        data = data or {}

        # Upstream module lookups
        su = self.cell_map.get('sources_uses', {})
        assumptions_map = self.cell_map.get('assumptions', {})
        om = self.cell_map.get('operating_model', {})
        ds = self.cell_map.get('debt_schedule', {})

        # Title
        self._add_title(ws, f"{self.company_name} - Equity Waterfall Analysis", 1, 1)

        # Cap Table
        self._add_section_header(ws, "CAPITALIZATION TABLE", 3, 1)

        cap_table = data.get('cap_table', [
            {'investor': 'Sponsor Fund', 'invested': 150.0, 'ownership': 0.80, 'type': 'Common'},
            {'investor': 'Co-Investors', 'invested': 25.0, 'ownership': 0.13, 'type': 'Common'},
            {'investor': 'Management', 'invested': 5.0, 'ownership': 0.05, 'type': 'Common'},
            {'investor': 'Rollover Equity', 'invested': 4.0, 'ownership': 0.02, 'type': 'Common'},
        ])

        headers = ["Investor", "Invested ($M)", "Ownership %", "Security Type"]
        row = 4
        for i, h in enumerate(headers):
            ws.cell(row=row, column=1 + i, value=h)
        self._format_header_row(ws, row, 1, 4)
        row += 1

        cap_start = row
        for investor in cap_table:
            ws.cell(row=row, column=1, value=investor['investor'])
            ws.cell(row=row, column=2, value=investor['invested'])
            ws.cell(row=row, column=2).number_format = '#,##0.0'
            self._format_input_cell(ws, row, 2)
            ws.cell(row=row, column=3, value=investor['ownership'])
            ws.cell(row=row, column=3).number_format = '0.0%'
            ws.cell(row=row, column=4, value=investor['type'])
            row += 1
        cap_end = row - 1

        row += 1
        ti_row = row  # Total Invested row
        ws.cell(row=row, column=1, value="Total")
        ws.cell(row=row, column=2, value=f"=SUM(B{cap_start}:B{cap_end})")
        ws.cell(row=row, column=2).number_format = '#,##0.0'
        ws.cell(row=row, column=2).font = Font(bold=True)
        ws.cell(row=row, column=3, value=f"=SUM(C{cap_start}:C{cap_end})")
        ws.cell(row=row, column=3).number_format = '0.0%'
        ws.cell(row=row, column=3).font = Font(bold=True)
        row += 1

        # ── S&U Equity Cross-check (when Sources & Uses exists) ──
        su_equity_cell = su.get('equity', None)
        if su_equity_cell:
            ws.cell(row=row, column=1, value="S&U Equity Cross-check")
            ws.cell(row=row, column=2,
                    value=f"='Sources & Uses'!{su_equity_cell}")
            ws.cell(row=row, column=2).number_format = '#,##0.0'
            ws.cell(row=row, column=2).font = Font(italic=True, color="808080")
            row += 1

        row += 2

        # Waterfall Terms
        self._add_section_header(ws, "WATERFALL STRUCTURE", row, 1)
        row += 1

        waterfall_terms = data.get('waterfall_terms', {
            'preferred_return': 0.08,
            'catch_up_pct': 1.00,
            'catch_up_split': 0.20,
            'carried_interest': 0.20,
            'gp_commitment': 0.02,
        })

        pref_row = row
        ws.cell(row=row, column=1, value="Preferred Return (Hurdle)")
        ws.cell(row=row, column=2, value=waterfall_terms['preferred_return'])
        ws.cell(row=row, column=2).number_format = '0.0%'
        self._format_input_cell(ws, row, 2)
        row += 1

        ws.cell(row=row, column=1, value="GP Catch-Up")
        ws.cell(row=row, column=2, value=waterfall_terms['catch_up_pct'])
        ws.cell(row=row, column=2).number_format = '0%'
        row += 1

        carry_row = row
        ws.cell(row=row, column=1, value="Carried Interest")
        ws.cell(row=row, column=2, value=waterfall_terms['carried_interest'])
        ws.cell(row=row, column=2).number_format = '0.0%'
        self._format_input_cell(ws, row, 2)
        row += 1

        ws.cell(row=row, column=1, value="GP Commitment")
        ws.cell(row=row, column=2, value=waterfall_terms['gp_commitment'])
        ws.cell(row=row, column=2).number_format = '0.0%'
        row += 1

        hold_row = row
        ws.cell(row=row, column=1, value="Hold Period (years)")
        # Reference Assumptions hold period when available; otherwise hardcoded 5
        if assumptions_map:
            ws.cell(row=row, column=2, value="='Assumptions'!B9")
        else:
            ws.cell(row=row, column=2, value=5)
            self._format_input_cell(ws, row, 2)
        row += 1

        # Preferred Return Threshold = Invested x (1+hurdle)^hold
        pt_row = row
        ws.cell(row=row, column=1, value="Pref Return Threshold")
        ws.cell(row=row, column=2, value=f"=B{ti_row}*((1+B{pref_row})^B{hold_row})")
        ws.cell(row=row, column=2).number_format = '#,##0.0'
        row += 2

        # Exit Scenarios Waterfall
        self._add_section_header(ws, "WATERFALL BY EXIT VALUE", row, 1)
        row += 1

        # Exit values as numbers for formula references
        ev_row = row
        ws.cell(row=row, column=1, value="Exit Equity Value ($M)")
        exit_values = [200, 300, 400, 500, 600]
        for i, ev in enumerate(exit_values):
            ws.cell(row=row, column=2 + i, value=ev)
            ws.cell(row=row, column=2 + i).number_format = '#,##0'
            self._format_input_cell(ws, row, 2 + i)
        self._format_header_row(ws, row, 1, 6)
        row += 1

        # ── Base Case Exit Equity from model (when OM + DS exist) ──
        base_case_exit_row = None
        if om and ds and om.get('ebitda_row') and ds.get('total_debt_row'):
            base_case_exit_row = row
            ws.cell(row=row, column=1, value="Base Case Exit Equity")
            ws.cell(row=row, column=1).font = Font(italic=True, color="808080")

            # Exit column = LTM col + hold period (from assumptions or default 5)
            ltm_col = om.get('ltm_col', 2)
            exit_col_num = ltm_col + a.hold_period
            exit_col_letter = get_column_letter(exit_col_num)
            ebitda_row_num = om['ebitda_row']
            total_debt_row_num = ds['total_debt_row']

            # Exit Equity = Exit EBITDA * Exit Multiple - Total Debt at Exit
            base_formula = (
                f"='Operating Model'!{exit_col_letter}{ebitda_row_num}"
                f"*'Assumptions'!B8"
                f"-'Debt Schedule'!{exit_col_letter}{total_debt_row_num}"
            )
            ws.cell(row=row, column=2, value=base_formula)
            ws.cell(row=row, column=2).number_format = '#,##0.0'
            ws.cell(row=row, column=2).font = Font(italic=True, color="808080")
            row += 1

        # Reference shortcuts for formulas
        ti = f"$B${ti_row}"    # Total Invested
        pt = f"$B${pt_row}"    # Pref Threshold
        cr = f"$B${carry_row}" # Carry %

        # Return of Capital
        ws.cell(row=row, column=1, value="Return of Capital")
        for i in range(5):
            c = get_column_letter(2 + i)
            ws.cell(row=row, column=2 + i, value=f"=MIN({c}${ev_row},{ti})")
            ws.cell(row=row, column=2 + i).number_format = '#,##0.0'
        row += 1

        # Preferred Return (to hurdle)
        ws.cell(row=row, column=1, value="Preferred Return (to hurdle)")
        for i in range(5):
            c = get_column_letter(2 + i)
            ws.cell(row=row, column=2 + i, value=f"=MIN(MAX(0,{c}${ev_row}-{ti}),{pt}-{ti})")
            ws.cell(row=row, column=2 + i).number_format = '#,##0.0'
        row += 1

        # GP Catch-Up
        ws.cell(row=row, column=1, value="GP Catch-Up")
        for i in range(5):
            c = get_column_letter(2 + i)
            ws.cell(row=row, column=2 + i,
                    value=f"=IF({c}${ev_row}>{pt},({c}${ev_row}-{ti})*{cr},0)")
            ws.cell(row=row, column=2 + i).number_format = '#,##0.0'
        row += 1

        # Remaining (80/20 split)
        ws.cell(row=row, column=1, value="Remaining (80/20 split)")
        for i in range(5):
            c = get_column_letter(2 + i)
            ws.cell(row=row, column=2 + i,
                    value=f"=IF({c}${ev_row}>{pt},MAX(0,{c}${ev_row}-{pt})*0.8,0)")
            ws.cell(row=row, column=2 + i).number_format = '#,##0.0'
        row += 1

        # Total to LPs
        lp_row = row
        ws.cell(row=row, column=1, value="Total to LPs")
        ws.cell(row=row, column=1).font = Font(bold=True)
        for i in range(5):
            c = get_column_letter(2 + i)
            ws.cell(row=row, column=2 + i,
                    value=f"=IF({c}${ev_row}<={pt},{c}${ev_row},{c}${ev_row}-({c}${ev_row}-{ti})*{cr})")
            ws.cell(row=row, column=2 + i).number_format = '#,##0.0'
            ws.cell(row=row, column=2 + i).font = Font(bold=True)
            ws.cell(row=row, column=2 + i).fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
        row += 1

        # Total to GP (Carry)
        gp_row = row
        ws.cell(row=row, column=1, value="Total to GP (Carry)")
        for i in range(5):
            c = get_column_letter(2 + i)
            ws.cell(row=row, column=2 + i,
                    value=f"=IF({c}${ev_row}<={pt},0,({c}${ev_row}-{ti})*{cr})")
            ws.cell(row=row, column=2 + i).number_format = '#,##0.0'
        row += 1

        # LP MOIC
        ws.cell(row=row, column=1, value="LP MOIC")
        ws.cell(row=row, column=1).font = Font(bold=True)
        for i in range(5):
            c = get_column_letter(2 + i)
            ws.cell(row=row, column=2 + i, value=f"=IFERROR({c}{lp_row}/{ti},0)")
            ws.cell(row=row, column=2 + i).number_format = '0.00x'
            ws.cell(row=row, column=2 + i).font = Font(bold=True)
        row += 1

        # GP Carry ($M)
        ws.cell(row=row, column=1, value="GP Carry ($M)")
        for i in range(5):
            c = get_column_letter(2 + i)
            ws.cell(row=row, column=2 + i, value=f"={c}{gp_row}")
            ws.cell(row=row, column=2 + i).number_format = '#,##0.0'

        # Column widths
        ws.column_dimensions['A'].width = 25
        for col_letter in ['B', 'C', 'D', 'E', 'F']:
            ws.column_dimensions[col_letter].width = 12

        self.cell_map['cap_table_waterfall'] = {
            'total_invested_row': ti_row,
            'pref_return_row': pref_row,
            'carry_row': carry_row,
            'pref_threshold_row': pt_row,
            'hold_period_row': hold_row,
            'base_case_exit_row': base_case_exit_row,
            'lp_total_row': lp_row,
            'gp_total_row': gp_row,
        }

        return self

    # ============================================================
    # MODULE: SPONSOR ECONOMICS
    # ============================================================

    def add_sponsor_economics(self, data: Dict = None) -> 'CapitalStructureMixin':
        """Add Sponsor Economics with live Excel formulas.

        Cross-references (when available):
        - Sources & Uses: deal equity investment
        - Assumptions: hold period
        - Returns Analysis: model MOIC for cross-check
        """
        ws = self.wb.create_sheet("Sponsor Economics")
        self.sheets_created.append("Sponsor Economics")

        a = self.assumptions
        data = data or {}

        # Upstream module lookups
        su = self.cell_map.get('sources_uses', {})
        assumptions_map = self.cell_map.get('assumptions', {})
        ra = self.cell_map.get('returns_analysis', {})

        # Title
        self._add_title(ws, f"{self.company_name} - Sponsor Economics Analysis", 1, 1)

        # ── FUND PARAMETERS (rows 3-10) — all inputs ──
        self._add_section_header(ws, "FUND PARAMETERS", 3, 1)

        fund_params = data.get('fund_params', {
            'fund_size': 500.0, 'gp_commitment': 0.02, 'management_fee': 0.02,
            'carried_interest': 0.20, 'hurdle_rate': 0.08,
            'fund_life': 10, 'investment_period': 5,
        })

        fund_inputs = [
            (4, "Fund Size ($M)", fund_params['fund_size'], '#,##0'),
            (5, "GP Commitment", fund_params['gp_commitment'], '0.0%'),
            (6, "Management Fee", fund_params['management_fee'], '0.0%'),
            (7, "Carried Interest", fund_params['carried_interest'], '0.0%'),
            (8, "Hurdle Rate", fund_params['hurdle_rate'], '0.0%'),
            (9, "Fund Life (years)", fund_params['fund_life'], '0'),
            (10, "Investment Period (years)", fund_params['investment_period'], '0'),
        ]
        for row, name, value, fmt in fund_inputs:
            ws.cell(row=row, column=1, value=name)
            ws.cell(row=row, column=2, value=value)
            ws.cell(row=row, column=2).number_format = fmt
            self._format_input_cell(ws, row, 2)

        # ── DEAL PARAMETERS (rows 12-17) — inputs with smart cross-refs ──
        self._add_section_header(ws, "DEAL PARAMETERS", 12, 1)

        deal_params = data.get('deal_params', {
            'deal_equity': 150.0, 'deal_pct_of_fund': 0.30,
            'entry_moic': 2.5, 'hold_period': 5,
        })

        # Row 13: Deal Equity Investment — reference S&U when available
        ws.cell(row=13, column=1, value="Deal Equity Investment ($M)")
        su_equity_cell = su.get('equity', None)
        if su_equity_cell:
            ws.cell(row=13, column=2, value=f"='Sources & Uses'!{su_equity_cell}")
            # Keep number format but don't apply input styling (it's a formula now)
        else:
            ws.cell(row=13, column=2, value=deal_params['deal_equity'])
            self._format_input_cell(ws, 13, 2)
        ws.cell(row=13, column=2).number_format = '#,##0.0'

        ws.cell(row=14, column=1, value="% of Fund")
        ws.cell(row=14, column=2, value="=IFERROR(B13/B4,0)")
        ws.cell(row=14, column=2).number_format = '0.0%'

        # Row 15: Expected MOIC — keep as input; add model cross-check in col C
        ws.cell(row=15, column=1, value="Expected MOIC")
        ws.cell(row=15, column=2, value=deal_params['entry_moic'])
        ws.cell(row=15, column=2).number_format = '0.0x'
        self._format_input_cell(ws, 15, 2)

        if ra and ra.get('moic_row') and ra.get('moic_col'):
            moic_col_letter = get_column_letter(ra['moic_col'])
            moic_row_num = ra['moic_row']
            ws.cell(row=15, column=3, value=f"='Returns Analysis'!{moic_col_letter}{moic_row_num}")
            ws.cell(row=15, column=3).number_format = '0.0x'
            ws.cell(row=15, column=3).font = Font(italic=True, color="808080")
            ws.cell(row=15, column=4, value="Model MOIC")
            ws.cell(row=15, column=4).font = Font(italic=True, color="808080")

        # Row 16: Hold Period — reference Assumptions when available
        ws.cell(row=16, column=1, value="Hold Period (years)")
        if assumptions_map:
            ws.cell(row=16, column=2, value="='Assumptions'!B9")
        else:
            ws.cell(row=16, column=2, value=deal_params['hold_period'])
            self._format_input_cell(ws, 16, 2)

        # ── GP ECONOMICS FROM THIS DEAL (rows 18-28) — formulas ──
        self._add_section_header(ws, "GP ECONOMICS FROM THIS DEAL", 18, 1)

        ws.cell(row=19, column=1, value="Deal Investment")
        ws.cell(row=19, column=2, value="=B13")
        ws.cell(row=19, column=2).number_format = '#,##0.0'

        ws.cell(row=20, column=1, value="Deal Proceeds")
        ws.cell(row=20, column=2, value="=B13*B15")
        ws.cell(row=20, column=2).number_format = '#,##0.0'

        ws.cell(row=21, column=1, value="Deal Profit")
        ws.cell(row=21, column=2, value="=B20-B19")
        ws.cell(row=21, column=2).number_format = '#,##0.0'

        ws.cell(row=22, column=1, value="Deal IRR")
        ws.cell(row=22, column=2, value="=IFERROR(B15^(1/B16)-1,0)")
        ws.cell(row=22, column=2).number_format = '0.0%'
        ws.cell(row=22, column=2).font = Font(bold=True)

        ws.cell(row=24, column=1, value="GP Carried Interest")
        ws.cell(row=24, column=2, value="=B21*B7")
        ws.cell(row=24, column=2).number_format = '#,##0.0'
        ws.cell(row=24, column=2).fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")

        ws.cell(row=25, column=1, value="GP Co-Invest Profit")
        ws.cell(row=25, column=2, value="=B13*B5*(B15-1)")
        ws.cell(row=25, column=2).number_format = '#,##0.0'

        ws.cell(row=26, column=1, value="Total GP Economics (this deal)")
        ws.cell(row=26, column=2, value="=B24+B25")
        ws.cell(row=26, column=2).number_format = '#,##0.0'
        ws.cell(row=26, column=2).font = Font(bold=True)

        # ── FUND-LEVEL ECONOMICS (rows 28-40) — formulas ──
        self._add_section_header(ws, "FUND-LEVEL ECONOMICS (ILLUSTRATIVE)", 28, 1)

        ws.cell(row=29, column=1, value="Fund Gross MOIC (assumed)")
        ws.cell(row=29, column=2, value=data.get('fund_gross_moic', 2.0))
        ws.cell(row=29, column=2).number_format = '0.0x'
        self._format_input_cell(ws, 29, 2)

        ws.cell(row=31, column=1, value="")
        ws.cell(row=31, column=2, value="Amount ($M)")
        ws.cell(row=31, column=3, value="% of Fund")
        self._format_header_row(ws, 31, 1, 3)

        ws.cell(row=32, column=1, value="Fund Size (Committed)")
        ws.cell(row=32, column=2, value="=B4")
        ws.cell(row=32, column=2).number_format = '#,##0.0'

        ws.cell(row=33, column=1, value="Gross Proceeds")
        ws.cell(row=33, column=2, value="=B4*B29")
        ws.cell(row=33, column=2).number_format = '#,##0.0'

        ws.cell(row=34, column=1, value="Gross Profit")
        ws.cell(row=34, column=2, value="=B33-B32")
        ws.cell(row=34, column=2).number_format = '#,##0.0'
        ws.cell(row=34, column=3, value="=IFERROR(B34/B32,0)")
        ws.cell(row=34, column=3).number_format = '0.0%'

        # Management Fees = Fund x Fee% x InvPeriod + Fund x 0.5 x Fee% x (Life-InvPeriod)
        ws.cell(row=36, column=1, value="Management Fees (total)")
        ws.cell(row=36, column=2, value="=B4*B6*B10+B4*0.5*B6*(B9-B10)")
        ws.cell(row=36, column=2).number_format = '#,##0.0'
        ws.cell(row=36, column=3, value="=IFERROR(B36/B32,0)")
        ws.cell(row=36, column=3).number_format = '0.0%'

        # Carry = Profit x Carry% (if hurdle met)
        ws.cell(row=37, column=1, value="Carried Interest")
        ws.cell(row=37, column=2, value="=IF(B33>B4*(1+B8)^B9,B34*B7,0)")
        ws.cell(row=37, column=2).number_format = '#,##0.0'
        ws.cell(row=37, column=3, value="=IFERROR(B37/B32,0)")
        ws.cell(row=37, column=3).number_format = '0.0%'

        ws.cell(row=38, column=1, value="GP Co-Invest Profit")
        ws.cell(row=38, column=2, value="=B4*B5*(B29-1)")
        ws.cell(row=38, column=2).number_format = '#,##0.0'
        ws.cell(row=38, column=3, value="=IFERROR(B38/B32,0)")
        ws.cell(row=38, column=3).number_format = '0.0%'

        ws.cell(row=40, column=1, value="Total GP Revenue")
        ws.cell(row=40, column=2, value="=B36+B37+B38")
        ws.cell(row=40, column=2).number_format = '#,##0.0'
        ws.cell(row=40, column=2).font = Font(bold=True)
        ws.cell(row=40, column=2).fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")

        # Column widths
        ws.column_dimensions['A'].width = 35
        ws.column_dimensions['B'].width = 15
        ws.column_dimensions['C'].width = 15
        ws.column_dimensions['D'].width = 15

        self.cell_map['sponsor_economics'] = {
            'fund_size': 'B4',
            'carry_pct': 'B7',
            'deal_equity': 'B13',
            'deal_moic': 'B15',
            'deal_irr': 'B22',
            'gp_carry': 'B24',
            'total_gp_deal': 'B26',
            'total_gp_fund': 'B40',
        }

        return self
