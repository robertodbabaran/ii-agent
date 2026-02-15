#!/usr/bin/env python3
"""
Value Creation Excel Modules Mixin.

Contains: Value Creation Bridge, 100-Day Plan,
Exit Readiness, Management Incentive Plan.

Smart formula upgrades (2026-02-14):
- Value Creation Bridge uses cross-sheet formulas from Assumptions,
  Operating Model, Sources & Uses, and Debt Schedule when available.
- Fixed S&U equity key lookup ('equity' instead of 'sponsor_equity').
- Added cell_map entries for all modules.
- SUMPRODUCT for Exit Readiness overall score.
"""

from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from .formula_builder import FormulaBuilder as FB
from ._base import (INPUT_FILL, INPUT_FONT, HEADER_FILL, HEADER_FONT,
                     BOTTOM_BORDER, DOUBLE_BORDER, ModelDepth)
from typing import Dict


class ValueCreationMixin:
    """Mixin providing value creation analysis modules."""

    def add_value_creation_bridge(self, data: Dict = None) -> 'ValueCreationMixin':
        """Add detailed Value Creation Bridge sheet.

        When upstream modules (assumptions, operating_model, sources_uses,
        debt_schedule) exist in self.cell_map, writes real Excel formulas
        for EBITDA Growth, Multiple Expansion, and Debt Paydown attribution.
        Falls back to hardcoded percentage-based attribution otherwise.
        """
        ws = self.wb.create_sheet("Value Creation")
        self.sheets_created.append("Value Creation")
        a = self.assumptions
        data = data or {}

        self._add_title(ws, f"{self.company_name} - Value Creation Bridge", 1, 1)

        # --- Detect upstream modules ---
        su = self.cell_map.get('sources_uses', {})
        om = self.cell_map.get('operating_model', {})
        asn = self.cell_map.get('assumptions', {})
        ds = self.cell_map.get('debt_schedule', {})

        # Sheet name constants for formula references
        A = "Assumptions"
        OM = "Operating Model"
        SU = "Sources & Uses"
        DS = "Debt Schedule"

        # Determine if we can build formula-driven attribution
        has_full_upstream = bool(om and asn and su and ds)

        # Exit column for projection end (e.g., G for 5-year)
        n = getattr(self, 'projection_years', 5)
        exit_col = get_column_letter(2 + n)  # col B=LTM, C=Y1, ... G=Y5

        # --- SECTION: VALUE CREATION ATTRIBUTION ---
        self._add_section_header(ws, "VALUE CREATION ATTRIBUTION", 3, 1)

        # Row 4: Entry Equity Value
        row = 4
        ws.cell(row=row, column=1, value="Entry Equity Value")
        su_equity = su.get('equity', None)
        if su_equity:
            ws.cell(row=row, column=2, value=f"='{SU}'!{su_equity}")
        else:
            # Hardcoded fallback: LTM EBITDA x Entry Multiple x ~55% equity
            entry_equity = data.get('entry_equity', a.ltm_ebitda * a.entry_multiple * 0.55)
            ws.cell(row=row, column=2, value=entry_equity)
        ws.cell(row=row, column=2).number_format = '#,##0.0'
        self._format_input_cell(ws, row, 2)

        # Row 5: blank spacer
        row = 6  # Attribution starts at row 6

        if has_full_upstream:
            # ========================================================
            # FORMULA-DRIVEN ATTRIBUTION (cross-sheet references)
            # ========================================================
            ebitda_row = om.get('ebitda_row', 13)
            ltm_ebitda_cell = asn.get('ltm_ebitda', 'B6')
            entry_mult_cell = asn.get('entry_multiple', 'B7')
            exit_mult_cell = asn.get('exit_multiple', 'B8')
            senior_debt_cell = su.get('senior_debt', 'B4')
            sub_debt_cell = su.get('sub_debt', 'B5')
            total_debt_row = ds.get('total_debt_row', None)

            attr_start = row

            # (+) EBITDA Growth = (Exit EBITDA - LTM EBITDA) x Exit Multiple
            ws.cell(row=row, column=1, value="(+) EBITDA Growth")
            ws.cell(row=row, column=2,
                    value=f"=('{OM}'!{exit_col}{ebitda_row}-'{A}'!{ltm_ebitda_cell})*'{A}'!{exit_mult_cell}")
            ws.cell(row=row, column=2).number_format = '#,##0.0'
            row += 1

            # (+) Multiple Expansion = Exit EBITDA x (Exit Mult - Entry Mult)
            ws.cell(row=row, column=1, value="(+) Multiple Expansion")
            ws.cell(row=row, column=2,
                    value=f"='{OM}'!{exit_col}{ebitda_row}*('{A}'!{exit_mult_cell}-'{A}'!{entry_mult_cell})")
            ws.cell(row=row, column=2).number_format = '#,##0.0'
            row += 1

            # (+) Debt Paydown = Entry Total Debt - Exit Total Debt
            ws.cell(row=row, column=1, value="(+) Debt Paydown")
            if total_debt_row:
                ws.cell(row=row, column=2,
                        value=f"=('{SU}'!{senior_debt_cell}+'{SU}'!{sub_debt_cell})-'{DS}'!{exit_col}{total_debt_row}")
            else:
                # Fallback: use senior ending + sub ending rows if total_debt_row unavailable
                sr_end = ds.get('senior_ending_row', 8)
                sub_end = ds.get('sub_ending_row', None)
                if sub_end:
                    ws.cell(row=row, column=2,
                            value=f"=('{SU}'!{senior_debt_cell}+'{SU}'!{sub_debt_cell})-('{DS}'!{exit_col}{sr_end}+'{DS}'!{exit_col}{sub_end})")
                else:
                    ws.cell(row=row, column=2,
                            value=f"=('{SU}'!{senior_debt_cell}+'{SU}'!{sub_debt_cell})-'{DS}'!{exit_col}{sr_end}")
            ws.cell(row=row, column=2).number_format = '#,##0.0'
            row += 1

            # (+) Other / Add-on M&A (user input or 0)
            ws.cell(row=row, column=1, value="(+) Other / Add-on M&A")
            other_value = data.get('other_value', 0)
            ws.cell(row=row, column=2, value=other_value)
            ws.cell(row=row, column=2).number_format = '#,##0.0'
            self._format_input_cell(ws, row, 2)
            attr_end = row
            row += 1

            # % of total attribution for each line
            for r in range(attr_start, attr_end + 1):
                ws.cell(row=r, column=3,
                        value=f"=IFERROR(B{r}/SUM(B${attr_start}:B${attr_end}),0)")
                ws.cell(row=r, column=3).number_format = '0%'

            # (=) Exit Equity = Entry Equity + SUM(attribution items)
            row += 1
            ws.cell(row=row, column=1, value="(=) Exit Equity")
            ws.cell(row=row, column=2, value=f"=B4+SUM(B{attr_start}:B{attr_end})")
            ws.cell(row=row, column=2).number_format = '#,##0.0'
            ws.cell(row=row, column=1).font = Font(bold=True)
            ws.cell(row=row, column=2).font = Font(bold=True)
            ws.cell(row=row, column=2).fill = PatternFill(
                start_color="90EE90", end_color="90EE90", fill_type="solid")
            exit_equity_row = row

            # MOIC = Exit Equity / Entry Equity
            row += 2
            ws.cell(row=row, column=1, value="MOIC")
            ws.cell(row=row, column=2, value=f"=IFERROR(B{exit_equity_row}/B4,0)")
            ws.cell(row=row, column=2).number_format = '0.00"x"'
            ws.cell(row=row, column=1).font = Font(bold=True)
            ws.cell(row=row, column=2).font = Font(bold=True)
            moic_row = row

            # IRR (approx) = MOIC^(1/hold_period) - 1
            row += 1
            hold_cell = asn.get('hold_period', 'B9')
            ws.cell(row=row, column=1, value="IRR (approx)")
            ws.cell(row=row, column=2,
                    value=f"=IFERROR(B{moic_row}^(1/'{A}'!{hold_cell})-1,0)")
            ws.cell(row=row, column=2).number_format = '0.0%'
            ws.cell(row=row, column=1).font = Font(bold=True)
            ws.cell(row=row, column=2).font = Font(bold=True)

            # cell_map for Value Creation module
            self.cell_map['value_creation'] = {
                'entry_equity_row': 4,
                'ebitda_growth_row': attr_start,
                'multiple_expansion_row': attr_start + 1,
                'debt_paydown_row': attr_start + 2,
                'other_row': attr_start + 3,
                'exit_equity_row': exit_equity_row,
                'moic_row': moic_row,
                'irr_row': moic_row + 1,
            }

        else:
            # ========================================================
            # HARDCODED FALLBACK (no upstream modules)
            # ========================================================
            entry_equity = data.get('entry_equity', a.ltm_ebitda * a.entry_multiple * 0.55)
            exit_equity = data.get('exit_equity', entry_equity * 2.5)
            value_created = exit_equity - entry_equity

            attribution = data.get('attribution', [
                {'name': 'Revenue Growth', 'value': value_created * 0.35},
                {'name': 'Margin Improvement', 'value': value_created * 0.25},
                {'name': 'Multiple Expansion', 'value': value_created * 0.15},
                {'name': 'Debt Paydown', 'value': value_created * 0.20},
                {'name': 'Add-on M&A', 'value': value_created * 0.05},
            ])
            attr_start = row
            attr_end = row + len(attribution) - 1
            for attr in attribution:
                ws.cell(row=row, column=1, value=attr['name'])
                ws.cell(row=row, column=2, value=attr['value'])
                ws.cell(row=row, column=2).number_format = '#,##0.0'
                self._format_input_cell(ws, row, 2)
                ws.cell(row=row, column=3, value=f"=IFERROR(B{row}/SUM(B${attr_start}:B${attr_end}),0)")
                ws.cell(row=row, column=3).number_format = '0%'
                row += 1

            row += 1
            ws.cell(row=row, column=1, value="Exit Equity Value")
            ws.cell(row=row, column=2, value=f"=B4+SUM(B{attr_start}:B{attr_end})")
            ws.cell(row=row, column=2).number_format = '#,##0.0'
            ws.cell(row=row, column=2).font = Font(bold=True)
            ws.cell(row=row, column=2).fill = PatternFill(
                start_color="90EE90", end_color="90EE90", fill_type="solid")

            # cell_map for fallback mode
            self.cell_map['value_creation'] = {
                'entry_equity_row': 4,
                'attr_start': attr_start,
                'attr_end': attr_end,
                'exit_equity_row': row,
            }

        ws.column_dimensions['A'].width = 25
        ws.column_dimensions['B'].width = 15
        ws.column_dimensions['C'].width = 10
        return self

    def add_hundred_day_plan(self, data: Dict = None) -> 'ValueCreationMixin':
        """Add 100-Day Plan Tracker sheet."""
        ws = self.wb.create_sheet("100-Day Plan")
        self.sheets_created.append("100-Day Plan")
        data = data or {}

        self._add_title(ws, f"{self.company_name} - 100-Day Plan", 1, 1)
        self._add_section_header(ws, "KEY INITIATIVES", 3, 1)

        headers = ["Initiative", "Owner", "Timeline", "Status", "Value ($M)"]
        row = 4
        for i, h in enumerate(headers):
            ws.cell(row=row, column=1+i, value=h)
        self._format_header_row(ws, row, 1, 5)
        row += 1

        initiatives = data.get('initiatives', [
            {'name': 'Management assessment', 'owner': 'CEO', 'timeline': 'Days 1-30', 'status': 'Complete', 'value': 0},
            {'name': 'Procurement RFP', 'owner': 'CFO', 'timeline': 'Days 30-90', 'status': 'In Progress', 'value': 3.0},
            {'name': 'Sales effectiveness', 'owner': 'CRO', 'timeline': 'Days 30-100', 'status': 'Planning', 'value': 5.0},
            {'name': 'Working capital optimization', 'owner': 'CFO', 'timeline': 'Days 15-75', 'status': 'In Progress', 'value': 1.5},
        ])
        init_start = row
        for init in initiatives:
            ws.cell(row=row, column=1, value=init['name'])
            ws.cell(row=row, column=2, value=init['owner'])
            ws.cell(row=row, column=3, value=init['timeline'])
            ws.cell(row=row, column=4, value=init['status'])
            ws.cell(row=row, column=5, value=init['value'] if init['value'] > 0 else "—")
            status_colors = {'Complete': "90EE90", 'In Progress': "FFFF99", 'Planning': "FFE4B5"}
            if init['status'] in status_colors:
                ws.cell(row=row, column=4).fill = PatternFill(start_color=status_colors[init['status']], end_color=status_colors[init['status']], fill_type="solid")
            if init['value'] > 0: ws.cell(row=row, column=5).number_format = '#,##0.0'
            row += 1
        init_end = row - 1

        row += 1
        ws.cell(row=row, column=1, value="Total Value at Stake")
        ws.cell(row=row, column=5, value=f"=SUM(E{init_start}:E{init_end})")
        ws.cell(row=row, column=5).number_format = '#,##0.0'
        ws.cell(row=row, column=5).font = Font(bold=True)
        ws.cell(row=row, column=5).fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")

        # cell_map for 100-Day Plan
        self.cell_map['hundred_day_plan'] = {
            'init_start': init_start,
            'init_end': init_end,
            'total_row': row,
        }

        ws.column_dimensions['A'].width = 30
        for c in ['B','C','D','E']: ws.column_dimensions[c].width = 12
        return self

    def add_exit_readiness(self, data: Dict = None) -> 'ValueCreationMixin':
        """Add Exit Readiness Assessment sheet."""
        ws = self.wb.create_sheet("Exit Readiness")
        self.sheets_created.append("Exit Readiness")
        data = data or {}

        self._add_title(ws, f"{self.company_name} - Exit Readiness", 1, 1)
        self._add_section_header(ws, "READINESS SCORECARD", 3, 1)

        headers = ["Category", "Score", "Weight", "Weighted"]
        row = 4
        for i, h in enumerate(headers):
            ws.cell(row=row, column=1+i, value=h)
        self._format_header_row(ws, row, 1, 4)
        row += 1

        categories = data.get('categories', [
            {'name': 'Financial Performance', 'score': 4, 'weight': 0.25},
            {'name': 'Management Team', 'score': 4, 'weight': 0.20},
            {'name': 'Growth Profile', 'score': 4, 'weight': 0.20},
            {'name': 'Market Position', 'score': 3, 'weight': 0.15},
            {'name': 'Financial Reporting', 'score': 3, 'weight': 0.10},
            {'name': 'Legal/Compliance', 'score': 4, 'weight': 0.10},
        ])
        cat_start = row
        for cat in categories:
            ws.cell(row=row, column=1, value=cat['name'])
            ws.cell(row=row, column=2, value=cat['score'])
            self._format_input_cell(ws, row, 2)
            ws.cell(row=row, column=3, value=cat['weight'])
            ws.cell(row=row, column=3).number_format = '0%'
            self._format_input_cell(ws, row, 3)
            ws.cell(row=row, column=4, value=f"=B{row}*C{row}")
            ws.cell(row=row, column=4).number_format = '0.0'
            color = "90EE90" if cat['score'] >= 4 else "FFFF99" if cat['score'] >= 3 else "FFB6C1"
            ws.cell(row=row, column=2).fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
            row += 1
        cat_end = row - 1

        row += 1
        ws.cell(row=row, column=1, value="Overall Score")
        # Use SUMPRODUCT for a cleaner weighted sum (equivalent to SUM of D column)
        ws.cell(row=row, column=4,
                value=f"=SUMPRODUCT(B{cat_start}:B{cat_end},C{cat_start}:C{cat_end})")
        ws.cell(row=row, column=4).number_format = '0.0'
        ws.cell(row=row, column=4).font = Font(bold=True)
        ws.cell(row=row, column=4).fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")

        # cell_map for Exit Readiness
        self.cell_map['exit_readiness'] = {
            'cat_start': cat_start,
            'cat_end': cat_end,
            'overall_score_row': row,
        }

        ws.column_dimensions['A'].width = 25
        for c in ['B','C','D']: ws.column_dimensions[c].width = 12
        return self

    def add_management_incentive_plan(self, data: Dict = None) -> 'ValueCreationMixin':
        """Add Management Incentive Plan (MIP) sheet."""
        ws = self.wb.create_sheet("Management MIP")
        self.sheets_created.append("Management MIP")
        a = self.assumptions
        data = data or {}

        self._add_title(ws, f"{self.company_name} - Management Incentive Plan", 1, 1)

        mip_pool = data.get('mip_pool_pct', 0.10)

        self._add_section_header(ws, "MIP STRUCTURE", 3, 1)
        ws.cell(row=4, column=1, value="MIP Pool (% of equity)")
        ws.cell(row=4, column=2, value=mip_pool)
        ws.cell(row=4, column=2).number_format = '0.0%'
        self._format_input_cell(ws, 4, 2)

        # Entry Equity from Sources & Uses (fixed key: 'equity' not 'sponsor_equity')
        ws.cell(row=5, column=1, value="Entry Equity ($M)")
        su = self.cell_map.get('sources_uses', {})
        su_equity = su.get('equity', None)
        if su_equity:
            ws.cell(row=5, column=2, value=f"='Sources & Uses'!{su_equity}")
        else:
            ws.cell(row=5, column=2, value="='Assumptions'!B6*'Assumptions'!B7*0.55")
        ws.cell(row=5, column=2).number_format = '#,##0.0'

        ws.cell(row=6, column=1, value="MIP Pool Value ($M)")
        ws.cell(row=6, column=2, value="=B5*B4")
        ws.cell(row=6, column=2).number_format = '#,##0.0'

        self._add_section_header(ws, "PAYOUT BY EXIT MOIC", 8, 1)
        headers = ["Exit MOIC", "Exit Equity", "MIP Value", "CEO (35%)"]
        for i, h in enumerate(headers):
            ws.cell(row=9, column=1+i, value=h)
        self._format_header_row(ws, 9, 1, 4)

        row = 10
        for moic in [1.5, 2.0, 2.5, 3.0]:
            ws.cell(row=row, column=1, value=moic)
            ws.cell(row=row, column=1).number_format = '0.0x'
            ws.cell(row=row, column=2, value=f"=$B$5*A{row}")     # Exit Equity
            ws.cell(row=row, column=3, value=f"=B{row}*$B$4")     # MIP Value
            ws.cell(row=row, column=4, value=f"=C{row}*0.35")     # CEO share
            for c in [2,3,4]: ws.cell(row=row, column=c).number_format = '#,##0.0'
            row += 1

        # cell_map for MIP module
        self.cell_map['management_mip'] = {
            'mip_pool_pct': 'B4',
            'entry_equity': 'B5',
            'mip_pool_value': 'B6',
            'payout_start_row': 10,
            'payout_end_row': row - 1,
        }

        ws.column_dimensions['A'].width = 20
        for c in ['B','C','D']: ws.column_dimensions[c].width = 15
        return self
