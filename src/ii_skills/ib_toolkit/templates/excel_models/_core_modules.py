#!/usr/bin/env python3
"""
Core Excel Modules Mixin.

Contains the 14 smart (formula-driven) modules that form the backbone
of every LBO model: Assumptions, Sources & Uses, Operating Model,
Revenue Build, Expense Build, Debt Schedule, WACC, Working Capital,
Returns Analysis, Sensitivity, Scenario Analysis, Management vs Buyer,
DCF Valuation, and Covenant Analysis.
"""

from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from .formula_builder import FormulaBuilder as FB
from ._base import (INPUT_FILL, INPUT_FONT, HEADER_FILL, HEADER_FONT,
                     BOTTOM_BORDER, DOUBLE_BORDER, ModelDepth)
from typing import Dict


class CoreModulesMixin:
    """Mixin providing the 14 core smart formula modules."""

    # ============================================================
    # MODULE: ASSUMPTIONS (Central Input Sheet)
    # ============================================================

    def add_assumptions_sheet(self) -> 'CoreModulesMixin':
        """Add central Assumptions sheet — all model inputs in one place.
        Every other formula-driven sheet references this sheet."""
        ws = self.wb.create_sheet("Assumptions")
        self.sheets_created.append("Assumptions")

        a = self.assumptions

        self._add_title(ws, f"{self.company_name} - Model Assumptions", 1, 1)

        # --- Company / Transaction ---
        self._add_section_header(ws, "Transaction Assumptions", 3, 1)
        self._format_header_row(ws, 3, 1, 2)

        inputs = [
            (5,  "LTM Revenue ($M)",        a.ltm_revenue,            '#,##0.0'),
            (6,  "LTM EBITDA ($M)",         a.ltm_ebitda,             '#,##0.0'),
            (7,  "Entry Multiple",          a.entry_multiple,         '0.0"x"'),
            (8,  "Exit Multiple",           a.exit_multiple,          '0.0"x"'),
            (9,  "Hold Period (Years)",     a.hold_period,            '#,##0'),
        ]

        for row, label, value, fmt in inputs:
            ws.cell(row=row, column=1, value=label)
            cell = ws.cell(row=row, column=2, value=value)
            cell.number_format = fmt
            self._format_input_cell(ws, row, 2)

        # --- Debt Structure ---
        self._add_section_header(ws, "Debt Structure", 11, 1)

        debt_inputs = [
            (12, "Senior Debt Multiple",    a.senior_debt_multiple,   '0.0"x"'),
            (13, "Senior Interest Rate",    a.senior_interest_rate,   '0.0%'),
            (14, "Senior Amortization %",   a.senior_amortization,    '0.0%'),
            (15, "Sub Debt Multiple",       a.sub_debt_multiple,      '0.0"x"'),
            (16, "Sub Interest Rate",       a.sub_interest_rate,      '0.0%'),
        ]

        for row, label, value, fmt in debt_inputs:
            ws.cell(row=row, column=1, value=label)
            cell = ws.cell(row=row, column=2, value=value)
            cell.number_format = fmt
            self._format_input_cell(ws, row, 2)

        # --- Fees & Rates ---
        self._add_section_header(ws, "Fees & Rates", 18, 1)

        fee_inputs = [
            (19, "Transaction Fee %",       a.transaction_fee_pct,    '0.0%'),
            (20, "Financing Fee %",         a.financing_fee_pct,      '0.0%'),
            (22, "CapEx % Revenue",         a.capex_pct_revenue,      '0.0%'),
            (23, "NWC % Revenue",           a.nwc_pct_revenue,        '0.0%'),
            (24, "Tax Rate",                a.tax_rate,               '0.0%'),
        ]

        for row, label, value, fmt in fee_inputs:
            ws.cell(row=row, column=1, value=label)
            cell = ws.cell(row=row, column=2, value=value)
            cell.number_format = fmt
            self._format_input_cell(ws, row, 2)

        # --- Projection Drivers (Year 1–5 across columns C–G) ---
        self._add_section_header(ws, "Projection Drivers", 26, 1)

        # Column headers
        for i in range(self.projection_years):
            ws.cell(row=27, column=3 + i, value=f"Year {i + 1}")
            ws.cell(row=27, column=3 + i).font = Font(bold=True)
            ws.cell(row=27, column=3 + i).alignment = Alignment(horizontal='center')

        # Revenue Growth (row 28, cols C–G)
        ws.cell(row=28, column=1, value="Revenue Growth %")
        for i, g in enumerate(a.revenue_growth[:self.projection_years]):
            cell = ws.cell(row=28, column=3 + i, value=g)
            cell.number_format = '0.0%'
            self._format_input_cell(ws, 28, 3 + i)

        # EBITDA Margin (row 29, cols C–G)
        ws.cell(row=29, column=1, value="EBITDA Margin %")
        for i, m in enumerate(a.ebitda_margin[:self.projection_years]):
            cell = ws.cell(row=29, column=3 + i, value=m)
            cell.number_format = '0.0%'
            self._format_input_cell(ws, 29, 3 + i)

        # Column widths
        ws.column_dimensions['A'].width = 25
        ws.column_dimensions['B'].width = 15
        for i in range(self.projection_years):
            ws.column_dimensions[get_column_letter(3 + i)].width = 12

        # Store cell locations for cross-references
        self.cell_map['assumptions'] = {
            'ltm_revenue': 'B5',
            'ltm_ebitda': 'B6',
            'entry_multiple': 'B7',
            'exit_multiple': 'B8',
            'hold_period': 'B9',
            'senior_debt_mult': 'B12',
            'senior_rate': 'B13',
            'senior_amort': 'B14',
            'sub_debt_mult': 'B15',
            'sub_rate': 'B16',
            'txn_fee_pct': 'B19',
            'fin_fee_pct': 'B20',
            'capex_pct': 'B22',
            'nwc_pct': 'B23',
            'tax_rate': 'B24',
            'rev_growth_row': 28,       # row number, cols C onward
            'ebitda_margin_row': 29,    # row number, cols C onward
        }

        return self

    # ============================================================
    # MODULE: SOURCES & USES
    # ============================================================

    def add_sources_uses(self, data: Dict = None) -> 'CoreModulesMixin':
        """Add Sources & Uses sheet with real Excel formulas referencing Assumptions."""
        ws = self.wb.create_sheet("Sources & Uses")
        self.sheets_created.append("Sources & Uses")

        A = "Assumptions"  # sheet name for cross-refs

        # Title
        self._add_title(ws, f"{self.company_name} - Sources & Uses", 1, 1)

        # Sources section
        self._add_section_header(ws, "Sources", 3, 1)
        ws.cell(row=3, column=2, value="$M")
        ws.cell(row=3, column=3, value="% of Total")
        self._format_header_row(ws, 3, 1, 3)

        # Row 4: Senior Debt = LTM EBITDA x Senior Debt Multiple
        ws.cell(row=4, column=1, value="Senior Secured Debt")
        ws.cell(row=4, column=2, value=f"='{A}'!B6*'{A}'!B12")
        ws.cell(row=4, column=2).number_format = '#,##0.0'
        ws.cell(row=4, column=3, value="=B4/$B$7")
        ws.cell(row=4, column=3).number_format = '0.0%'
        self._format_input_cell(ws, 4, 2)

        # Row 5: Sub Debt = LTM EBITDA x Sub Debt Multiple
        ws.cell(row=5, column=1, value="Subordinated Debt")
        ws.cell(row=5, column=2, value=f"='{A}'!B6*'{A}'!B15")
        ws.cell(row=5, column=2).number_format = '#,##0.0'
        ws.cell(row=5, column=3, value="=B5/$B$7")
        ws.cell(row=5, column=3).number_format = '0.0%'
        self._format_input_cell(ws, 5, 2)

        # Row 6: Sponsor Equity = Total Uses - Senior - Sub
        ws.cell(row=6, column=1, value="Sponsor Equity")
        ws.cell(row=6, column=2, value="=B12-B4-B5")
        ws.cell(row=6, column=2).number_format = '#,##0.0'
        ws.cell(row=6, column=3, value="=B6/$B$7")
        ws.cell(row=6, column=3).number_format = '0.0%'

        # Row 7: Total Sources
        ws.cell(row=7, column=1, value="Total Sources")
        ws.cell(row=7, column=2, value="=SUM(B4:B6)")
        ws.cell(row=7, column=2).number_format = '#,##0.0'
        ws.cell(row=7, column=3, value="=1")
        ws.cell(row=7, column=3).number_format = '0.0%'
        for col in range(1, 4):
            ws.cell(row=7, column=col).font = Font(bold=True)
            ws.cell(row=7, column=col).border = DOUBLE_BORDER

        # Uses section (row 9)
        self._add_section_header(ws, "Uses", 9, 1)
        ws.cell(row=9, column=2, value="$M")
        ws.cell(row=9, column=3, value="% of Total")
        self._format_header_row(ws, 9, 1, 3)

        # Row 10: Purchase EV = LTM EBITDA x Entry Multiple
        ws.cell(row=10, column=1, value="Purchase Enterprise Value")
        ws.cell(row=10, column=2, value=f"='{A}'!B6*'{A}'!B7")
        ws.cell(row=10, column=2).number_format = '#,##0.0'
        ws.cell(row=10, column=3, value="=B10/$B$12")
        ws.cell(row=10, column=3).number_format = '0.0%'

        # Row 11: Transaction Fees = EV x Fee %
        ws.cell(row=11, column=1, value="Transaction Fees")
        ws.cell(row=11, column=2, value=f"=B10*'{A}'!B19")
        ws.cell(row=11, column=2).number_format = '#,##0.0'
        ws.cell(row=11, column=3, value="=B11/$B$12")
        ws.cell(row=11, column=3).number_format = '0.0%'

        # Row 12 (was Financing Fees, move to 12 for Fin Fees): Financing Fees = Total Debt x Fin Fee %
        # Actually let's put Fin Fees at 12 and Total at 13
        ws.cell(row=12, column=1, value="Financing Fees")
        ws.cell(row=12, column=2, value=f"=(B4+B5)*'{A}'!B20")
        ws.cell(row=12, column=2).number_format = '#,##0.0'
        ws.cell(row=12, column=3, value="=B12/$B$13")
        ws.cell(row=12, column=3).number_format = '0.0%'

        # Fix % of total references to point to row 13 (total uses)
        ws.cell(row=10, column=3, value="=B10/$B$13")
        ws.cell(row=11, column=3, value="=B11/$B$13")
        ws.cell(row=12, column=3, value="=B12/$B$13")

        # Row 13: Total Uses
        ws.cell(row=13, column=1, value="Total Uses")
        ws.cell(row=13, column=2, value="=SUM(B10:B12)")
        ws.cell(row=13, column=2).number_format = '#,##0.0'
        ws.cell(row=13, column=3, value="=1")
        ws.cell(row=13, column=3).number_format = '0.0%'
        for col in range(1, 4):
            ws.cell(row=13, column=col).font = Font(bold=True)
            ws.cell(row=13, column=col).border = DOUBLE_BORDER

        # Key metrics (row 15)
        self._add_section_header(ws, "Key Metrics", 15, 1)

        ws.cell(row=16, column=1, value="Entry EV / EBITDA")
        ws.cell(row=16, column=2, value=f"='{A}'!B7")
        ws.cell(row=16, column=2).number_format = '0.0"x"'

        ws.cell(row=17, column=1, value="Total Debt / EBITDA")
        ws.cell(row=17, column=2, value=f"=(B4+B5)/'{A}'!B6")
        ws.cell(row=17, column=2).number_format = '0.0"x"'

        ws.cell(row=18, column=1, value="Senior Debt / EBITDA")
        ws.cell(row=18, column=2, value=f"=B4/'{A}'!B6")
        ws.cell(row=18, column=2).number_format = '0.0"x"'

        ws.cell(row=19, column=1, value="Equity Check ($M)")
        ws.cell(row=19, column=2, value="=B6")
        ws.cell(row=19, column=2).number_format = '#,##0.0'

        ws.cell(row=20, column=1, value="Equity / EV")
        ws.cell(row=20, column=2, value="=B6/B10")
        ws.cell(row=20, column=2).number_format = '0.0%'

        # Set column widths
        ws.column_dimensions['A'].width = 25
        ws.column_dimensions['B'].width = 15
        ws.column_dimensions['C'].width = 12

        # Store cell map
        self.cell_map['sources_uses'] = {
            'senior_debt': 'B4',
            'sub_debt': 'B5',
            'equity': 'B6',
            'total_sources': 'B7',
            'purchase_ev': 'B10',
            'total_uses': 'B13',
        }

        return self

    # ============================================================
    # MODULE: OPERATING MODEL
    # ============================================================

    def add_operating_model(self, data: Dict = None) -> 'CoreModulesMixin':
        """Add Operating Model sheet with real Excel formulas referencing Assumptions."""
        ws = self.wb.create_sheet("Operating Model")
        self.sheets_created.append("Operating Model")

        A = "Assumptions"
        n = self.projection_years  # typically 5
        # Column layout: B=LTM, C=Year1, D=Year2, ...
        ltm_col = 2        # column B
        y1_col = 3          # column C

        # Title
        self._add_title(ws, f"{self.company_name} - Operating Model", 1, 1)

        # Year headers
        self._year_headers(ws, 3, 2)

        # --- Assumptions section (rows 4-7) ---
        self._add_section_header(ws, "Assumptions", 4, 1)

        # Row 5: Revenue Growth % — LTM is "—", Y1+ reference Assumptions row 28
        ws.cell(row=5, column=1, value="Revenue Growth %")
        ws.cell(row=5, column=ltm_col, value="—")
        for i in range(n):
            col = y1_col + i
            # Reference Assumptions!C28, D28, etc.
            ws.cell(row=5, column=col, value=f"='{A}'!{get_column_letter(3 + i)}28")
            ws.cell(row=5, column=col).number_format = '0.0%'
            self._format_input_cell(ws, 5, col)

        # Row 6: EBITDA Margin % — LTM = EBITDA/Revenue, Y1+ reference Assumptions row 29
        ws.cell(row=6, column=1, value="EBITDA Margin %")
        ws.cell(row=6, column=ltm_col, value=f"='{A}'!B6/'{A}'!B5")
        ws.cell(row=6, column=ltm_col).number_format = '0.0%'
        for i in range(n):
            col = y1_col + i
            ws.cell(row=6, column=col, value=f"='{A}'!{get_column_letter(3 + i)}29")
            ws.cell(row=6, column=col).number_format = '0.0%'
            self._format_input_cell(ws, 6, col)

        # Row 7: CapEx % Revenue — all periods reference Assumptions!B22
        ws.cell(row=7, column=1, value="CapEx % Revenue")
        for i in range(n + 1):
            col = ltm_col + i
            ws.cell(row=7, column=col, value=f"='{A}'!B22")
            ws.cell(row=7, column=col).number_format = '0.0%'
            self._format_input_cell(ws, 7, col)

        # --- Operating Model section (rows 9+) ---
        self._add_section_header(ws, "Operating Model ($M)", 9, 1)

        # Row 10: Revenue — LTM from Assumptions, Y1+ = prior x (1 + growth)
        ws.cell(row=10, column=1, value="Revenue")
        ws.cell(row=10, column=ltm_col, value=f"='{A}'!B5")
        ws.cell(row=10, column=ltm_col).number_format = '#,##0.0'
        for i in range(n):
            col = y1_col + i
            prev = get_column_letter(col - 1)
            this = get_column_letter(col)
            ws.cell(row=10, column=col, value=f"={prev}10*(1+{this}5)")
            ws.cell(row=10, column=col).number_format = '#,##0.0'

        # Row 11: % Growth — LTM is "—", Y1+ = (this/prior)-1
        ws.cell(row=11, column=1, value="  % Growth")
        ws.cell(row=11, column=ltm_col, value="—")
        for i in range(n):
            col = y1_col + i
            prev = get_column_letter(col - 1)
            this = get_column_letter(col)
            ws.cell(row=11, column=col, value=f"={this}10/{prev}10-1")
            ws.cell(row=11, column=col).number_format = '0.0%'
            ws.cell(row=11, column=col).font = Font(italic=True, color="666666")

        # Row 13: EBITDA — LTM from Assumptions, Y1+ = Revenue x Margin
        ws.cell(row=13, column=1, value="EBITDA")
        ws.cell(row=13, column=ltm_col, value=f"='{A}'!B6")
        ws.cell(row=13, column=ltm_col).number_format = '#,##0.0'
        for i in range(n):
            col = y1_col + i
            this = get_column_letter(col)
            ws.cell(row=13, column=col, value=f"={this}10*{this}6")
            ws.cell(row=13, column=col).number_format = '#,##0.0'

        # Row 14: % Margin = EBITDA / Revenue
        ws.cell(row=14, column=1, value="  % Margin")
        for i in range(n + 1):
            col = ltm_col + i
            this = get_column_letter(col)
            ws.cell(row=14, column=col, value=f"={this}13/{this}10")
            ws.cell(row=14, column=col).number_format = '0.0%'
            ws.cell(row=14, column=col).font = Font(italic=True, color="666666")

        # Row 16: CapEx = -Revenue x CapEx%
        ws.cell(row=16, column=1, value="CapEx")
        for i in range(n + 1):
            col = ltm_col + i
            this = get_column_letter(col)
            ws.cell(row=16, column=col, value=f"=-{this}10*{this}7")
            ws.cell(row=16, column=col).number_format = '(#,##0.0)'

        # Row 17: D&A = -CapEx (same magnitude, positive)
        ws.cell(row=17, column=1, value="D&A")
        for i in range(n + 1):
            col = ltm_col + i
            this = get_column_letter(col)
            ws.cell(row=17, column=col, value=f"=-{this}16")
            ws.cell(row=17, column=col).number_format = '#,##0.0'

        # Row 19: EBITDA - CapEx
        ws.cell(row=19, column=1, value="EBITDA - CapEx")
        for i in range(n + 1):
            col = ltm_col + i
            this = get_column_letter(col)
            ws.cell(row=19, column=col, value=f"={this}13+{this}16")
            ws.cell(row=19, column=col).number_format = '#,##0.0'
            ws.cell(row=19, column=col).font = Font(bold=True)

        # Set column widths
        ws.column_dimensions['A'].width = 20
        for i in range(n + 2):
            ws.column_dimensions[get_column_letter(2 + i)].width = 12

        # Store cell map
        self.cell_map['operating_model'] = {
            'revenue_row': 10,
            'ebitda_row': 13,
            'capex_row': 16,
            'ebitda_minus_capex_row': 19,
            'ltm_col': ltm_col,
            'y1_col': y1_col,
        }

        return self

    # ============================================================
    # MODULE: REVENUE BUILD
    # ============================================================

    def add_revenue_build(self, data: Dict = None) -> 'CoreModulesMixin':
        """Add detailed Revenue Build sheet with real Excel formulas."""
        ws = self.wb.create_sheet("Revenue Build")
        self.sheets_created.append("Revenue Build")

        A = "Assumptions"
        a = self.assumptions
        n = self.projection_years

        # Title
        self._add_title(ws, f"{self.company_name} - Revenue Build", 1, 1)

        # Year headers
        self._year_headers(ws, 3, 2)

        if self.depth == ModelDepth.QUICK:
            # Simple: total revenue with formulas referencing Assumptions
            self._add_section_header(ws, "Revenue Projections ($M)", 4, 1)

            # Row 5: Total Revenue — LTM from Assumptions, projections via growth formula
            ws.cell(row=5, column=1, value="Total Revenue")
            ws.cell(row=5, column=2, value=f"='{A}'!B5")  # LTM Revenue
            ws.cell(row=5, column=2).number_format = '#,##0.0'
            for i in range(n):
                prev_col = get_column_letter(2 + i)
                growth_col = get_column_letter(3 + i)
                ws.cell(row=5, column=3 + i,
                        value=f"={prev_col}5*(1+'{A}'!{growth_col}28)")
                ws.cell(row=5, column=3 + i).number_format = '#,##0.0'

            # Row 6: Growth % — references from Assumptions row 28
            ws.cell(row=6, column=1, value="  % Growth")
            ws.cell(row=6, column=2, value="—")
            for i in range(n):
                growth_col = get_column_letter(3 + i)
                ws.cell(row=6, column=3 + i, value=f"='{A}'!{growth_col}28")
                ws.cell(row=6, column=3 + i).number_format = '0.0%'
                self._format_input_cell(ws, 6, 3 + i)

            self.cell_map['revenue_build'] = {
                'total_revenue_row': 5,
                'growth_row': 6,
                'ltm_col': 2,
                'y1_col': 3,
            }

        else:
            # Standard/Comprehensive: Segment breakdown with formulas
            data = data or {}
            segments = data.get('segments', [
                {"name": "Segment A", "pct": 0.50, "growth": [0.10, 0.09, 0.08, 0.07, 0.06]},
                {"name": "Segment B", "pct": 0.30, "growth": [0.06, 0.05, 0.05, 0.04, 0.04]},
                {"name": "Segment C", "pct": 0.20, "growth": [0.04, 0.03, 0.03, 0.02, 0.02]},
            ])

            # Segment assumptions (input cells — stay as values)
            self._add_section_header(ws, "Segment Assumptions", 4, 1)
            ws.cell(row=5, column=1, value="Segment")
            ws.cell(row=5, column=2, value="LTM %")
            for i in range(n):
                ws.cell(row=5, column=3 + i, value=f"Y{i+1} Growth")
            self._format_header_row(ws, 5, 1, 2 + n)

            for j, seg in enumerate(segments):
                row = 6 + j
                ws.cell(row=row, column=1, value=seg["name"])
                ws.cell(row=row, column=2, value=seg["pct"])
                ws.cell(row=row, column=2).number_format = '0.0%'
                self._format_input_cell(ws, row, 2)

                for i, g in enumerate(seg["growth"][:n]):
                    cell = ws.cell(row=row, column=3 + i, value=g)
                    cell.number_format = '0.0%'
                    self._format_input_cell(ws, row, 3 + i)

            # Revenue by segment — formulas referencing Assumptions + segment inputs
            seg_start = 6 + len(segments) + 1
            self._add_section_header(ws, "Revenue by Segment ($M)", seg_start, 1)

            for j, seg in enumerate(segments):
                row = seg_start + 1 + j
                seg_input_row = 6 + j  # Row with this segment's pct and growth
                ws.cell(row=row, column=1, value=seg["name"])

                # LTM: =Assumptions!B5 * segment_pct
                pct_cell = f"B{seg_input_row}"
                ws.cell(row=row, column=2, value=f"='{A}'!B5*{pct_cell}")
                ws.cell(row=row, column=2).number_format = '#,##0.0'

                # Projected: =prev_revenue * (1 + segment_growth)
                for i in range(n):
                    prev_col = get_column_letter(2 + i)
                    growth_col = get_column_letter(3 + i)
                    ws.cell(row=row, column=3 + i,
                            value=f"={prev_col}{row}*(1+{growth_col}{seg_input_row})")
                    ws.cell(row=row, column=3 + i).number_format = '#,##0.0'

            # Total revenue — SUM formula
            total_row = seg_start + 1 + len(segments)
            ws.cell(row=total_row, column=1, value="Total Revenue")

            for i in range(n + 1):
                start_row = seg_start + 1
                end_row = total_row - 1
                col_letter = get_column_letter(2 + i)
                ws.cell(row=total_row, column=2 + i,
                       value=f"=SUM({col_letter}{start_row}:{col_letter}{end_row})")
                ws.cell(row=total_row, column=2 + i).number_format = '#,##0.0'
                ws.cell(row=total_row, column=2 + i).font = Font(bold=True)
                ws.cell(row=total_row, column=2 + i).border = DOUBLE_BORDER

            self.cell_map['revenue_build'] = {
                'total_revenue_row': total_row,
                'seg_start_row': seg_start + 1,
                'seg_count': len(segments),
                'ltm_col': 2,
                'y1_col': 3,
            }

        # Set column widths
        ws.column_dimensions['A'].width = 20
        for i in range(n + 2):
            ws.column_dimensions[get_column_letter(2 + i)].width = 12

        return self

    # ============================================================
    # MODULE: EXPENSE BUILD (SG&A BREAKOUT)
    # ============================================================

    def add_expense_build(self, data: Dict = None) -> 'CoreModulesMixin':
        """Add Expense / SG&A Build sheet with real Excel formulas."""
        ws = self.wb.create_sheet("Expense Build")
        self.sheets_created.append("Expense Build")

        A = "Assumptions"
        OM = "Operating Model"
        a = self.assumptions
        n = self.projection_years

        # Title
        self._add_title(ws, f"{self.company_name} - Expense Build", 1, 1)

        # Year headers
        self._year_headers(ws, 3, 2)

        # Revenue reference row — pull from Operating Model row 10
        om_rev_row = self.cell_map.get('operating_model', {}).get('revenue_row', 10)

        if self.depth == ModelDepth.QUICK:
            # Simple: COGS and SG&A as % of revenue (input cells stay as values)
            self._add_section_header(ws, "Cost Assumptions (% Revenue)", 4, 1)

            cogs_pct = 0.60
            sga_pct = 0.20

            ws.cell(row=5, column=1, value="COGS %")
            for i in range(n + 1):
                cell = ws.cell(row=5, column=2 + i, value=cogs_pct)
                cell.number_format = '0.0%'
                self._format_input_cell(ws, 5, 2 + i)

            ws.cell(row=6, column=1, value="SG&A %")
            for i in range(n + 1):
                cell = ws.cell(row=6, column=2 + i, value=sga_pct)
                cell.number_format = '0.0%'
                self._format_input_cell(ws, 6, 2 + i)

            # Calculated values — formulas referencing OM revenue x local %
            self._add_section_header(ws, "Expense Summary ($M)", 8, 1)

            ws.cell(row=9, column=1, value="COGS")
            for i in range(n + 1):
                col = get_column_letter(2 + i)
                ws.cell(row=9, column=2 + i,
                        value=f"='{OM}'!{col}{om_rev_row}*{col}5")
                ws.cell(row=9, column=2 + i).number_format = '#,##0.0'

            ws.cell(row=10, column=1, value="Gross Profit")
            for i in range(n + 1):
                col = get_column_letter(2 + i)
                ws.cell(row=10, column=2 + i,
                        value=f"='{OM}'!{col}{om_rev_row}-{col}9")
                ws.cell(row=10, column=2 + i).number_format = '#,##0.0'
                ws.cell(row=10, column=2 + i).font = Font(bold=True)

            ws.cell(row=12, column=1, value="SG&A")
            for i in range(n + 1):
                col = get_column_letter(2 + i)
                ws.cell(row=12, column=2 + i,
                        value=f"='{OM}'!{col}{om_rev_row}*{col}6")
                ws.cell(row=12, column=2 + i).number_format = '#,##0.0'

            self.cell_map['expense_build'] = {
                'cogs_pct_row': 5,
                'sga_pct_row': 6,
                'cogs_row': 9,
                'gross_profit_row': 10,
                'sga_row': 12,
            }

        else:
            # Standard/Comprehensive: Detailed breakdown with formulas
            data = data or {}
            self._add_section_header(ws, "COGS Breakdown ($M)", 4, 1)

            cogs_items = data.get('cogs_items', [
                {"name": "Materials", "pct": 0.35},
                {"name": "Direct Labor", "pct": 0.15},
                {"name": "Manufacturing OH", "pct": 0.10},
            ])

            # COGS line items: formula = OM Revenue x item_pct (input in adjacent row)
            # First write pct input rows, then formula rows
            # Layout: Row 5+ = item name + pct input | formula = OM_rev * pct
            for j, item in enumerate(cogs_items):
                row = 5 + j
                ws.cell(row=row, column=1, value=item["name"])
                for i in range(n + 1):
                    col = get_column_letter(2 + i)
                    # Write pct as input value (these are the driver assumptions)
                    ws.cell(row=row, column=2 + i, value=item["pct"])
                    ws.cell(row=row, column=2 + i).number_format = '0.0%'
                    self._format_input_cell(ws, row, 2 + i)

            # COGS $ rows (formulas)
            cogs_dollar_start = 5 + len(cogs_items) + 1
            self._add_section_header(ws, "COGS ($M)", cogs_dollar_start - 1, 1)

            for j, item in enumerate(cogs_items):
                row = cogs_dollar_start + j
                pct_row = 5 + j
                ws.cell(row=row, column=1, value=f"{item['name']} $")
                for i in range(n + 1):
                    col = get_column_letter(2 + i)
                    ws.cell(row=row, column=2 + i,
                            value=f"='{OM}'!{col}{om_rev_row}*{col}{pct_row}")
                    ws.cell(row=row, column=2 + i).number_format = '#,##0.0'

            cogs_total_row = cogs_dollar_start + len(cogs_items)
            ws.cell(row=cogs_total_row, column=1, value="Total COGS")
            for i in range(n + 1):
                col = get_column_letter(2 + i)
                ws.cell(row=cogs_total_row, column=2 + i,
                        value=f"=SUM({col}{cogs_dollar_start}:{col}{cogs_total_row - 1})")
                ws.cell(row=cogs_total_row, column=2 + i).number_format = '#,##0.0'
                ws.cell(row=cogs_total_row, column=2 + i).font = Font(bold=True)
                ws.cell(row=cogs_total_row, column=2 + i).border = BOTTOM_BORDER

            # SG&A breakdown
            sga_pct_start = cogs_total_row + 2
            self._add_section_header(ws, "SG&A Assumptions (% Revenue)", sga_pct_start, 1)

            sga_items = data.get('sga_items', [
                {"name": "Sales & Marketing", "pct": 0.08},
                {"name": "General & Admin", "pct": 0.05},
                {"name": "R&D", "pct": 0.04},
                {"name": "Other", "pct": 0.03},
            ])

            for j, item in enumerate(sga_items):
                row = sga_pct_start + 1 + j
                ws.cell(row=row, column=1, value=item["name"])
                for i in range(n + 1):
                    ws.cell(row=row, column=2 + i, value=item["pct"])
                    ws.cell(row=row, column=2 + i).number_format = '0.0%'
                    self._format_input_cell(ws, row, 2 + i)

            # SG&A $ rows (formulas)
            sga_dollar_start = sga_pct_start + 1 + len(sga_items) + 1
            self._add_section_header(ws, "SG&A ($M)", sga_dollar_start - 1, 1)

            for j, item in enumerate(sga_items):
                row = sga_dollar_start + j
                pct_row = sga_pct_start + 1 + j
                ws.cell(row=row, column=1, value=f"{item['name']} $")
                for i in range(n + 1):
                    col = get_column_letter(2 + i)
                    ws.cell(row=row, column=2 + i,
                            value=f"='{OM}'!{col}{om_rev_row}*{col}{pct_row}")
                    ws.cell(row=row, column=2 + i).number_format = '#,##0.0'

            sga_total_row = sga_dollar_start + len(sga_items)
            ws.cell(row=sga_total_row, column=1, value="Total SG&A")
            for i in range(n + 1):
                col = get_column_letter(2 + i)
                ws.cell(row=sga_total_row, column=2 + i,
                        value=f"=SUM({col}{sga_dollar_start}:{col}{sga_total_row - 1})")
                ws.cell(row=sga_total_row, column=2 + i).number_format = '#,##0.0'
                ws.cell(row=sga_total_row, column=2 + i).font = Font(bold=True)
                ws.cell(row=sga_total_row, column=2 + i).border = BOTTOM_BORDER

            # P&L Summary — all formulas
            summary_start = sga_total_row + 2
            self._add_section_header(ws, "P&L Summary ($M)", summary_start, 1)

            ws.cell(row=summary_start + 1, column=1, value="Revenue")
            for i in range(n + 1):
                col = get_column_letter(2 + i)
                ws.cell(row=summary_start + 1, column=2 + i,
                        value=f"='{OM}'!{col}{om_rev_row}")
                ws.cell(row=summary_start + 1, column=2 + i).number_format = '#,##0.0'

            ws.cell(row=summary_start + 2, column=1, value="Gross Profit")
            for i in range(n + 1):
                col = get_column_letter(2 + i)
                ws.cell(row=summary_start + 2, column=2 + i,
                        value=f"={col}{summary_start + 1}-{col}{cogs_total_row}")
                ws.cell(row=summary_start + 2, column=2 + i).number_format = '#,##0.0'

            ws.cell(row=summary_start + 3, column=1, value="  Gross Margin %")
            for i in range(n + 1):
                col = get_column_letter(2 + i)
                ws.cell(row=summary_start + 3, column=2 + i,
                        value=f"=IFERROR({col}{summary_start + 2}/{col}{summary_start + 1},0)")
                ws.cell(row=summary_start + 3, column=2 + i).number_format = '0.0%'
                ws.cell(row=summary_start + 3, column=2 + i).font = Font(italic=True, color="666666")

            ws.cell(row=summary_start + 5, column=1, value="EBITDA")
            for i in range(n + 1):
                col = get_column_letter(2 + i)
                ws.cell(row=summary_start + 5, column=2 + i,
                        value=f"={col}{summary_start + 2}-{col}{sga_total_row}")
                ws.cell(row=summary_start + 5, column=2 + i).number_format = '#,##0.0'
                ws.cell(row=summary_start + 5, column=2 + i).font = Font(bold=True)

            ws.cell(row=summary_start + 6, column=1, value="  EBITDA Margin %")
            for i in range(n + 1):
                col = get_column_letter(2 + i)
                ws.cell(row=summary_start + 6, column=2 + i,
                        value=f"=IFERROR({col}{summary_start + 5}/{col}{summary_start + 1},0)")
                ws.cell(row=summary_start + 6, column=2 + i).number_format = '0.0%'
                ws.cell(row=summary_start + 6, column=2 + i).font = Font(italic=True, color="666666")

            self.cell_map['expense_build'] = {
                'cogs_pct_start': 5,
                'cogs_dollar_start': cogs_dollar_start,
                'cogs_total_row': cogs_total_row,
                'sga_pct_start': sga_pct_start + 1,
                'sga_dollar_start': sga_dollar_start,
                'sga_total_row': sga_total_row,
                'summary_revenue_row': summary_start + 1,
                'summary_ebitda_row': summary_start + 5,
            }

        # Set column widths
        ws.column_dimensions['A'].width = 20
        for i in range(self.projection_years + 2):
            ws.column_dimensions[get_column_letter(2 + i)].width = 12

        return self

    # ============================================================
    # MODULE: DEBT SCHEDULE
    # ============================================================

    def add_debt_schedule(self, data: Dict = None) -> 'CoreModulesMixin':
        """Add Debt Schedule sheet with real Excel formulas."""
        ws = self.wb.create_sheet("Debt Schedule")
        self.sheets_created.append("Debt Schedule")

        A = "Assumptions"
        OM = "Operating Model"
        n = self.projection_years
        # Column layout: B=Entry, C=Year1, D=Year2, ...
        entry_col = 2
        y1_col = 3

        # Title
        self._add_title(ws, f"{self.company_name} - Debt Schedule", 1, 1)

        # Year headers
        ws.cell(row=3, column=1, value="")
        ws.cell(row=3, column=2, value="Entry")
        for i in range(n):
            ws.cell(row=3, column=y1_col + i, value=f"Year {i + 1}")
        self._format_header_row(ws, 3, 1, entry_col + n)

        if self.depth == ModelDepth.QUICK:
            # --- QUICK: Single tranche with flat paydown ---
            self._add_section_header(ws, "Debt Assumptions", 4, 1)

            # Row 5: Initial Debt = EBITDA x (Senior + Sub multiples)
            ws.cell(row=5, column=1, value="Initial Debt ($M)")
            ws.cell(row=5, column=2, value=f"='{A}'!B6*('{A}'!B12+'{A}'!B15)")
            ws.cell(row=5, column=2).number_format = '#,##0.0'
            self._format_input_cell(ws, 5, 2)

            # Row 6: Blended Interest Rate = average of senior and sub
            ws.cell(row=6, column=1, value="Interest Rate")
            ws.cell(row=6, column=2, value=f"=('{A}'!B13+'{A}'!B16)/2")
            ws.cell(row=6, column=2).number_format = '0.0%'
            self._format_input_cell(ws, 6, 2)

            # Row 7: Annual Paydown = 10% of initial debt
            ws.cell(row=7, column=1, value="Annual Paydown ($M)")
            ws.cell(row=7, column=2, value="=B5*0.1")
            ws.cell(row=7, column=2).number_format = '#,##0.0'
            self._format_input_cell(ws, 7, 2)

            # Debt schedule
            self._add_section_header(ws, "Debt Schedule ($M)", 9, 1)

            # Row 10: Beginning Balance — Entry=Initial, Y1+=prior ending
            ws.cell(row=10, column=1, value="Beginning Balance")
            ws.cell(row=10, column=entry_col, value="=B5")
            ws.cell(row=10, column=entry_col).number_format = '#,##0.0'
            for i in range(n):
                col = y1_col + i
                prev = get_column_letter(col - 1)
                this = get_column_letter(col)
                if i == 0:
                    ws.cell(row=10, column=col, value="=B5")  # Entry = initial
                else:
                    ws.cell(row=10, column=col, value=f"={prev}12")  # prior ending
                ws.cell(row=10, column=col).number_format = '#,##0.0'

            # Row 11: Paydown
            ws.cell(row=11, column=1, value="Paydown")
            for i in range(n):
                col = y1_col + i
                ws.cell(row=11, column=col, value="=-$B$7")
                ws.cell(row=11, column=col).number_format = '(#,##0.0)'

            # Row 12: Ending Balance = MAX(0, Beg + Paydown)
            ws.cell(row=12, column=1, value="Ending Balance")
            for i in range(n):
                col = y1_col + i
                this = get_column_letter(col)
                ws.cell(row=12, column=col, value=f"=MAX(0,{this}10+{this}11)")
                ws.cell(row=12, column=col).number_format = '#,##0.0'
                ws.cell(row=12, column=col).font = Font(bold=True)

            # Row 14: Interest = avg(beg, end) x rate
            ws.cell(row=14, column=1, value="Interest Expense")
            for i in range(n):
                col = y1_col + i
                this = get_column_letter(col)
                ws.cell(row=14, column=col, value=f"=({this}10+{this}12)/2*$B$6")
                ws.cell(row=14, column=col).number_format = '#,##0.0'

            # Cell map for QUICK mode
            self.cell_map['debt_schedule'] = {
                'mode': 'quick',
                'ending_balance_row': 12,
                'interest_row': 14,
                'total_debt_row': 12,  # same as ending in quick mode
            }

        else:
            # --- STANDARD/COMPREHENSIVE: Multi-tranche with cash sweep ---

            # Senior Debt section
            self._add_section_header(ws, "Senior Secured Debt ($M)", 4, 1)

            # Row 5: Senior Beg Balance — Entry col = EBITDA x Senior mult
            ws.cell(row=5, column=1, value="Beginning Balance")
            ws.cell(row=5, column=entry_col, value=f"='{A}'!B6*'{A}'!B12")
            ws.cell(row=5, column=entry_col).number_format = '#,##0.0'
            for i in range(n):
                col = y1_col + i
                prev = get_column_letter(col - 1)
                if i == 0:
                    ws.cell(row=5, column=col, value=f"=B5")  # Entry initial
                else:
                    ws.cell(row=5, column=col, value=f"={prev}8")  # prior ending
                ws.cell(row=5, column=col).number_format = '#,##0.0'

            # Row 6: Mandatory Amortization = -Initial x amort%
            ws.cell(row=6, column=1, value="Mandatory Amortization")
            for i in range(n):
                col = y1_col + i
                ws.cell(row=6, column=col, value=f"=-$B$5*'{A}'!B14")
                ws.cell(row=6, column=col).number_format = '(#,##0.0)'

            # Row 7: Cash Sweep = -MAX(0, (EBITDA - CapEx - Interest - Amort) x 50%)
            ws.cell(row=7, column=1, value="Cash Sweep")
            for i in range(n):
                col = y1_col + i
                this = get_column_letter(col)
                om_ebitda = f"'{OM}'!{this}13"
                om_capex = f"'{OM}'!{this}16"
                sr_interest = f"{this}10"
                sr_amort = f"{this}6"
                ws.cell(row=7, column=col,
                        value=f"=-MAX(0,({om_ebitda}+{om_capex}-{sr_interest}+{sr_amort})*0.5)")
                ws.cell(row=7, column=col).number_format = '(#,##0.0)'

            # Row 8: Senior Ending Balance = MAX(0, Beg + Amort + Sweep)
            ws.cell(row=8, column=1, value="Ending Balance")
            for i in range(n):
                col = y1_col + i
                this = get_column_letter(col)
                ws.cell(row=8, column=col, value=f"=MAX(0,{this}5+{this}6+{this}7)")
                ws.cell(row=8, column=col).number_format = '#,##0.0'
                ws.cell(row=8, column=col).font = Font(bold=True)

            # Row 9: Senior Interest Rate (input reference)
            ws.cell(row=9, column=1, value="Interest Rate")
            ws.cell(row=9, column=2, value=f"='{A}'!B13")
            ws.cell(row=9, column=2).number_format = '0.0%'
            self._format_input_cell(ws, 9, 2)

            # Row 10: Senior Interest Expense = avg(beg, end) x rate
            ws.cell(row=10, column=1, value="Interest Expense")
            for i in range(n):
                col = y1_col + i
                this = get_column_letter(col)
                ws.cell(row=10, column=col, value=f"=({this}5+{this}8)/2*$B$9")
                ws.cell(row=10, column=col).number_format = '#,##0.0'

            # --- Subordinated Debt section (row 12) ---
            sub_start = 12
            self._add_section_header(ws, "Subordinated Debt ($M)", sub_start, 1)

            # Row 13: Sub Balance = constant = EBITDA x Sub mult
            ws.cell(row=sub_start + 1, column=1, value="Beginning Balance")
            ws.cell(row=sub_start + 1, column=entry_col, value=f"='{A}'!B6*'{A}'!B15")
            ws.cell(row=sub_start + 1, column=entry_col).number_format = '#,##0.0'
            for i in range(n):
                col = y1_col + i
                ws.cell(row=sub_start + 1, column=col, value=f"=$B${sub_start + 1}")
                ws.cell(row=sub_start + 1, column=col).number_format = '#,##0.0'

            # Row 14: Sub Ending = same (no amort)
            ws.cell(row=sub_start + 2, column=1, value="Ending Balance")
            for i in range(n):
                col = y1_col + i
                this = get_column_letter(col)
                ws.cell(row=sub_start + 2, column=col, value=f"={this}{sub_start + 1}")
                ws.cell(row=sub_start + 2, column=col).number_format = '#,##0.0'
                ws.cell(row=sub_start + 2, column=col).font = Font(bold=True)

            # Row 15: Sub Interest Rate
            ws.cell(row=sub_start + 3, column=1, value="Interest Rate")
            ws.cell(row=sub_start + 3, column=2, value=f"='{A}'!B16")
            ws.cell(row=sub_start + 3, column=2).number_format = '0.0%'
            self._format_input_cell(ws, sub_start + 3, 2)

            # Row 16: Sub Interest = Balance x Rate
            ws.cell(row=sub_start + 4, column=1, value="Interest Expense")
            for i in range(n):
                col = y1_col + i
                this = get_column_letter(col)
                ws.cell(row=sub_start + 4, column=col,
                        value=f"={this}{sub_start + 1}*$B${sub_start + 3}")
                ws.cell(row=sub_start + 4, column=col).number_format = '#,##0.0'

            # --- Debt Summary (row 18) ---
            summary_start = sub_start + 6  # row 18
            self._add_section_header(ws, "Debt Summary", summary_start, 1)

            # Row 19: Total Debt = Senior Ending + Sub Ending
            td_row = summary_start + 1
            ws.cell(row=td_row, column=1, value="Total Debt")
            # Entry column: initial senior + sub
            ws.cell(row=td_row, column=entry_col, value=f"=B5+B{sub_start + 1}")
            ws.cell(row=td_row, column=entry_col).number_format = '#,##0.0'
            ws.cell(row=td_row, column=entry_col).font = Font(bold=True)
            for i in range(n):
                col = y1_col + i
                this = get_column_letter(col)
                ws.cell(row=td_row, column=col, value=f"={this}8+{this}{sub_start + 2}")
                ws.cell(row=td_row, column=col).number_format = '#,##0.0'
                ws.cell(row=td_row, column=col).font = Font(bold=True)

            # Row 20: Total Interest = Senior Interest + Sub Interest
            ti_row = summary_start + 2
            ws.cell(row=ti_row, column=1, value="Total Interest")
            for i in range(n):
                col = y1_col + i
                this = get_column_letter(col)
                ws.cell(row=ti_row, column=col, value=f"={this}10+{this}{sub_start + 4}")
                ws.cell(row=ti_row, column=col).number_format = '#,##0.0'

            # Row 22: Debt / EBITDA
            lever_row = summary_start + 4
            ws.cell(row=lever_row, column=1, value="Debt / EBITDA")
            # Entry
            ws.cell(row=lever_row, column=entry_col,
                    value=f"=IFERROR(B{td_row}/'{OM}'!B13,0)")
            ws.cell(row=lever_row, column=entry_col).number_format = '0.0"x"'
            for i in range(n):
                col = y1_col + i
                this = get_column_letter(col)
                ws.cell(row=lever_row, column=col,
                        value=f"=IFERROR({this}{td_row}/'{OM}'!{this}13,0)")
                ws.cell(row=lever_row, column=col).number_format = '0.0"x"'

            # Cell map for STANDARD mode
            self.cell_map['debt_schedule'] = {
                'mode': 'standard',
                'senior_ending_row': 8,
                'sub_ending_row': sub_start + 2,
                'total_debt_row': td_row,
                'total_interest_row': ti_row,
                'leverage_row': lever_row,
            }

        # Set column widths
        ws.column_dimensions['A'].width = 22
        for i in range(n + 2):
            ws.column_dimensions[get_column_letter(2 + i)].width = 12

        return self

    # ============================================================
    # MODULE: WACC CALCULATION
    # ============================================================

    def add_wacc_calculation(self, data: Dict = None) -> 'CoreModulesMixin':
        """Add WACC Calculation sheet with live Excel formulas."""
        ws = self.wb.create_sheet("WACC")
        self.sheets_created.append("WACC")

        a = self.assumptions

        # Title
        self._add_title(ws, f"{self.company_name} - WACC Calculation", 1, 1)

        if self.depth == ModelDepth.QUICK:
            # Simple: Just show WACC assumption
            self._add_section_header(ws, "WACC Assumption", 3, 1)

            ws.cell(row=4, column=1, value="WACC")
            ws.cell(row=4, column=2, value=0.10)
            ws.cell(row=4, column=2).number_format = '0.0%'
            self._format_input_cell(ws, 4, 2)

            self.cell_map['wacc'] = {'wacc': 'B4'}

        else:
            # Standard/Comprehensive: Full CAPM with live formulas

            # Cost of Equity section
            self._add_section_header(ws, "Cost of Equity (CAPM)", 3, 1)

            # Row 4-7: CAPM inputs (blue font = editable)
            capm_inputs = [
                (4, "Risk-Free Rate (10Y Treasury)", a.risk_free_rate, '0.00%'),
                (5, "Equity Risk Premium", a.equity_risk_premium, '0.00%'),
                (6, "Beta (Levered)", a.beta, '0.00'),
                (7, "Size Premium", a.size_premium, '0.00%'),
            ]
            for row, name, value, fmt in capm_inputs:
                ws.cell(row=row, column=1, value=name)
                ws.cell(row=row, column=2, value=value)
                ws.cell(row=row, column=2).number_format = fmt
                self._format_input_cell(ws, row, 2)

            # Row 9: Cost of Equity = Rf + (Beta x ERP) + Size Premium
            ws.cell(row=9, column=1, value="Cost of Equity")
            ws.cell(row=9, column=2, value="=B4+(B6*B5)+B7")
            ws.cell(row=9, column=2).number_format = '0.0%'
            ws.cell(row=9, column=2).font = Font(bold=True)
            ws.cell(row=9, column=2).border = DOUBLE_BORDER

            # Cost of Debt section
            self._add_section_header(ws, "Cost of Debt", 11, 1)

            ws.cell(row=12, column=1, value="Pre-Tax Cost of Debt")
            ws.cell(row=12, column=2, value=a.cost_of_debt)
            ws.cell(row=12, column=2).number_format = '0.0%'
            self._format_input_cell(ws, 12, 2)

            ws.cell(row=13, column=1, value="Tax Rate")
            ws.cell(row=13, column=2, value=a.tax_rate)
            ws.cell(row=13, column=2).number_format = '0.0%'
            self._format_input_cell(ws, 13, 2)

            # Row 14: After-Tax Cost of Debt = Rate x (1 - Tax)
            ws.cell(row=14, column=1, value="After-Tax Cost of Debt")
            ws.cell(row=14, column=2, value="=B12*(1-B13)")
            ws.cell(row=14, column=2).number_format = '0.0%'
            ws.cell(row=14, column=2).font = Font(bold=True)
            ws.cell(row=14, column=2).border = DOUBLE_BORDER

            # Capital Structure
            self._add_section_header(ws, "Capital Structure", 16, 1)

            ws.cell(row=17, column=1, value="Target Debt / Equity")
            ws.cell(row=17, column=2, value=a.target_debt_equity)
            ws.cell(row=17, column=2).number_format = '0.0%'
            self._format_input_cell(ws, 17, 2)

            # Row 18-19: Weights as formulas
            ws.cell(row=18, column=1, value="Debt Weight")
            ws.cell(row=18, column=2, value="=B17/(1+B17)")
            ws.cell(row=18, column=2).number_format = '0.0%'

            ws.cell(row=19, column=1, value="Equity Weight")
            ws.cell(row=19, column=2, value="=1-B18")
            ws.cell(row=19, column=2).number_format = '0.0%'

            # WACC Calculation
            self._add_section_header(ws, "WACC Calculation", 21, 1)

            # Row 22: WACC = (E/V x Re) + (D/V x Rd x(1-T))
            ws.cell(row=22, column=1, value="WACC")
            ws.cell(row=22, column=2, value="=(B19*B9)+(B18*B14)")
            ws.cell(row=22, column=2).number_format = '0.0%'
            ws.cell(row=22, column=2).font = Font(bold=True, size=14)
            ws.cell(row=22, column=2).fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")

            # Formula reference
            ws.cell(row=24, column=1, value="Formula: (E/V × Re) + (D/V × Rd × (1-T))")

            self.cell_map['wacc'] = {
                'risk_free_rate': 'B4',
                'erp': 'B5',
                'beta': 'B6',
                'size_premium': 'B7',
                'cost_of_equity': 'B9',
                'cost_of_debt': 'B12',
                'tax_rate': 'B13',
                'after_tax_debt': 'B14',
                'de_ratio': 'B17',
                'debt_weight': 'B18',
                'equity_weight': 'B19',
                'wacc': 'B22',
            }

        # Set column widths
        ws.column_dimensions['A'].width = 30
        ws.column_dimensions['B'].width = 15

        return self

    # ============================================================
    # MODULE: WORKING CAPITAL
    # ============================================================

    def add_working_capital(self, data: Dict = None) -> 'CoreModulesMixin':
        """Add Working Capital Analysis sheet with real Excel formulas."""
        ws = self.wb.create_sheet("Working Capital")
        self.sheets_created.append("Working Capital")

        A = "Assumptions"
        OM = "Operating Model"
        n = self.projection_years
        ltm_col = 2
        y1_col = 3

        # Title
        self._add_title(ws, f"{self.company_name} - Working Capital Analysis", 1, 1)

        # Year headers
        self._year_headers(ws, 3, 2)

        if self.depth == ModelDepth.QUICK:
            # Simple: NWC as % of revenue, referencing Assumptions + Operating Model
            self._add_section_header(ws, "Working Capital Assumptions", 4, 1)

            # Row 5: NWC % Revenue — reference Assumptions!B23
            ws.cell(row=5, column=1, value="NWC % of Revenue")
            for i in range(n + 1):
                col = ltm_col + i
                ws.cell(row=5, column=col, value=f"='{A}'!B23")
                ws.cell(row=5, column=col).number_format = '0.0%'
                self._format_input_cell(ws, 5, col)

            self._add_section_header(ws, "Working Capital ($M)", 7, 1)

            # Row 8: NWC = Revenue x NWC%
            ws.cell(row=8, column=1, value="Net Working Capital")
            for i in range(n + 1):
                col = ltm_col + i
                this = get_column_letter(col)
                ws.cell(row=8, column=col, value=f"='{OM}'!{this}10*{this}5")
                ws.cell(row=8, column=col).number_format = '#,##0.0'

            # Row 9: Change in NWC
            ws.cell(row=9, column=1, value="Change in NWC")
            ws.cell(row=9, column=ltm_col, value="—")
            for i in range(n):
                col = y1_col + i
                prev = get_column_letter(col - 1)
                this = get_column_letter(col)
                ws.cell(row=9, column=col, value=f"={this}8-{prev}8")
                ws.cell(row=9, column=col).number_format = '#,##0.0'

            self.cell_map['working_capital'] = {
                'nwc_row': 8,
                'change_nwc_row': 9,
            }

        else:
            # Standard/Comprehensive: Days-based calculation with formulas
            data = data or {}
            self._add_section_header(ws, "Working Capital Assumptions (Days)", 4, 1)

            days = data.get('days', {
                'ar_days': 45, 'inventory_days': 60, 'ap_days': 30,
                'other_ca_pct': 0.02, 'other_cl_pct': 0.03,
            })

            # Row 5-7: Input assumptions (still hardcoded input values, not formula refs)
            ws.cell(row=5, column=1, value="Accounts Receivable Days")
            for i in range(n + 1):
                ws.cell(row=5, column=ltm_col + i, value=days['ar_days'])
                self._format_input_cell(ws, 5, ltm_col + i)

            ws.cell(row=6, column=1, value="Inventory Days")
            for i in range(n + 1):
                ws.cell(row=6, column=ltm_col + i, value=days['inventory_days'])
                self._format_input_cell(ws, 6, ltm_col + i)

            ws.cell(row=7, column=1, value="Accounts Payable Days")
            for i in range(n + 1):
                ws.cell(row=7, column=ltm_col + i, value=days['ap_days'])
                self._format_input_cell(ws, 7, ltm_col + i)

            # Current Assets (formulas referencing Operating Model revenue)
            self._add_section_header(ws, "Current Assets ($M)", 9, 1)

            # Row 10: AR = Revenue x AR Days / 365
            ws.cell(row=10, column=1, value="Accounts Receivable")
            for i in range(n + 1):
                col = ltm_col + i
                this = get_column_letter(col)
                ws.cell(row=10, column=col, value=f"='{OM}'!{this}10*{this}5/365")
                ws.cell(row=10, column=col).number_format = '#,##0.0'

            # Row 11: Inventory = Revenue x 0.60 x Inv Days / 365
            ws.cell(row=11, column=1, value="Inventory")
            for i in range(n + 1):
                col = ltm_col + i
                this = get_column_letter(col)
                ws.cell(row=11, column=col, value=f"='{OM}'!{this}10*0.6*{this}6/365")
                ws.cell(row=11, column=col).number_format = '#,##0.0'

            # Row 12: Other CA = Revenue x other_ca_pct
            ws.cell(row=12, column=1, value="Other Current Assets")
            for i in range(n + 1):
                col = ltm_col + i
                this = get_column_letter(col)
                ws.cell(row=12, column=col, value=f"='{OM}'!{this}10*{days['other_ca_pct']}")
                ws.cell(row=12, column=col).number_format = '#,##0.0'

            # Row 13: Total CA
            ws.cell(row=13, column=1, value="Total Current Assets")
            for i in range(n + 1):
                col = ltm_col + i
                this = get_column_letter(col)
                ws.cell(row=13, column=col, value=f"=SUM({this}10:{this}12)")
                ws.cell(row=13, column=col).number_format = '#,##0.0'
                ws.cell(row=13, column=col).font = Font(bold=True)

            # Current Liabilities
            self._add_section_header(ws, "Current Liabilities ($M)", 15, 1)

            # Row 16: AP = Revenue x 0.60 x AP Days / 365
            ws.cell(row=16, column=1, value="Accounts Payable")
            for i in range(n + 1):
                col = ltm_col + i
                this = get_column_letter(col)
                ws.cell(row=16, column=col, value=f"='{OM}'!{this}10*0.6*{this}7/365")
                ws.cell(row=16, column=col).number_format = '#,##0.0'

            # Row 17: Other CL = Revenue x other_cl_pct
            ws.cell(row=17, column=1, value="Other Current Liabilities")
            for i in range(n + 1):
                col = ltm_col + i
                this = get_column_letter(col)
                ws.cell(row=17, column=col, value=f"='{OM}'!{this}10*{days['other_cl_pct']}")
                ws.cell(row=17, column=col).number_format = '#,##0.0'

            # Row 18: Total CL
            ws.cell(row=18, column=1, value="Total Current Liabilities")
            for i in range(n + 1):
                col = ltm_col + i
                this = get_column_letter(col)
                ws.cell(row=18, column=col, value=f"=SUM({this}16:{this}17)")
                ws.cell(row=18, column=col).number_format = '#,##0.0'
                ws.cell(row=18, column=col).font = Font(bold=True)

            # Net Working Capital
            self._add_section_header(ws, "Net Working Capital ($M)", 20, 1)

            # Row 21: NWC = Total CA - Total CL
            ws.cell(row=21, column=1, value="Net Working Capital")
            for i in range(n + 1):
                col = ltm_col + i
                this = get_column_letter(col)
                ws.cell(row=21, column=col, value=f"={this}13-{this}18")
                ws.cell(row=21, column=col).number_format = '#,##0.0'
                ws.cell(row=21, column=col).font = Font(bold=True)

            # Row 22: NWC % Revenue
            ws.cell(row=22, column=1, value="NWC % of Revenue")
            for i in range(n + 1):
                col = ltm_col + i
                this = get_column_letter(col)
                ws.cell(row=22, column=col, value=f"={this}21/'{OM}'!{this}10")
                ws.cell(row=22, column=col).number_format = '0.0%'

            # Row 23: Change in NWC
            ws.cell(row=23, column=1, value="Change in NWC")
            ws.cell(row=23, column=ltm_col, value="—")
            for i in range(n):
                col = y1_col + i
                prev = get_column_letter(col - 1)
                this = get_column_letter(col)
                ws.cell(row=23, column=col, value=f"={this}21-{prev}21")
                ws.cell(row=23, column=col).number_format = '#,##0.0'

            # Cash Conversion Cycle (formulas referencing local day inputs)
            self._add_section_header(ws, "Cash Conversion Cycle", 25, 1)

            ws.cell(row=26, column=1, value="DSO (Days Sales Outstanding)")
            ws.cell(row=26, column=2, value="=B5")

            ws.cell(row=27, column=1, value="DIO (Days Inventory Outstanding)")
            ws.cell(row=27, column=2, value="=B6")

            ws.cell(row=28, column=1, value="DPO (Days Payable Outstanding)")
            ws.cell(row=28, column=2, value="=B7")

            ws.cell(row=29, column=1, value="Cash Conversion Cycle")
            ws.cell(row=29, column=2, value="=B26+B27-B28")
            ws.cell(row=29, column=2).font = Font(bold=True)

            self.cell_map['working_capital'] = {
                'nwc_row': 21,
                'change_nwc_row': 23,
            }

        # Set column widths
        ws.column_dimensions['A'].width = 28
        for i in range(n + 2):
            ws.column_dimensions[get_column_letter(2 + i)].width = 12

        return self

    # ============================================================
    # MODULE: RETURNS ANALYSIS
    # ============================================================

    def add_returns_analysis(self, data: Dict = None) -> 'CoreModulesMixin':
        """Add Returns Analysis sheet with real Excel formulas referencing other sheets."""
        ws = self.wb.create_sheet("Returns Analysis")
        self.sheets_created.append("Returns Analysis")

        A = "Assumptions"
        SU = "Sources & Uses"
        OM = "Operating Model"
        DS = "Debt Schedule"
        n = self.projection_years

        # Determine debt schedule total debt row based on mode
        ds_info = self.cell_map.get('debt_schedule', {})
        ds_mode = ds_info.get('mode', 'quick')
        td_row = ds_info.get('total_debt_row', 12)

        # Title
        self._add_title(ws, f"{self.company_name} - Returns Analysis", 1, 1)

        # --- Entry Assumptions (rows 3-8) — formulas referencing S&U ---
        self._add_section_header(ws, "Entry Assumptions", 3, 1)

        ws.cell(row=4, column=1, value="LTM EBITDA ($M)")
        ws.cell(row=4, column=2, value=f"='{A}'!B6")
        ws.cell(row=4, column=2).number_format = '#,##0.0'

        ws.cell(row=5, column=1, value="Entry Multiple")
        ws.cell(row=5, column=2, value=f"='{A}'!B7")
        ws.cell(row=5, column=2).number_format = '0.0"x"'

        ws.cell(row=6, column=1, value="Enterprise Value ($M)")
        ws.cell(row=6, column=2, value=f"='{SU}'!B10")
        ws.cell(row=6, column=2).number_format = '#,##0.0'

        ws.cell(row=7, column=1, value="Total Debt ($M)")
        ws.cell(row=7, column=2, value=f"='{SU}'!B4+'{SU}'!B5")
        ws.cell(row=7, column=2).number_format = '#,##0.0'

        ws.cell(row=8, column=1, value="Equity Investment ($M)")
        ws.cell(row=8, column=2, value=f"='{SU}'!B6")
        ws.cell(row=8, column=2).number_format = '#,##0.0'

        # --- Exit Assumptions (rows 10-12) ---
        self._add_section_header(ws, "Exit Assumptions", 10, 1)

        ws.cell(row=11, column=1, value="Exit Multiple")
        ws.cell(row=11, column=2, value=f"='{A}'!B8")
        ws.cell(row=11, column=2).number_format = '0.0"x"'
        self._format_input_cell(ws, 11, 2)

        ws.cell(row=12, column=1, value="Hold Period (Years)")
        ws.cell(row=12, column=2, value=f"='{A}'!B9")
        self._format_input_cell(ws, 12, 2)

        # --- Returns by Exit Year (rows 14+) ---
        self._add_section_header(ws, "Returns by Exit Year", 14, 1)

        headers = ["Exit Year", "EBITDA", "Exit EV", "Net Debt", "Equity Value", "MOIC", "IRR"]
        for i, header in enumerate(headers):
            ws.cell(row=15, column=1 + i, value=header)
        self._format_header_row(ws, 15, 1, len(headers))

        for year in range(3, n + 1):
            row = 16 + (year - 3)
            # Column in Operating Model / Debt Schedule for this year
            # Year 3 = column E (col 5), Year 4 = F, Year 5 = G
            yr_col = get_column_letter(2 + year)  # B=LTM, C=Y1, D=Y2, E=Y3...

            ws.cell(row=row, column=1, value=f"Year {year}")

            # Col B: Exit EBITDA from Operating Model
            ws.cell(row=row, column=2, value=f"='{OM}'!{yr_col}13")
            ws.cell(row=row, column=2).number_format = '#,##0.0'

            # Col C: Exit EV = Exit EBITDA x Exit Multiple
            ws.cell(row=row, column=3, value=f"=B{row}*$B$11")
            ws.cell(row=row, column=3).number_format = '#,##0.0'

            # Col D: Net Debt at exit from Debt Schedule
            ws.cell(row=row, column=4, value=f"='{DS}'!{yr_col}{td_row}")
            ws.cell(row=row, column=4).number_format = '#,##0.0'

            # Col E: Equity Value = Exit EV - Net Debt
            ws.cell(row=row, column=5, value=f"=C{row}-D{row}")
            ws.cell(row=row, column=5).number_format = '#,##0.0'

            # Col F: MOIC = Equity Value / Entry Equity
            ws.cell(row=row, column=6, value=f"=IFERROR(E{row}/$B$8,0)")
            ws.cell(row=row, column=6).number_format = '0.00"x"'

            # Col G: IRR = (MOIC^(1/year)) - 1
            ws.cell(row=row, column=7, value=f"=IFERROR((F{row}^(1/{year}))-1,0)")
            ws.cell(row=row, column=7).number_format = '0.0%'

        # Highlight base case (Year 5)
        if n >= 5:
            base_row = 16 + (5 - 3)
            for col in range(1, 8):
                ws.cell(row=base_row, column=col).fill = PatternFill(
                    start_color="90EE90", end_color="90EE90", fill_type="solid")

        # --- Value Creation Bridge (Standard/Comprehensive only) ---
        if self.depth != ModelDepth.QUICK:
            bridge_start = 16 + (n - 2) + 2
            self._add_section_header(ws, "Value Creation Bridge (Base Case)", bridge_start, 1)

            # Use last exit year row for base case references
            base_data_row = 16 + (n - 3)  # row for Year n (the last one)
            exit_yr_col = get_column_letter(2 + n)  # column for exit year in OM/DS

            br = bridge_start + 1
            # Entry Equity
            ws.cell(row=br, column=1, value="Entry Equity")
            ws.cell(row=br, column=2, value="=$B$8")
            ws.cell(row=br, column=2).number_format = '#,##0.0'

            # EBITDA Growth value = (Exit EBITDA - LTM EBITDA) x Exit Multiple
            ws.cell(row=br + 1, column=1, value="(+) EBITDA Growth")
            ws.cell(row=br + 1, column=2,
                    value=f"=('{OM}'!{exit_yr_col}13-'{A}'!B6)*$B$11")
            ws.cell(row=br + 1, column=2).number_format = '#,##0.0'

            # Multiple Expansion = Exit EBITDA x (Exit Mult - Entry Mult)
            ws.cell(row=br + 2, column=1, value="(+) Multiple Expansion")
            ws.cell(row=br + 2, column=2,
                    value=f"='{OM}'!{exit_yr_col}13*('{A}'!B8-'{A}'!B7)")
            ws.cell(row=br + 2, column=2).number_format = '#,##0.0'

            # Debt Paydown = Entry Debt - Exit Debt
            ws.cell(row=br + 3, column=1, value="(+) Debt Paydown")
            ws.cell(row=br + 3, column=2,
                    value=f"=$B$7-'{DS}'!{exit_yr_col}{td_row}")
            ws.cell(row=br + 3, column=2).number_format = '#,##0.0'

            # Exit Equity = sum of bridge
            ws.cell(row=br + 4, column=1, value="(=) Exit Equity")
            ws.cell(row=br + 4, column=2,
                    value=f"=SUM(B{br}:B{br + 3})")
            ws.cell(row=br + 4, column=2).number_format = '#,##0.0'
            ws.cell(row=br + 4, column=1).font = Font(bold=True)
            ws.cell(row=br + 4, column=2).font = Font(bold=True)
            ws.cell(row=br + 4, column=2).border = DOUBLE_BORDER

        # Set column widths
        ws.column_dimensions['A'].width = 22
        for i in range(7):
            ws.column_dimensions[get_column_letter(2 + i)].width = 12

        # cell_map for cross-sheet references
        base_row = 16  # Year 3 row (first exit year)
        self.cell_map['returns_analysis'] = {
            'entry_equity': 'B8',
            'exit_multiple': 'B11',
            'moic_row': base_row,      # MOIC is column F (6) for each exit year
            'irr_row': base_row,       # IRR is column G (7) for each exit year
            'moic_col': 6,
            'irr_col': 7,
        }

        return self

    # ============================================================
    # MODULE: SENSITIVITY TABLES
    # ============================================================

    def add_sensitivity_tables(self, data: Dict = None) -> 'CoreModulesMixin':
        """Add Sensitivity Analysis sheet with formula grids referencing other sheets."""
        ws = self.wb.create_sheet("Sensitivity")
        self.sheets_created.append("Sensitivity")

        A = "Assumptions"
        OM = "Operating Model"
        DS = "Debt Schedule"
        n = self.projection_years

        # Exit year column in Operating Model (Year n)
        exit_col = get_column_letter(2 + n)  # e.g. G for year 5

        # Debt schedule total debt row
        ds_info = self.cell_map.get('debt_schedule', {})
        td_row = ds_info.get('total_debt_row', 12)

        # We store reference cells in hidden rows for the formula grid
        # Row 2: helper cells
        # H2 = Exit EBITDA, H3 = Total Debt at exit, H4 = Fees, H5 = Hold period
        ws.cell(row=2, column=8, value=f"='{OM}'!{exit_col}13")    # Exit EBITDA
        ws.cell(row=2, column=9, value=f"='{DS}'!{exit_col}{td_row}")  # Debt at exit
        ws.cell(row=2, column=10, value=f"='{A}'!B6*('{A}'!B12+'{A}'!B15)")  # Total entry debt
        ws.cell(row=2, column=11, value=f"='{A}'!B9")  # Hold period
        ws.cell(row=2, column=12, value=f"='{A}'!B19")  # Txn fee %
        ws.cell(row=2, column=13, value=f"='{A}'!B20")  # Fin fee %
        ws.cell(row=2, column=14, value=f"='{A}'!B6")   # LTM EBITDA
        # Labels for reference
        for c, label in [(8, "Exit EBITDA"), (9, "Debt@Exit"), (10, "Entry Debt"),
                         (11, "Hold Yrs"), (12, "Txn%"), (13, "Fin%"), (14, "LTM EBITDA")]:
            ws.cell(row=1, column=c, value=label)
            ws.cell(row=1, column=c).font = Font(color="999999", size=8)

        entry_multiples = [7.0, 7.5, 8.0, 8.5, 9.0]
        exit_multiples = [7.0, 7.5, 8.0, 8.5, 9.0]

        # Title
        self._add_title(ws, f"{self.company_name} - Sensitivity Analysis", 3, 1)

        # --- IRR Matrix (rows 5-10) ---
        self._add_section_header(ws, "IRR Sensitivity: Entry Multiple vs Exit Multiple", 5, 1)

        ws.cell(row=6, column=1, value="Entry \\ Exit")
        for i, em in enumerate(exit_multiples):
            ws.cell(row=6, column=2 + i, value=em)
            ws.cell(row=6, column=2 + i).number_format = '0.0"x"'
        self._format_header_row(ws, 6, 1, 1 + len(exit_multiples))

        for j, entry_mult in enumerate(entry_multiples):
            row = 7 + j
            ws.cell(row=row, column=1, value=entry_mult)
            ws.cell(row=row, column=1).number_format = '0.0"x"'
            ws.cell(row=row, column=1).font = HEADER_FONT
            ws.cell(row=row, column=1).fill = HEADER_FILL

            for i, exit_mult in enumerate(exit_multiples):
                # MOIC = (exit_ebitda * exit_mult - debt_at_exit) /
                #        (entry_mult * ltm_ebitda + fees - entry_debt)
                # fees = entry_mult*ltm_ebitda*txn_fee_pct + entry_debt*fin_fee_pct
                # IRR = (MOIC^(1/hold))-1
                exit_mult_ref = f"${get_column_letter(2 + i)}$6"  # exit mult from header
                entry_mult_ref = f"$A${row}"  # entry mult from row header
                # Build MOIC formula inline
                moic_f = (
                    f"(($H$2*{exit_mult_ref}-$I$2)"
                    f"/({entry_mult_ref}*$N$2"
                    f"+{entry_mult_ref}*$N$2*$L$2+$J$2*$M$2"
                    f"-$J$2))"
                )
                irr_f = f"=IFERROR(({moic_f}^(1/$K$2))-1,0)"

                cell = ws.cell(row=row, column=2 + i, value=irr_f)
                cell.number_format = '0.0%'

        # --- MOIC Matrix ---
        moic_start = 7 + len(entry_multiples) + 2  # row 14
        self._add_section_header(ws, "MOIC Sensitivity: Entry Multiple vs Exit Multiple", moic_start, 1)

        ws.cell(row=moic_start + 1, column=1, value="Entry \\ Exit")
        for i, em in enumerate(exit_multiples):
            ws.cell(row=moic_start + 1, column=2 + i, value=em)
            ws.cell(row=moic_start + 1, column=2 + i).number_format = '0.0"x"'
        self._format_header_row(ws, moic_start + 1, 1, 1 + len(exit_multiples))

        for j, entry_mult in enumerate(entry_multiples):
            row = moic_start + 2 + j
            ws.cell(row=row, column=1, value=entry_mult)
            ws.cell(row=row, column=1).number_format = '0.0"x"'
            ws.cell(row=row, column=1).font = HEADER_FONT
            ws.cell(row=row, column=1).fill = HEADER_FILL

            for i, exit_mult in enumerate(exit_multiples):
                exit_mult_ref = f"${get_column_letter(2 + i)}${moic_start + 1}"
                entry_mult_ref = f"$A${row}"
                moic_f = (
                    f"=IFERROR(($H$2*{exit_mult_ref}-$I$2)"
                    f"/({entry_mult_ref}*$N$2"
                    f"+{entry_mult_ref}*$N$2*$L$2+$J$2*$M$2"
                    f"-$J$2),0)"
                )

                cell = ws.cell(row=row, column=2 + i, value=moic_f)
                cell.number_format = '0.00"x"'

        # Legend
        legend_start = moic_start + 2 + len(entry_multiples) + 2
        ws.cell(row=legend_start, column=1, value="Legend:")
        ws.cell(row=legend_start + 1, column=1, value="Note: Values update automatically when Assumptions change")
        ws.cell(row=legend_start + 1, column=1).font = Font(italic=True, color="666666")

        # --- Conditional Formatting ---
        # Color scale on IRR matrix (red-yellow-green)
        irr_range = f"B7:{get_column_letter(1 + len(exit_multiples))}{6 + len(entry_multiples)}"
        self._add_color_scale(ws, irr_range)
        # Color scale on MOIC matrix
        moic_range = f"B{moic_start + 2}:{get_column_letter(1 + len(exit_multiples))}{moic_start + 1 + len(entry_multiples)}"
        self._add_color_scale(ws, moic_range)

        # Set column widths
        ws.column_dimensions['A'].width = 15
        for i in range(6):
            ws.column_dimensions[get_column_letter(2 + i)].width = 10

        return self

    # ============================================================
    # MODULE: SCENARIO ANALYSIS (Bull/Bear/Base)
    # ============================================================

    def add_scenario_analysis(self, data: Dict = None) -> 'CoreModulesMixin':
        """Add Bull/Bear/Base scenario comparison sheet with live formulas."""
        ws = self.wb.create_sheet("Scenario Analysis")
        self.sheets_created.append("Scenario Analysis")

        A = "Assumptions"
        a = self.assumptions
        data = data or {}
        n = self.projection_years

        # Define scenarios (input assumptions — written as editable values)
        scenarios = data.get('scenarios', {
            'bear': {
                'name': 'Bear Case',
                'revenue_growth': [0.03, 0.02, 0.02, 0.02, 0.02],
                'ebitda_margin': [0.18, 0.18, 0.18, 0.18, 0.18],
                'exit_multiple': a.exit_multiple - 1.0,
                'probability': 0.25,
            },
            'base': {
                'name': 'Base Case',
                'revenue_growth': list(a.revenue_growth),
                'ebitda_margin': list(a.ebitda_margin),
                'exit_multiple': a.exit_multiple,
                'probability': 0.50,
            },
            'bull': {
                'name': 'Bull Case',
                'revenue_growth': [0.12, 0.10, 0.09, 0.08, 0.07],
                'ebitda_margin': [0.22, 0.24, 0.25, 0.26, 0.26],
                'exit_multiple': a.exit_multiple + 1.0,
                'probability': 0.25,
            },
        })

        # Title
        self._add_title(ws, f"{self.company_name} - Scenario Analysis", 1, 1)

        # -- SECTION 1: Scenario Input Assumptions (values, user-editable) --
        self._add_section_header(ws, "Scenario Assumptions", 3, 1)

        headers = ["Assumption", "Bear Case", "Base Case", "Bull Case"]
        for i, h in enumerate(headers):
            ws.cell(row=4, column=1 + i, value=h)
        self._format_header_row(ws, 4, 1, 4)

        # Row 5: Revenue Growth (Y1)
        ws.cell(row=5, column=1, value="Revenue Growth (Y1)")
        ws.cell(row=5, column=2, value=scenarios['bear']['revenue_growth'][0])
        ws.cell(row=5, column=3, value=scenarios['base']['revenue_growth'][0])
        ws.cell(row=5, column=4, value=scenarios['bull']['revenue_growth'][0])
        for col in range(2, 5):
            ws.cell(row=5, column=col).number_format = '0.0%'
            self._format_input_cell(ws, 5, col)

        # Row 6: Exit EBITDA Margin
        ws.cell(row=6, column=1, value="Exit EBITDA Margin")
        ws.cell(row=6, column=2, value=scenarios['bear']['ebitda_margin'][-1])
        ws.cell(row=6, column=3, value=scenarios['base']['ebitda_margin'][-1])
        ws.cell(row=6, column=4, value=scenarios['bull']['ebitda_margin'][-1])
        for col in range(2, 5):
            ws.cell(row=6, column=col).number_format = '0.0%'
            self._format_input_cell(ws, 6, col)

        # Row 7: Exit Multiple
        ws.cell(row=7, column=1, value="Exit Multiple")
        ws.cell(row=7, column=2, value=scenarios['bear']['exit_multiple'])
        ws.cell(row=7, column=3, value=scenarios['base']['exit_multiple'])
        ws.cell(row=7, column=4, value=scenarios['bull']['exit_multiple'])
        for col in range(2, 5):
            ws.cell(row=7, column=col).number_format = '0.0x'
            self._format_input_cell(ws, 7, col)

        # Row 8: Probability Weight
        ws.cell(row=8, column=1, value="Probability Weight")
        ws.cell(row=8, column=2, value=scenarios['bear']['probability'])
        ws.cell(row=8, column=3, value=scenarios['base']['probability'])
        ws.cell(row=8, column=4, value=scenarios['bull']['probability'])
        for col in range(2, 5):
            ws.cell(row=8, column=col).number_format = '0%'
            self._format_input_cell(ws, 8, col)

        # -- SECTION 2: Shared entry-level assumptions (formulas from Assumptions) --
        # Row 9: blank separator
        self._add_section_header(ws, "Entry Assumptions (from Assumptions sheet)", 9, 1)

        # B10: LTM Revenue, B11: LTM EBITDA, B12: Entry Multiple
        # B13: Total Debt, B14: Fees (4% of EV), B15: Equity Check
        ws.cell(row=10, column=1, value="LTM Revenue ($M)")
        ws.cell(row=10, column=2, value=f"='{A}'!B5")
        ws.cell(row=10, column=2).number_format = '#,##0.0'

        ws.cell(row=11, column=1, value="LTM EBITDA ($M)")
        ws.cell(row=11, column=2, value=f"='{A}'!B6")
        ws.cell(row=11, column=2).number_format = '#,##0.0'

        ws.cell(row=12, column=1, value="Entry Multiple")
        ws.cell(row=12, column=2, value=f"='{A}'!B7")
        ws.cell(row=12, column=2).number_format = '0.0"x"'

        ws.cell(row=13, column=1, value="Total Debt ($M)")
        ws.cell(row=13, column=2, value=f"=B11*('{A}'!B12+'{A}'!B15)")
        ws.cell(row=13, column=2).number_format = '#,##0.0'

        ws.cell(row=14, column=1, value="Fees ($M, 4% EV)")
        ws.cell(row=14, column=2, value="=B11*B12*0.04")
        ws.cell(row=14, column=2).number_format = '#,##0.0'

        ws.cell(row=15, column=1, value="Equity Check ($M)")
        ws.cell(row=15, column=2, value="=B11*B12+B14-B13")
        ws.cell(row=15, column=2).number_format = '#,##0.0'
        ws.cell(row=15, column=2).font = Font(bold=True)

        ws.cell(row=16, column=1, value="Hold Period (yrs)")
        ws.cell(row=16, column=2, value=f"='{A}'!B9")

        # -- SECTION 3: Scenario Outputs (ALL formulas) --
        self._add_section_header(ws, "Scenario Outputs ($M)", 18, 1)

        output_headers = ["Metric", "Bear Case", "Base Case", "Bull Case"]
        for i, h in enumerate(output_headers):
            ws.cell(row=19, column=1 + i, value=h)
        self._format_header_row(ws, 19, 1, 4)

        # Row 20: Exit Revenue = LTM_Rev * (1+growth)^hold_period (simplified)
        ws.cell(row=20, column=1, value="Exit Revenue")
        for col_idx in range(2, 5):
            c = get_column_letter(col_idx)
            ws.cell(row=20, column=col_idx,
                    value=f"=$B$10*(1+{c}5)^$B$16")
            ws.cell(row=20, column=col_idx).number_format = '#,##0.0'

        # Row 21: Exit EBITDA = Exit Revenue x Exit EBITDA Margin
        ws.cell(row=21, column=1, value="Exit EBITDA")
        for col_idx in range(2, 5):
            c = get_column_letter(col_idx)
            ws.cell(row=21, column=col_idx,
                    value=f"={c}20*{c}6")
            ws.cell(row=21, column=col_idx).number_format = '#,##0.0'

        # Row 22: Exit EV = Exit EBITDA x Exit Multiple
        ws.cell(row=22, column=1, value="Exit EV")
        for col_idx in range(2, 5):
            c = get_column_letter(col_idx)
            ws.cell(row=22, column=col_idx,
                    value=f"={c}21*{c}7")
            ws.cell(row=22, column=col_idx).number_format = '#,##0.0'

        # Row 23: Exit Equity = Exit EV - Debt at Exit (50% paydown simplified)
        ws.cell(row=23, column=1, value="Exit Equity Value")
        for col_idx in range(2, 5):
            c = get_column_letter(col_idx)
            ws.cell(row=23, column=col_idx,
                    value=f"={c}22-$B$13*0.5")
            ws.cell(row=23, column=col_idx).number_format = '#,##0.0'

        # Row 24: MOIC = Exit Equity / Equity Check
        ws.cell(row=24, column=1, value="MOIC")
        for col_idx in range(2, 5):
            c = get_column_letter(col_idx)
            ws.cell(row=24, column=col_idx,
                    value=f"=IFERROR({c}23/$B$15,0)")
            ws.cell(row=24, column=col_idx).number_format = '0.00"x"'
            ws.cell(row=24, column=col_idx).font = Font(bold=True)

        # Row 25: IRR = MOIC^(1/hold) - 1
        ws.cell(row=25, column=1, value="IRR")
        for col_idx in range(2, 5):
            c = get_column_letter(col_idx)
            ws.cell(row=25, column=col_idx,
                    value=f"=IFERROR({c}24^(1/$B$16)-1,0)")
            ws.cell(row=25, column=col_idx).number_format = '0.0%'
            ws.cell(row=25, column=col_idx).font = Font(bold=True)

        # -- SECTION 4: Probability-Weighted Returns (formulas) --
        self._add_section_header(ws, "Probability-Weighted Returns", 27, 1)

        # Row 28: Expected MOIC = SUMPRODUCT(MOIC row, Probability row)
        ws.cell(row=28, column=1, value="Expected MOIC")
        ws.cell(row=28, column=2, value="=SUMPRODUCT(B24:D24,B8:D8)")
        ws.cell(row=28, column=2).number_format = '0.00"x"'
        ws.cell(row=28, column=2).font = Font(bold=True, size=14)
        ws.cell(row=28, column=2).fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")

        # Row 29: Expected IRR = SUMPRODUCT(IRR row, Probability row)
        ws.cell(row=29, column=1, value="Expected IRR")
        ws.cell(row=29, column=2, value="=SUMPRODUCT(B25:D25,B8:D8)")
        ws.cell(row=29, column=2).number_format = '0.0%'
        ws.cell(row=29, column=2).font = Font(bold=True, size=14)
        ws.cell(row=29, column=2).fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")

        # -- SECTION 5: Downside Protection (formulas) --
        self._add_section_header(ws, "Downside Protection", 31, 1)

        ws.cell(row=32, column=1, value="Bear Case MOIC")
        ws.cell(row=32, column=2, value="=B24")
        ws.cell(row=32, column=2).number_format = '0.00"x"'

        ws.cell(row=33, column=1, value="Capital Protected?")
        ws.cell(row=33, column=2, value='=IF(B32>=1,"Yes","No")')

        # Note about live formulas
        ws.cell(row=35, column=1, value="Note: All outputs update automatically when scenario assumptions change")
        ws.cell(row=35, column=1).font = Font(italic=True, color="666666")

        self.cell_map['scenario_analysis'] = {
            'rev_growth_row': 5,
            'ebitda_margin_row': 6,
            'exit_multiple_row': 7,
            'probability_row': 8,
            'moic_row': 24,
            'irr_row': 25,
            'expected_moic_cell': 'B28',
            'expected_irr_cell': 'B29',
        }

        # Set column widths
        ws.column_dimensions['A'].width = 28
        for col in ['B', 'C', 'D']:
            ws.column_dimensions[col].width = 14

        return self

    # ============================================================
    # MODULE: MANAGEMENT VS BUYER CASE
    # ============================================================

    def add_management_vs_buyer(self, data: Dict = None) -> 'CoreModulesMixin':
        """Add Management Case vs Buyer Case comparison with live Excel formulas."""
        ws = self.wb.create_sheet("Mgmt vs Buyer Case")
        self.sheets_created.append("Mgmt vs Buyer Case")

        a = self.assumptions
        data = data or {}
        n = self.projection_years

        # Define case assumptions (inputs)
        mgmt_case = data.get('management', {
            'revenue_growth': [0.12, 0.10, 0.09, 0.08, 0.07],
            'ebitda_margin': [0.22, 0.24, 0.25, 0.26, 0.27],
        })
        buyer_case = data.get('buyer', {
            'revenue_growth': [0.08, 0.07, 0.06, 0.05, 0.05],
            'ebitda_margin': [0.20, 0.21, 0.22, 0.22, 0.22],
        })

        # Title
        self._add_title(ws, f"{self.company_name} - Management vs Buyer Case", 1, 1)

        # Year headers (row 3)
        ws.cell(row=3, column=1, value="")
        ws.cell(row=3, column=2, value="LTM")
        for i in range(n):
            ws.cell(row=3, column=3 + i, value=f"Year {i + 1}")
        self._format_header_row(ws, 3, 1, 2 + n)

        # -- MANAGEMENT CASE (rows 4-8) --
        self._add_section_header(ws, "MANAGEMENT CASE", 4, 1)

        # Row 5: Revenue — LTM from Assumptions, projections as formulas
        ws.cell(row=5, column=1, value="Revenue")
        ws.cell(row=5, column=2, value=f"='Assumptions'!B5")
        ws.cell(row=5, column=2).number_format = '#,##0.0'
        for i in range(n):
            col = get_column_letter(3 + i)
            prev = get_column_letter(2 + i)
            ws.cell(row=5, column=3 + i, value=f"={prev}5*(1+{col}6)")
            ws.cell(row=5, column=3 + i).number_format = '#,##0.0'

        # Row 6: Growth % (inputs)
        ws.cell(row=6, column=1, value="  % Growth")
        ws.cell(row=6, column=2, value="—")
        for i in range(n):
            g = mgmt_case['revenue_growth'][i] if i < len(mgmt_case['revenue_growth']) else 0.05
            ws.cell(row=6, column=3 + i, value=g)
            ws.cell(row=6, column=3 + i).number_format = '0.0%'
            self._format_input_cell(ws, 6, 3 + i)

        # Row 7: EBITDA = Revenue x Margin
        ws.cell(row=7, column=1, value="EBITDA")
        ws.cell(row=7, column=2, value=f"='Assumptions'!B6")
        ws.cell(row=7, column=2).number_format = '#,##0.0'
        ws.cell(row=7, column=2).font = Font(bold=True)
        for i in range(n):
            col = get_column_letter(3 + i)
            ws.cell(row=7, column=3 + i, value=f"={col}5*{col}8")
            ws.cell(row=7, column=3 + i).number_format = '#,##0.0'
            ws.cell(row=7, column=3 + i).font = Font(bold=True)

        # Row 8: Margin % (inputs)
        ws.cell(row=8, column=1, value="  % Margin")
        ws.cell(row=8, column=2, value="=B7/B5")
        ws.cell(row=8, column=2).number_format = '0.0%'
        for i in range(n):
            m = mgmt_case['ebitda_margin'][i] if i < len(mgmt_case['ebitda_margin']) else 0.22
            ws.cell(row=8, column=3 + i, value=m)
            ws.cell(row=8, column=3 + i).number_format = '0.0%'
            self._format_input_cell(ws, 8, 3 + i)

        # -- BUYER CASE (rows 10-14) --
        self._add_section_header(ws, "BUYER CASE (HAIRCUT)", 10, 1)

        # Row 11: Revenue
        ws.cell(row=11, column=1, value="Revenue")
        ws.cell(row=11, column=2, value="=B5")  # Same LTM
        ws.cell(row=11, column=2).number_format = '#,##0.0'
        for i in range(n):
            col = get_column_letter(3 + i)
            prev = get_column_letter(2 + i)
            ws.cell(row=11, column=3 + i, value=f"={prev}11*(1+{col}12)")
            ws.cell(row=11, column=3 + i).number_format = '#,##0.0'

        # Row 12: Growth % (inputs)
        ws.cell(row=12, column=1, value="  % Growth")
        ws.cell(row=12, column=2, value="—")
        for i in range(n):
            g = buyer_case['revenue_growth'][i] if i < len(buyer_case['revenue_growth']) else 0.05
            ws.cell(row=12, column=3 + i, value=g)
            ws.cell(row=12, column=3 + i).number_format = '0.0%'
            self._format_input_cell(ws, 12, 3 + i)

        # Row 13: EBITDA = Revenue x Margin
        ws.cell(row=13, column=1, value="EBITDA")
        ws.cell(row=13, column=2, value="=B11*B14")
        ws.cell(row=13, column=2).number_format = '#,##0.0'
        ws.cell(row=13, column=2).font = Font(bold=True)
        for i in range(n):
            col = get_column_letter(3 + i)
            ws.cell(row=13, column=3 + i, value=f"={col}11*{col}14")
            ws.cell(row=13, column=3 + i).number_format = '#,##0.0'
            ws.cell(row=13, column=3 + i).font = Font(bold=True)

        # Row 14: Margin % (inputs)
        ws.cell(row=14, column=1, value="  % Margin")
        ws.cell(row=14, column=2, value="=B7/B5")
        ws.cell(row=14, column=2).number_format = '0.0%'
        for i in range(n):
            m = buyer_case['ebitda_margin'][i] if i < len(buyer_case['ebitda_margin']) else 0.20
            ws.cell(row=14, column=3 + i, value=m)
            ws.cell(row=14, column=3 + i).number_format = '0.0%'
            self._format_input_cell(ws, 14, 3 + i)

        # -- VARIANCE ANALYSIS (rows 16-21) --
        self._add_section_header(ws, "VARIANCE ANALYSIS", 16, 1)

        # Row 17: Revenue Variance = Buyer - Mgmt
        ws.cell(row=17, column=1, value="Revenue Variance")
        for i in range(n + 1):
            col = get_column_letter(2 + i)
            ws.cell(row=17, column=2 + i, value=f"={col}11-{col}5")
            ws.cell(row=17, column=2 + i).number_format = '#,##0.0'

        # Row 18: % Variance
        ws.cell(row=18, column=1, value="  % Variance")
        for i in range(n + 1):
            col = get_column_letter(2 + i)
            ws.cell(row=18, column=2 + i, value=f"=IFERROR({col}17/{col}5,0)")
            ws.cell(row=18, column=2 + i).number_format = '0.0%'

        # Row 20: EBITDA Variance
        ws.cell(row=20, column=1, value="EBITDA Variance")
        for i in range(n + 1):
            col = get_column_letter(2 + i)
            ws.cell(row=20, column=2 + i, value=f"={col}13-{col}7")
            ws.cell(row=20, column=2 + i).number_format = '#,##0.0'

        # Row 21: % Variance
        ws.cell(row=21, column=1, value="  % Variance")
        for i in range(n + 1):
            col = get_column_letter(2 + i)
            ws.cell(row=21, column=2 + i, value=f"=IFERROR({col}20/{col}7,0)")
            ws.cell(row=21, column=2 + i).number_format = '0.0%'

        # -- RETURNS COMPARISON (rows 23-28) --
        self._add_section_header(ws, "RETURNS COMPARISON", 23, 1)

        ws.cell(row=24, column=1, value="")
        ws.cell(row=24, column=2, value="Management")
        ws.cell(row=24, column=3, value="Buyer")
        ws.cell(row=24, column=4, value="Difference")
        self._format_header_row(ws, 24, 1, 4)

        # Exit year column (last projection year)
        exit_col = get_column_letter(2 + n)

        # Row 25: Exit EBITDA — reference last year of each case
        ws.cell(row=25, column=1, value="Exit EBITDA")
        ws.cell(row=25, column=2, value=f"={exit_col}7")
        ws.cell(row=25, column=3, value=f"={exit_col}13")
        ws.cell(row=25, column=4, value="=C25-B25")
        for col in range(2, 5):
            ws.cell(row=25, column=col).number_format = '#,##0.0'

        # Row 26: Exit EV = Exit EBITDA x Exit Multiple
        ws.cell(row=26, column=1, value="Exit EV")
        ws.cell(row=26, column=2, value=f"=B25*'Assumptions'!B8")
        ws.cell(row=26, column=3, value=f"=C25*'Assumptions'!B8")
        ws.cell(row=26, column=4, value="=C26-B26")
        for col in range(2, 5):
            ws.cell(row=26, column=col).number_format = '#,##0.0'

        # Row 27: Exit Equity = Exit EV - Debt at exit (50% paydown)
        su = self.cell_map.get('sources_uses', {})
        su_debt = su.get('total_debt', 'B8')  # fallback
        ws.cell(row=27, column=1, value="Exit Equity")
        ws.cell(row=27, column=2, value=f"=B26-'Sources & Uses'!{su_debt}*0.5")
        ws.cell(row=27, column=3, value=f"=C26-'Sources & Uses'!{su_debt}*0.5")
        ws.cell(row=27, column=4, value="=C27-B27")
        for col in range(2, 5):
            ws.cell(row=27, column=col).number_format = '#,##0.0'

        # Row 28: MOIC = Exit Equity / Initial Equity
        su_equity = su.get('sponsor_equity', 'B13')
        ws.cell(row=28, column=1, value="MOIC")
        ws.cell(row=28, column=2, value=f"=IFERROR(B27/'Sources & Uses'!{su_equity},0)")
        ws.cell(row=28, column=3, value=f"=IFERROR(C27/'Sources & Uses'!{su_equity},0)")
        ws.cell(row=28, column=4, value="=C28-B28")
        for col in range(2, 5):
            ws.cell(row=28, column=col).number_format = '0.00x'
            ws.cell(row=28, column=col).font = Font(bold=True)

        # Row 29: IRR = MOIC^(1/hold) - 1
        ws.cell(row=29, column=1, value="IRR")
        ws.cell(row=29, column=2, value=f"=IFERROR(B28^(1/'Assumptions'!B9)-1,0)")
        ws.cell(row=29, column=3, value=f"=IFERROR(C28^(1/'Assumptions'!B9)-1,0)")
        ws.cell(row=29, column=4, value="=C29-B29")
        for col in range(2, 5):
            ws.cell(row=29, column=col).number_format = '0.0%'
            ws.cell(row=29, column=col).font = Font(bold=True)

        # Highlight buyer case column
        for row in range(25, 30):
            ws.cell(row=row, column=3).fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")

        # Set column widths
        ws.column_dimensions['A'].width = 20
        for i in range(n + 3):
            ws.column_dimensions[get_column_letter(2 + i)].width = 12

        self.cell_map['mgmt_vs_buyer'] = {
            'mgmt_revenue_row': 5,
            'mgmt_growth_row': 6,
            'mgmt_ebitda_row': 7,
            'mgmt_margin_row': 8,
            'buyer_revenue_row': 11,
            'buyer_growth_row': 12,
            'buyer_ebitda_row': 13,
            'buyer_margin_row': 14,
            'moic_row': 28,
            'irr_row': 29,
        }

        return self

    # ============================================================
    # MODULE: DCF VALUATION
    # ============================================================

    def add_dcf_valuation(self, data: Dict = None) -> 'CoreModulesMixin':
        """Add DCF Valuation sheet with real Excel formulas."""
        ws = self.wb.create_sheet("DCF Valuation")
        self.sheets_created.append("DCF Valuation")

        A = "Assumptions"
        OM = "Operating Model"
        n = self.projection_years
        data = data or {}

        # WACC and terminal growth are local DCF inputs (not on Assumptions sheet)
        wacc = data.get('wacc', 0.10)
        terminal_growth = data.get('terminal_growth', 0.025)

        # Title
        self._add_title(ws, f"{self.company_name} - DCF Valuation", 1, 1)

        # DCF Assumptions (rows 3-6)
        self._add_section_header(ws, "DCF Assumptions", 3, 1)

        ws.cell(row=4, column=1, value="WACC")
        ws.cell(row=4, column=2, value=wacc)
        ws.cell(row=4, column=2).number_format = '0.0%'
        self._format_input_cell(ws, 4, 2)

        ws.cell(row=5, column=1, value="Terminal Growth Rate")
        ws.cell(row=5, column=2, value=terminal_growth)
        ws.cell(row=5, column=2).number_format = '0.0%'
        self._format_input_cell(ws, 5, 2)

        ws.cell(row=6, column=1, value="Tax Rate")
        ws.cell(row=6, column=2, value=f"='{A}'!B24")
        ws.cell(row=6, column=2).number_format = '0.0%'
        self._format_input_cell(ws, 6, 2)

        # Year headers (row 8)
        ws.cell(row=8, column=1, value="")
        for i in range(n):
            ws.cell(row=8, column=2 + i, value=f"Year {i + 1}")
        ws.cell(row=8, column=2 + n, value="Terminal")
        self._format_header_row(ws, 8, 1, 2 + n)

        # DCF column layout: B=Year1, C=Year2, ... last_col+1=Terminal
        # Year i (1-indexed) is in column (1+i) = col 2+i-1... actually col 2+i-1
        # Wait: Year 1 = col B (2), Year 2 = col C (3), ... Year n = col 1+n
        # But Operating Model has: B=LTM, C=Y1, D=Y2, ... so Y(i) = col 2+i
        # DCF has: B=Y1, C=Y2, ... so Y(i) = col 1+i
        # We need to map: DCF col for year i = 1+i, OM col for year i = 2+i

        # --- Unlevered Free Cash Flow Build (rows 9-18) ---
        self._add_section_header(ws, "Unlevered Free Cash Flow ($M)", 9, 1)

        # Row 10: EBITDA — from Operating Model
        ws.cell(row=10, column=1, value="EBITDA")
        for i in range(n):
            dcf_col = 2 + i  # B, C, D, ...
            om_col_letter = get_column_letter(3 + i)  # C, D, E, ... (Y1+ in OM)
            ws.cell(row=10, column=dcf_col, value=f"='{OM}'!{om_col_letter}13")
            ws.cell(row=10, column=dcf_col).number_format = '#,##0.0'

        # Row 11: Less: D&A = CapEx from Operating Model (negative)
        ws.cell(row=11, column=1, value="Less: D&A")
        for i in range(n):
            dcf_col = 2 + i
            om_col_letter = get_column_letter(3 + i)
            ws.cell(row=11, column=dcf_col, value=f"='{OM}'!{om_col_letter}16")  # CapEx is already negative
            ws.cell(row=11, column=dcf_col).number_format = '(#,##0.0)'

        # Row 12: EBIT = EBITDA + D&A (D&A is negative)
        ws.cell(row=12, column=1, value="EBIT")
        for i in range(n):
            dcf_col = 2 + i
            c = get_column_letter(dcf_col)
            ws.cell(row=12, column=dcf_col, value=f"={c}10+{c}11")
            ws.cell(row=12, column=dcf_col).number_format = '#,##0.0'

        # Row 13: Less: Taxes = -EBIT x Tax Rate
        ws.cell(row=13, column=1, value="Less: Taxes")
        for i in range(n):
            dcf_col = 2 + i
            c = get_column_letter(dcf_col)
            ws.cell(row=13, column=dcf_col, value=f"=-{c}12*$B$6")
            ws.cell(row=13, column=dcf_col).number_format = '(#,##0.0)'

        # Row 14: NOPAT = EBIT x (1 - Tax Rate)
        ws.cell(row=14, column=1, value="NOPAT")
        for i in range(n):
            dcf_col = 2 + i
            c = get_column_letter(dcf_col)
            ws.cell(row=14, column=dcf_col, value=f"={c}12*(1-$B$6)")
            ws.cell(row=14, column=dcf_col).number_format = '#,##0.0'

        # Row 15: Plus: D&A = -Row 11 (add back)
        ws.cell(row=15, column=1, value="Plus: D&A")
        for i in range(n):
            dcf_col = 2 + i
            c = get_column_letter(dcf_col)
            ws.cell(row=15, column=dcf_col, value=f"=-{c}11")
            ws.cell(row=15, column=dcf_col).number_format = '#,##0.0'

        # Row 16: Less: CapEx (same as D&A, negative)
        ws.cell(row=16, column=1, value="Less: CapEx")
        for i in range(n):
            dcf_col = 2 + i
            c = get_column_letter(dcf_col)
            ws.cell(row=16, column=dcf_col, value=f"={c}11")  # same as D&A (already negative)
            ws.cell(row=16, column=dcf_col).number_format = '(#,##0.0)'

        # Row 17: Less: Change in NWC
        # NWC = Revenue x NWC%, Change = -(this_NWC - prev_NWC)
        ws.cell(row=17, column=1, value="Less: Change in NWC")
        for i in range(n):
            dcf_col = 2 + i
            om_this = get_column_letter(3 + i)   # Y(i+1) in OM
            om_prev = get_column_letter(2 + i)   # Y(i) in OM (or LTM for i=0)
            ws.cell(row=17, column=dcf_col,
                    value=f"=-('{OM}'!{om_this}10*'{A}'!B23-'{OM}'!{om_prev}10*'{A}'!B23)")
            ws.cell(row=17, column=dcf_col).number_format = '(#,##0.0)'

        # Row 18: Unlevered FCF = NOPAT + D&A + CapEx + Change in NWC
        # Since D&A and CapEx cancel, UFCF = NOPAT + Change_NWC(row 17)
        ws.cell(row=18, column=1, value="Unlevered FCF")
        for i in range(n):
            dcf_col = 2 + i
            c = get_column_letter(dcf_col)
            ws.cell(row=18, column=dcf_col, value=f"={c}14+{c}15+{c}16+{c}17")
            ws.cell(row=18, column=dcf_col).number_format = '#,##0.0'
            ws.cell(row=18, column=dcf_col).font = Font(bold=True)
            ws.cell(row=18, column=dcf_col).border = DOUBLE_BORDER

        # Terminal Value = Last UFCF x (1+g) / (WACC - g)
        term_col = 2 + n  # Terminal column
        last_fcf_col = get_column_letter(1 + n)  # last year UFCF column
        ws.cell(row=18, column=term_col,
                value=f"={last_fcf_col}18*(1+$B$5)/($B$4-$B$5)")
        ws.cell(row=18, column=term_col).number_format = '#,##0.0'
        ws.cell(row=18, column=term_col).font = Font(bold=True)

        # --- Present Value Calculation (rows 20-22) ---
        self._add_section_header(ws, "Present Value Calculation", 20, 1)

        # Row 21: Discount Factor = 1/(1+WACC)^year
        ws.cell(row=21, column=1, value="Discount Factor")
        for i in range(n):
            dcf_col = 2 + i
            ws.cell(row=21, column=dcf_col, value=f"=1/(1+$B$4)^{i + 1}")
            ws.cell(row=21, column=dcf_col).number_format = '0.000'

        # Terminal discount factor (same as last year)
        ws.cell(row=21, column=term_col, value=f"=1/(1+$B$4)^{n}")
        ws.cell(row=21, column=term_col).number_format = '0.000'

        # Row 22: Present Value = UFCF x Discount Factor
        ws.cell(row=22, column=1, value="Present Value")
        for i in range(n):
            dcf_col = 2 + i
            c = get_column_letter(dcf_col)
            ws.cell(row=22, column=dcf_col, value=f"={c}18*{c}21")
            ws.cell(row=22, column=dcf_col).number_format = '#,##0.0'

        # PV of Terminal Value
        tc = get_column_letter(term_col)
        ws.cell(row=22, column=term_col, value=f"={tc}18*{tc}21")
        ws.cell(row=22, column=term_col).number_format = '#,##0.0'

        # --- Valuation Summary (rows 24-31) ---
        self._add_section_header(ws, "Valuation Summary ($M)", 24, 1)

        # PV of Projection FCF = SUM of PV row (excluding terminal)
        first_pv = get_column_letter(2)
        last_pv = get_column_letter(1 + n)
        ws.cell(row=25, column=1, value="PV of Projection Period FCF")
        ws.cell(row=25, column=2, value=f"=SUM({first_pv}22:{last_pv}22)")
        ws.cell(row=25, column=2).number_format = '#,##0.0'

        # PV of Terminal Value
        ws.cell(row=26, column=1, value="PV of Terminal Value")
        ws.cell(row=26, column=2, value=f"={tc}22")
        ws.cell(row=26, column=2).number_format = '#,##0.0'

        # Enterprise Value = sum
        ws.cell(row=27, column=1, value="Enterprise Value")
        ws.cell(row=27, column=2, value="=B25+B26")
        ws.cell(row=27, column=2).number_format = '#,##0.0'
        ws.cell(row=27, column=2).font = Font(bold=True, size=14)
        ws.cell(row=27, column=2).fill = PatternFill(
            start_color="90EE90", end_color="90EE90", fill_type="solid")

        # Implied multiples
        ws.cell(row=29, column=1, value="Implied EV / LTM EBITDA")
        ws.cell(row=29, column=2, value=f"=IFERROR(B27/'{A}'!B6,0)")
        ws.cell(row=29, column=2).number_format = '0.0"x"'

        # Exit EBITDA = last year in OM
        exit_om_col = get_column_letter(2 + n)
        ws.cell(row=30, column=1, value="Implied EV / Exit EBITDA")
        ws.cell(row=30, column=2, value=f"=IFERROR(B27/'{OM}'!{exit_om_col}13,0)")
        ws.cell(row=30, column=2).number_format = '0.0"x"'

        # Terminal Value % of EV
        ws.cell(row=31, column=1, value="Terminal Value % of EV")
        ws.cell(row=31, column=2, value="=IFERROR(B26/B27,0)")
        ws.cell(row=31, column=2).number_format = '0.0%'

        # Set column widths
        ws.column_dimensions['A'].width = 25
        for i in range(n + 2):
            ws.column_dimensions[get_column_letter(2 + i)].width = 12

        return self

    # ============================================================
    # MODULE: COVENANT ANALYSIS
    # ============================================================

    def add_covenant_analysis(self, data: Dict = None) -> 'CoreModulesMixin':
        """Add Covenant Analysis sheet with live formulas referencing OM and Debt Schedule."""
        ws = self.wb.create_sheet("Covenant Analysis")
        self.sheets_created.append("Covenant Analysis")

        A = "Assumptions"
        OM = "Operating Model"
        DS = "Debt Schedule"
        a = self.assumptions
        n = self.projection_years
        data = data or {}

        # Resolve Debt Schedule cell_map (quick vs standard mode)
        ds_map = self.cell_map.get('debt_schedule', {})
        ds_mode = ds_map.get('mode', 'quick')
        om_map = self.cell_map.get('operating_model', {})
        om_ebitda_row = om_map.get('ebitda_row', 13)

        # Debt ending-balance row in DS sheet
        if ds_mode == 'standard':
            # Standard mode: total debt row = senior_ending + sub_ending via a total row
            ds_total_debt_row = ds_map.get('total_debt_row', 20)
            ds_interest_row = ds_map.get('total_interest_row', 22)
        else:
            # Quick mode
            ds_total_debt_row = ds_map.get('ending_balance_row', 12)
            ds_interest_row = ds_map.get('interest_row', 14)

        # Covenant thresholds (input cells — user-editable)
        covenants = data.get('covenants', {
            'max_leverage': 6.0,
            'min_interest_coverage': 2.0,
            'min_fixed_charge': 1.1,
        })

        # Title
        self._add_title(ws, f"{self.company_name} - Covenant Analysis", 1, 1)

        # -- Covenant Thresholds (input values) --
        self._add_section_header(ws, "Covenant Thresholds", 3, 1)

        ws.cell(row=4, column=1, value="Maximum Leverage (Debt/EBITDA)")
        ws.cell(row=4, column=2, value=covenants['max_leverage'])
        ws.cell(row=4, column=2).number_format = '0.0"x"'
        self._format_input_cell(ws, 4, 2)

        ws.cell(row=5, column=1, value="Minimum Interest Coverage")
        ws.cell(row=5, column=2, value=covenants['min_interest_coverage'])
        ws.cell(row=5, column=2).number_format = '0.0"x"'
        self._format_input_cell(ws, 5, 2)

        ws.cell(row=6, column=1, value="Minimum Fixed Charge Coverage")
        ws.cell(row=6, column=2, value=covenants['min_fixed_charge'])
        ws.cell(row=6, column=2).number_format = '0.0"x"'
        self._format_input_cell(ws, 6, 2)

        # -- Year headers --
        ws.cell(row=8, column=1, value="")
        ws.cell(row=8, column=2, value="Entry")
        for i in range(n):
            ws.cell(row=8, column=3 + i, value=f"Year {i + 1}")
        self._format_header_row(ws, 8, 1, 2 + n)

        # -- Leverage Ratio (formulas referencing DS and OM) --
        self._add_section_header(ws, "Leverage Ratio (Debt / EBITDA)", 9, 1)

        # Row 10: Total Debt — pull from Debt Schedule ending balance
        ws.cell(row=10, column=1, value="Total Debt")
        for i in range(n + 1):
            col = get_column_letter(2 + i)
            ws.cell(row=10, column=2 + i,
                    value=f"='{DS}'!{col}{ds_total_debt_row}")
            ws.cell(row=10, column=2 + i).number_format = '#,##0.0'

        # Row 11: EBITDA — pull from Operating Model
        ws.cell(row=11, column=1, value="EBITDA")
        for i in range(n + 1):
            col = get_column_letter(2 + i)
            ws.cell(row=11, column=2 + i,
                    value=f"='{OM}'!{col}{om_ebitda_row}")
            ws.cell(row=11, column=2 + i).number_format = '#,##0.0'

        # Row 12: Leverage Ratio = Debt / EBITDA
        ws.cell(row=12, column=1, value="Leverage Ratio")
        for i in range(n + 1):
            col = get_column_letter(2 + i)
            ws.cell(row=12, column=2 + i,
                    value=f"=IFERROR({col}10/{col}11,0)")
            ws.cell(row=12, column=2 + i).number_format = '0.0"x"'
            ws.cell(row=12, column=2 + i).font = Font(bold=True)

        # Row 13: Covenant threshold (repeated for comparison)
        ws.cell(row=13, column=1, value="Covenant")
        for i in range(n + 1):
            ws.cell(row=13, column=2 + i, value="=$B$4")
            ws.cell(row=13, column=2 + i).number_format = '0.0"x"'
            ws.cell(row=13, column=2 + i).font = Font(italic=True, color="666666")

        # Row 14: Headroom = Covenant - Leverage
        ws.cell(row=14, column=1, value="Headroom")
        for i in range(n + 1):
            col = get_column_letter(2 + i)
            ws.cell(row=14, column=2 + i,
                    value=f"=$B$4-{col}12")
            ws.cell(row=14, column=2 + i).number_format = '0.0"x"'

        # -- Interest Coverage (formulas) --
        self._add_section_header(ws, "Interest Coverage (EBITDA / Interest)", 16, 1)

        # Row 17: EBITDA (projected years only, cols C+)
        ws.cell(row=17, column=1, value="EBITDA")
        for i in range(n):
            col = get_column_letter(3 + i)
            ws.cell(row=17, column=3 + i,
                    value=f"='{OM}'!{col}{om_ebitda_row}")
            ws.cell(row=17, column=3 + i).number_format = '#,##0.0'

        # Row 18: Interest Expense — pull from Debt Schedule interest row
        ws.cell(row=18, column=1, value="Interest Expense")
        for i in range(n):
            col = get_column_letter(3 + i)
            ws.cell(row=18, column=3 + i,
                    value=f"='{DS}'!{col}{ds_interest_row}")
            ws.cell(row=18, column=3 + i).number_format = '#,##0.0'

        # Row 19: Interest Coverage = EBITDA / Interest
        ws.cell(row=19, column=1, value="Interest Coverage")
        for i in range(n):
            col = get_column_letter(3 + i)
            ws.cell(row=19, column=3 + i,
                    value=f"=IFERROR({col}17/{col}18,99)")
            ws.cell(row=19, column=3 + i).number_format = '0.0"x"'
            ws.cell(row=19, column=3 + i).font = Font(bold=True)

        # Row 20: Covenant threshold
        ws.cell(row=20, column=1, value="Covenant")
        for i in range(n):
            ws.cell(row=20, column=3 + i, value="=$B$5")
            ws.cell(row=20, column=3 + i).number_format = '0.0"x"'
            ws.cell(row=20, column=3 + i).font = Font(italic=True, color="666666")

        # -- Compliance Summary (formulas) --
        self._add_section_header(ws, "Compliance Summary", 22, 1)

        # Row 23: Overall Compliance — AND of all leverage <= max AND coverage >= min
        # Build formula: check last year leverage and first year coverage as proxy
        last_yr_col = get_column_letter(2 + n)
        first_yr_col = get_column_letter(3)
        ws.cell(row=23, column=1, value="Overall Compliance")
        ws.cell(row=23, column=2,
                value=f'=IF(AND({last_yr_col}12<=$B$4,{first_yr_col}19>=$B$5),"PASS","FAIL")')
        ws.cell(row=23, column=2).font = Font(bold=True, size=14)

        # Note
        ws.cell(row=25, column=1,
                value="Note: Ratios update automatically from Operating Model and Debt Schedule")
        ws.cell(row=25, column=1).font = Font(italic=True, color="666666")

        # --- Conditional Formatting (RAG) ---
        # Leverage: lower is better — green if <= covenant, red if > covenant
        lev_range = f"B12:{get_column_letter(2 + n)}12"
        self._add_rag_cells(ws, lev_range,
                            green_threshold='$B$4',
                            amber_threshold=None,
                            higher_is_better=False)
        # Headroom: higher is better — green if > 0, red if < 0
        head_range = f"B14:{get_column_letter(2 + n)}14"
        self._add_rag_cells(ws, head_range,
                            green_threshold=0,
                            higher_is_better=True)
        # Coverage: higher is better — green if >= covenant
        cov_range = f"C19:{get_column_letter(2 + n)}19"
        self._add_rag_cells(ws, cov_range,
                            green_threshold='$B$5',
                            higher_is_better=True)
        # PASS/FAIL cell
        from openpyxl.formatting.rule import FormulaRule
        ws.conditional_formatting.add(
            'B23',
            FormulaRule(formula=['B23="PASS"'],
                        fill=PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')))
        ws.conditional_formatting.add(
            'B23',
            FormulaRule(formula=['B23="FAIL"'],
                        fill=PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')))

        self.cell_map['covenant_analysis'] = {
            'max_leverage_cell': 'B4',
            'min_coverage_cell': 'B5',
            'leverage_row': 12,
            'coverage_row': 19,
            'compliance_cell': 'B23',
        }

        # Set column widths
        ws.column_dimensions['A'].width = 28
        for i in range(n + 2):
            ws.column_dimensions[get_column_letter(2 + i)].width = 12

        return self
