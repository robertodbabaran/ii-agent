#!/usr/bin/env python3
"""
Due Diligence Excel Modules Mixin.

Contains: Quality of Earnings, NWC Normalization,
Customer Revenue Quality, Credit/Debt Sizing.

Smart cross-sheet formula upgrade: When upstream modules (Operating Model,
Sources & Uses, Debt Schedule) exist in cell_map, these modules write live
cross-sheet Excel formulas. When upstream modules are absent, hardcoded
fallback values are used instead.
"""

from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from .formula_builder import FormulaBuilder as FB
from ._base import (INPUT_FILL, INPUT_FONT, HEADER_FILL, HEADER_FONT,
                     BOTTOM_BORDER, DOUBLE_BORDER, ModelDepth)
from typing import Dict


class DDModulesMixin:
    """Mixin providing due diligence analysis modules."""

    # ============================================================
    # MODULE: QUALITY OF EARNINGS (QoE)
    # ============================================================

    def add_quality_of_earnings(self, data: Dict = None) -> 'DDModulesMixin':
        """
        Add Quality of Earnings analysis with live Excel formulas for totals.

        Smart cross-ref: When Operating Model exists in cell_map, Row 5
        (Reported EBITDA LTM) uses a cross-sheet formula to pull EBITDA
        from the Operating Model's LTM column. Falls back to hardcoded
        values when OM is not built.
        """
        ws = self.wb.create_sheet("Quality of Earnings")
        self.sheets_created.append("Quality of Earnings")

        a = self.assumptions
        data = data or {}

        # Check if Operating Model has been built
        om = self.cell_map.get('operating_model', {})

        # Title
        self._add_title(ws, f"{self.company_name} - Quality of Earnings Analysis", 1, 1)

        # Reported EBITDA Section
        self._add_section_header(ws, "REPORTED EBITDA ($M)", 3, 1)

        ws.cell(row=4, column=1, value="")
        ws.cell(row=4, column=2, value="LTM")
        ws.cell(row=4, column=3, value="FY-1")
        ws.cell(row=4, column=4, value="FY-2")
        self._format_header_row(ws, 4, 1, 4)

        # Row 5: Reported figures
        # LTM column: cross-ref Operating Model EBITDA if available
        if om and 'ebitda_row' in om and 'ltm_col' in om:
            om_ltm_col_letter = get_column_letter(om['ltm_col'])
            om_ebitda_row = om['ebitda_row']
            reported_ltm_value = f"='Operating Model'!{om_ltm_col_letter}{om_ebitda_row}"
        else:
            reported_ltm_value = data.get('reported_ebitda_ltm', a.ltm_ebitda)

        # FY-1 and FY-2 are always inputs (historical, no upstream source)
        reported_fy1 = data.get('reported_ebitda_fy1', a.ltm_ebitda * 0.92)
        reported_fy2 = data.get('reported_ebitda_fy2', a.ltm_ebitda * 0.85)

        ws.cell(row=5, column=1, value="Reported EBITDA")
        ws.cell(row=5, column=2, value=reported_ltm_value)
        ws.cell(row=5, column=3, value=reported_fy1)
        ws.cell(row=5, column=4, value=reported_fy2)
        for col in range(2, 5):
            ws.cell(row=5, column=col).number_format = '#,##0.0'
            ws.cell(row=5, column=col).font = Font(bold=True)
            self._format_input_cell(ws, 5, col)

        # Adjustments Section
        self._add_section_header(ws, "EBITDA ADJUSTMENTS", 7, 1)

        adjustments = data.get('adjustments', [
            {'name': 'Owner compensation normalization', 'ltm': 2.0, 'fy1': 1.8, 'fy2': 1.5},
            {'name': 'One-time legal/settlement costs', 'ltm': 1.5, 'fy1': 0.5, 'fy2': 2.0},
            {'name': 'Non-recurring consulting fees', 'ltm': 0.8, 'fy1': 1.2, 'fy2': 0.3},
            {'name': 'Related party transaction adjustment', 'ltm': 0.5, 'fy1': 0.5, 'fy2': 0.5},
            {'name': 'Inventory write-down (non-recurring)', 'ltm': 0.0, 'fy1': 1.0, 'fy2': 0.0},
            {'name': 'Stock-based compensation', 'ltm': 1.2, 'fy1': 1.0, 'fy2': 0.8},
            {'name': 'Transaction costs (add-back)', 'ltm': 0.5, 'fy1': 0.0, 'fy2': 0.0},
            {'name': 'Cost savings (run-rate)', 'ltm': 1.5, 'fy1': 0.0, 'fy2': 0.0},
        ])

        # Rows 8+: Adjustment line items (inputs)
        adj_start_row = 8
        for i, adj in enumerate(adjustments):
            row = adj_start_row + i
            ws.cell(row=row, column=1, value=f"  {adj['name']}")
            ws.cell(row=row, column=2, value=adj['ltm'])
            ws.cell(row=row, column=3, value=adj['fy1'])
            ws.cell(row=row, column=4, value=adj['fy2'])
            for col in range(2, 5):
                ws.cell(row=row, column=col).number_format = '#,##0.0'
                ws.cell(row=row, column=col).font = Font(color="0066CC")
        adj_end_row = adj_start_row + len(adjustments) - 1

        # Total Adjustments row (SUM formula)
        total_row = adj_end_row + 2
        ws.cell(row=total_row, column=1, value="Total Adjustments")
        for col in range(2, 5):
            col_letter = get_column_letter(col)
            ws.cell(row=total_row, column=col,
                    value=f"=SUM({col_letter}{adj_start_row}:{col_letter}{adj_end_row})")
            ws.cell(row=total_row, column=col).number_format = '#,##0.0'
            ws.cell(row=total_row, column=col).font = Font(bold=True)
            ws.cell(row=total_row, column=col).border = Border(top=Side(style='thin'))

        # Adjusted EBITDA section
        adj_header_row = total_row + 2
        self._add_section_header(ws, "ADJUSTED EBITDA", adj_header_row, 1)
        adj_ebitda_row = adj_header_row + 1

        ws.cell(row=adj_ebitda_row, column=1, value="Adjusted EBITDA")
        for col in range(2, 5):
            col_letter = get_column_letter(col)
            ws.cell(row=adj_ebitda_row, column=col,
                    value=f"={col_letter}5+{col_letter}{total_row}")
            ws.cell(row=adj_ebitda_row, column=col).number_format = '#,##0.0'
            ws.cell(row=adj_ebitda_row, column=col).font = Font(bold=True)
            ws.cell(row=adj_ebitda_row, column=col).fill = PatternFill(
                start_color="90EE90", end_color="90EE90", fill_type="solid")

        # % Adjustment
        pct_row = adj_ebitda_row + 1
        ws.cell(row=pct_row, column=1, value="  % Adjustment")
        for col in range(2, 5):
            col_letter = get_column_letter(col)
            ws.cell(row=pct_row, column=col,
                    value=f"=IFERROR({col_letter}{total_row}/{col_letter}5,0)")
            ws.cell(row=pct_row, column=col).number_format = '0.0%'
            ws.cell(row=pct_row, column=col).font = Font(italic=True, color="666666")

        # Run-Rate Analysis
        rr_header_row = pct_row + 2
        self._add_section_header(ws, "RUN-RATE EBITDA ANALYSIS", rr_header_row, 1)
        rr_start = rr_header_row + 1

        # Base = Adjusted LTM EBITDA (formula reference)
        ws.cell(row=rr_start, column=1, value="Adjusted LTM EBITDA")
        ws.cell(row=rr_start, column=2, value=f"=B{adj_ebitda_row}")
        ws.cell(row=rr_start, column=2).number_format = '#,##0.0'

        run_rate_adds = data.get('run_rate_adds', [
            {'name': 'Full-year impact of price increase', 'value': 1.5},
            {'name': 'Annualized new customer wins', 'value': 2.0},
            {'name': 'Full-year cost savings', 'value': 1.0},
            {'name': 'Lost customer annualization', 'value': -0.5},
        ])

        for i, item in enumerate(run_rate_adds):
            row = rr_start + 1 + i
            ws.cell(row=row, column=1, value=f"  {item['name']}")
            ws.cell(row=row, column=2, value=item['value'])
            ws.cell(row=row, column=2).number_format = '#,##0.0'
            self._format_input_cell(ws, row, 2)
        rr_end = rr_start + len(run_rate_adds)

        # Run-Rate EBITDA total (SUM formula)
        rr_total_row = rr_end + 2
        ws.cell(row=rr_total_row, column=1, value="Run-Rate EBITDA")
        ws.cell(row=rr_total_row, column=2,
                value=f"=SUM(B{rr_start}:B{rr_end})")
        ws.cell(row=rr_total_row, column=2).number_format = '#,##0.0'
        ws.cell(row=rr_total_row, column=2).font = Font(bold=True)
        ws.cell(row=rr_total_row, column=2).fill = PatternFill(
            start_color="FFFF99", end_color="FFFF99", fill_type="solid")

        # QoE Summary
        sum_header = rr_total_row + 3
        self._add_section_header(ws, "QoE SUMMARY", sum_header, 1)
        sr = sum_header + 1

        ws.cell(row=sr, column=1, value="Reported LTM EBITDA")
        ws.cell(row=sr, column=2, value="=B5")
        ws.cell(row=sr, column=2).number_format = '#,##0.0'

        ws.cell(row=sr + 1, column=1, value="Adjusted LTM EBITDA")
        ws.cell(row=sr + 1, column=2, value=f"=B{adj_ebitda_row}")
        ws.cell(row=sr + 1, column=2).number_format = '#,##0.0'

        ws.cell(row=sr + 2, column=1, value="Run-Rate EBITDA")
        ws.cell(row=sr + 2, column=2, value=f"=B{rr_total_row}")
        ws.cell(row=sr + 2, column=2).number_format = '#,##0.0'

        ws.cell(row=sr + 3, column=1, value="Implied Adjustment %")
        ws.cell(row=sr + 3, column=2, value=f"=IFERROR((B{rr_total_row}-B5)/B5,0)")
        ws.cell(row=sr + 3, column=2).number_format = '0.0%'
        ws.cell(row=sr + 3, column=2).font = Font(bold=True)

        # Column widths
        ws.column_dimensions['A'].width = 35
        ws.column_dimensions['B'].width = 15
        ws.column_dimensions['C'].width = 15
        ws.column_dimensions['D'].width = 15

        self.cell_map['quality_of_earnings'] = {
            'reported_ebitda_row': 5,
            'adj_start_row': adj_start_row,
            'adj_end_row': adj_end_row,
            'total_adj_row': total_row,
            'adjusted_ebitda_row': adj_ebitda_row,
            'run_rate_row': rr_total_row,
        }

        return self

    # ============================================================
    # MODULE: WORKING CAPITAL NORMALIZATION
    # ============================================================

    def add_working_capital_normalization(self, data: Dict = None) -> 'DDModulesMixin':
        """
        Add Working Capital Normalization analysis sheet.
        Includes target vs actual NWC, peg mechanism analysis.

        Smart cross-ref: When Operating Model exists in cell_map, a
        "Projected NWC" section is added below the historical analysis
        that pulls revenue from the Operating Model for each projection
        year and calculates projected NWC = Projected Revenue x Average
        NWC %. Historical section remains input-based (always manual).
        """
        ws = self.wb.create_sheet("NWC Normalization")
        self.sheets_created.append("NWC Normalization")

        a = self.assumptions
        data = data or {}

        # Check if Operating Model has been built
        om = self.cell_map.get('operating_model', {})

        # Title
        self._add_title(ws, f"{self.company_name} - Working Capital Normalization", 1, 1)

        # Historical NWC Section
        self._add_section_header(ws, "HISTORICAL WORKING CAPITAL ($M)", 3, 1)

        # Headers
        headers = ["", "LTM", "FY-1", "FY-2", "FY-3", "Average"]
        for i, h in enumerate(headers):
            ws.cell(row=4, column=1 + i, value=h)
        self._format_header_row(ws, 4, 1, len(headers))

        # Historical data
        historical = data.get('historical_nwc', {
            'accounts_receivable': [25.0, 23.0, 22.0, 20.0],
            'inventory': [15.0, 14.0, 13.0, 12.0],
            'prepaid_expenses': [3.0, 2.8, 2.5, 2.3],
            'accounts_payable': [18.0, 16.0, 15.0, 14.0],
            'accrued_expenses': [8.0, 7.5, 7.0, 6.5],
            'deferred_revenue': [5.0, 4.5, 4.0, 3.5],
            'revenue': [200.0, 185.0, 170.0, 155.0],
        })

        row = 5
        # Current Assets
        ws.cell(row=row, column=1, value="Current Assets")
        ws.cell(row=row, column=1).font = Font(bold=True, italic=True)
        row += 1

        ca_start = row
        for item_name, values in [
            ("  Accounts Receivable", historical['accounts_receivable']),
            ("  Inventory", historical['inventory']),
            ("  Prepaid Expenses", historical['prepaid_expenses']),
        ]:
            ws.cell(row=row, column=1, value=item_name)
            for i, v in enumerate(values):
                ws.cell(row=row, column=2 + i, value=v)
                ws.cell(row=row, column=2 + i).number_format = '#,##0.0'
                self._format_input_cell(ws, row, 2 + i)
            ws.cell(row=row, column=6, value=f"=AVERAGE(B{row}:E{row})")
            ws.cell(row=row, column=6).number_format = '#,##0.0'
            row += 1
        ca_end = row - 1

        # Total Current Assets
        ca_total_row = row
        ws.cell(row=row, column=1, value="Total Current Assets")
        for i in range(4):
            c = get_column_letter(2 + i)
            ws.cell(row=row, column=2 + i, value=f"=SUM({c}{ca_start}:{c}{ca_end})")
            ws.cell(row=row, column=2 + i).number_format = '#,##0.0'
        ws.cell(row=row, column=6, value=f"=AVERAGE(B{row}:E{row})")
        ws.cell(row=row, column=6).number_format = '#,##0.0'
        ws.cell(row=row, column=1).font = Font(bold=True)
        row += 2

        # Current Liabilities
        ws.cell(row=row, column=1, value="Current Liabilities")
        ws.cell(row=row, column=1).font = Font(bold=True, italic=True)
        row += 1

        cl_item_start = row
        for item_name, values in [
            ("  Accounts Payable", historical['accounts_payable']),
            ("  Accrued Expenses", historical['accrued_expenses']),
            ("  Deferred Revenue", historical['deferred_revenue']),
        ]:
            ws.cell(row=row, column=1, value=item_name)
            for i, v in enumerate(values):
                ws.cell(row=row, column=2 + i, value=v)
                ws.cell(row=row, column=2 + i).number_format = '#,##0.0'
                self._format_input_cell(ws, row, 2 + i)
            ws.cell(row=row, column=6, value=f"=AVERAGE(B{row}:E{row})")
            ws.cell(row=row, column=6).number_format = '#,##0.0'
            row += 1
        cl_item_end = row - 1

        # Total Current Liabilities
        cl_total_row = row
        ws.cell(row=row, column=1, value="Total Current Liabilities")
        for i in range(4):
            c = get_column_letter(2 + i)
            ws.cell(row=row, column=2 + i, value=f"=SUM({c}{cl_item_start}:{c}{cl_item_end})")
            ws.cell(row=row, column=2 + i).number_format = '#,##0.0'
        ws.cell(row=row, column=6, value=f"=AVERAGE(B{row}:E{row})")
        ws.cell(row=row, column=6).number_format = '#,##0.0'
        ws.cell(row=row, column=1).font = Font(bold=True)
        row += 2

        # Net Working Capital
        nwc_row = row
        ws.cell(row=row, column=1, value="Net Working Capital")
        for i in range(4):
            c = get_column_letter(2 + i)
            ws.cell(row=row, column=2 + i, value=f"={c}{ca_total_row}-{c}{cl_total_row}")
            ws.cell(row=row, column=2 + i).number_format = '#,##0.0'
            ws.cell(row=row, column=2 + i).font = Font(bold=True)
        ws.cell(row=row, column=6, value=f"=AVERAGE(B{row}:E{row})")
        ws.cell(row=row, column=6).number_format = '#,##0.0'
        ws.cell(row=row, column=6).font = Font(bold=True)
        ws.cell(row=row, column=6).fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
        row += 2

        # Revenue reference row (needed for NWC % and Days calculations)
        revenues = historical['revenue']
        rev_row = row
        ws.cell(row=row, column=1, value="Revenue")
        for i, v in enumerate(revenues):
            ws.cell(row=row, column=2 + i, value=v)
            ws.cell(row=row, column=2 + i).number_format = '#,##0.0'
            self._format_input_cell(ws, row, 2 + i)
        ws.cell(row=row, column=6, value=f"=AVERAGE(B{row}:E{row})")
        ws.cell(row=row, column=6).number_format = '#,##0.0'
        row += 1

        # NWC as % of Revenue
        nwc_pct_row = row
        ws.cell(row=row, column=1, value="NWC as % of Revenue")
        for i in range(4):
            c = get_column_letter(2 + i)
            ws.cell(row=row, column=2 + i, value=f"=IFERROR({c}{nwc_row}/{c}{rev_row},0)")
            ws.cell(row=row, column=2 + i).number_format = '0.0%'
        ws.cell(row=row, column=6, value=f"=AVERAGE(B{row}:E{row})")
        ws.cell(row=row, column=6).number_format = '0.0%'
        ws.cell(row=row, column=6).font = Font(bold=True)
        row += 3

        # ── PROJECTED NWC SECTION (cross-ref Operating Model) ──
        # Only added when Operating Model exists in cell_map
        proj_nwc_start_row = None  # Track for cell_map
        if om and 'revenue_row' in om and 'y1_col' in om:
            n = self.projection_years
            om_revenue_row = om['revenue_row']
            om_y1_col = om['y1_col']

            self._add_section_header(ws, "PROJECTED NWC ($M)", row, 1)
            row += 1

            # Projection year headers
            proj_header_row = row
            ws.cell(row=row, column=1, value="")
            for i in range(n):
                ws.cell(row=row, column=2 + i, value=f"Year {i + 1}")
            self._format_header_row(ws, row, 1, 1 + n)
            row += 1

            # Projected Revenue (cross-ref from Operating Model)
            proj_rev_row = row
            ws.cell(row=row, column=1, value="Projected Revenue")
            for i in range(n):
                om_col_letter = get_column_letter(om_y1_col + i)
                ws.cell(row=row, column=2 + i,
                        value=f"='Operating Model'!{om_col_letter}{om_revenue_row}")
                ws.cell(row=row, column=2 + i).number_format = '#,##0.0'
            row += 1

            # Average NWC % (reference from the historical average computed above)
            avg_nwc_pct_row = row
            ws.cell(row=row, column=1, value="Average NWC % of Revenue")
            for i in range(n):
                ws.cell(row=row, column=2 + i, value=f"=$F${nwc_pct_row}")
                ws.cell(row=row, column=2 + i).number_format = '0.0%'
                ws.cell(row=row, column=2 + i).font = Font(italic=True, color="666666")
            row += 1

            # Projected NWC = Projected Revenue x Average NWC %
            proj_nwc_start_row = row
            ws.cell(row=row, column=1, value="Projected NWC")
            for i in range(n):
                c = get_column_letter(2 + i)
                ws.cell(row=row, column=2 + i,
                        value=f"={c}{proj_rev_row}*{c}{avg_nwc_pct_row}")
                ws.cell(row=row, column=2 + i).number_format = '#,##0.0'
                ws.cell(row=row, column=2 + i).font = Font(bold=True)
                ws.cell(row=row, column=2 + i).fill = PatternFill(
                    start_color="90EE90", end_color="90EE90", fill_type="solid")
            row += 1

            # Change in NWC (Y/Y) for cash flow impact
            ws.cell(row=row, column=1, value="Change in NWC (Y/Y)")
            for i in range(n):
                c = get_column_letter(2 + i)
                if i == 0:
                    # First projection year vs LTM NWC
                    ws.cell(row=row, column=2 + i,
                            value=f"={c}{proj_nwc_start_row}-B{nwc_row}")
                else:
                    prev_c = get_column_letter(2 + i - 1)
                    ws.cell(row=row, column=2 + i,
                            value=f"={c}{proj_nwc_start_row}-{prev_c}{proj_nwc_start_row}")
                ws.cell(row=row, column=2 + i).number_format = '#,##0.0'
                ws.cell(row=row, column=2 + i).font = Font(color="FF0000")
            row += 3

        # Target NWC / Peg Analysis
        self._add_section_header(ws, "NWC PEG MECHANISM", row, 1)
        row += 1

        target_row_num = row
        ws.cell(row=row, column=1, value="Target NWC (Peg)")
        if 'target_nwc' in data:
            ws.cell(row=row, column=2, value=data['target_nwc'])
        else:
            ws.cell(row=row, column=2, value=f"=F{nwc_row}")  # Default = average NWC
        ws.cell(row=row, column=2).number_format = '#,##0.0'
        self._format_input_cell(ws, row, 2)
        row += 1

        closing_row_num = row
        ws.cell(row=row, column=1, value="Estimated Closing NWC")
        if 'actual_nwc' in data:
            ws.cell(row=row, column=2, value=data['actual_nwc'])
        else:
            ws.cell(row=row, column=2, value=f"=B{nwc_row}")  # Default = LTM NWC
        ws.cell(row=row, column=2).number_format = '#,##0.0'
        self._format_input_cell(ws, row, 2)
        row += 2

        ws.cell(row=row, column=1, value="NWC Variance")
        ws.cell(row=row, column=2, value=f"=B{closing_row_num}-B{target_row_num}")
        ws.cell(row=row, column=2).number_format = '#,##0.0'
        ws.cell(row=row, column=2).font = Font(bold=True)
        ws.cell(row=row, column=3, value=f'=IF(B{row}<0,"← Buyer receives adjustment","← Seller receives adjustment")')
        ws.cell(row=row, column=3).font = Font(italic=True)
        row += 2

        # Days Analysis
        self._add_section_header(ws, "DAYS ANALYSIS", row, 1)
        row += 1

        ws.cell(row=row, column=1, value="")
        ws.cell(row=row, column=2, value="Current")
        ws.cell(row=row, column=3, value="Target")
        ws.cell(row=row, column=4, value="Benchmark")
        self._format_header_row(ws, row, 1, 4)
        row += 1

        # AR = ca_start, Inventory = ca_start+1, AP = cl_item_start
        ar_row = ca_start
        inv_row_num = ca_start + 1
        ap_row = cl_item_start

        dso_row = row
        ws.cell(row=row, column=1, value="Days Sales Outstanding (DSO)")
        ws.cell(row=row, column=2, value=f"=IFERROR(B{ar_row}/B{rev_row}*365,0)")
        ws.cell(row=row, column=3, value=data.get('target_dso', 40))
        ws.cell(row=row, column=4, value=data.get('benchmark_dso', 45))
        for col in range(2, 5):
            ws.cell(row=row, column=col).number_format = '0'
        self._format_input_cell(ws, row, 3)
        self._format_input_cell(ws, row, 4)
        row += 1

        dio_row = row
        ws.cell(row=row, column=1, value="Days Inventory Outstanding (DIO)")
        ws.cell(row=row, column=2, value=f"=IFERROR(B{inv_row_num}/(B{rev_row}*0.6)*365,0)")
        ws.cell(row=row, column=3, value=data.get('target_dio', 35))
        ws.cell(row=row, column=4, value=data.get('benchmark_dio', 40))
        for col in range(2, 5):
            ws.cell(row=row, column=col).number_format = '0'
        self._format_input_cell(ws, row, 3)
        self._format_input_cell(ws, row, 4)
        row += 1

        dpo_row = row
        ws.cell(row=row, column=1, value="Days Payable Outstanding (DPO)")
        ws.cell(row=row, column=2, value=f"=IFERROR(B{ap_row}/(B{rev_row}*0.6)*365,0)")
        ws.cell(row=row, column=3, value=data.get('target_dpo', 45))
        ws.cell(row=row, column=4, value=data.get('benchmark_dpo', 40))
        for col in range(2, 5):
            ws.cell(row=row, column=col).number_format = '0'
        self._format_input_cell(ws, row, 3)
        self._format_input_cell(ws, row, 4)
        row += 1

        # Cash Conversion Cycle
        row += 1
        ws.cell(row=row, column=1, value="Cash Conversion Cycle")
        ws.cell(row=row, column=2, value=f"=B{dso_row}+B{dio_row}-B{dpo_row}")
        ws.cell(row=row, column=2).number_format = '0'
        ws.cell(row=row, column=2).font = Font(bold=True)

        # Column widths
        ws.column_dimensions['A'].width = 30
        for col in ['B', 'C', 'D', 'E', 'F']:
            ws.column_dimensions[col].width = 12

        cell_map_entry = {
            'ca_start': ca_start,
            'ca_end': ca_end,
            'ca_total_row': ca_total_row,
            'cl_start': cl_item_start,
            'cl_end': cl_item_end,
            'cl_total_row': cl_total_row,
            'nwc_row': nwc_row,
            'rev_row': rev_row,
            'nwc_pct_row': nwc_pct_row,
            'target_row': target_row_num,
            'closing_row': closing_row_num,
            'dso_row': dso_row,
            'dio_row': dio_row,
            'dpo_row': dpo_row,
        }
        # Include projected NWC row reference if section was built
        if proj_nwc_start_row is not None:
            cell_map_entry['proj_nwc_row'] = proj_nwc_start_row
        self.cell_map['nwc_normalization'] = cell_map_entry

        return self

    # ============================================================
    # MODULE: CUSTOMER/REVENUE QUALITY
    # ============================================================

    def add_customer_revenue_quality(self, data: Dict = None) -> 'DDModulesMixin':
        """
        Add Customer/Revenue Quality analysis sheet.
        Includes cohort analysis, churn, LTV/CAC, customer concentration.

        Smart cross-ref: When Operating Model exists in cell_map, a
        cross-check row is added below the Total Revenue row that
        references Operating Model LTM revenue for reconciliation.
        """
        ws = self.wb.create_sheet("Revenue Quality")
        self.sheets_created.append("Revenue Quality")

        a = self.assumptions
        data = data or {}

        # Check if Operating Model has been built
        om = self.cell_map.get('operating_model', {})

        # Title
        self._add_title(ws, f"{self.company_name} - Customer & Revenue Quality", 1, 1)

        # Customer Concentration
        self._add_section_header(ws, "CUSTOMER CONCENTRATION", 3, 1)

        headers = ["Customer", "Revenue ($M)", "% of Total", "Tenure (Yrs)", "Contract Type"]
        for i, h in enumerate(headers):
            ws.cell(row=4, column=1 + i, value=h)
        self._format_header_row(ws, 4, 1, len(headers))

        customers = data.get('top_customers', [
            {'name': 'Customer A', 'revenue': 25.0, 'tenure': 8, 'contract': 'Multi-year'},
            {'name': 'Customer B', 'revenue': 18.0, 'tenure': 5, 'contract': 'Annual'},
            {'name': 'Customer C', 'revenue': 15.0, 'tenure': 6, 'contract': 'Multi-year'},
            {'name': 'Customer D', 'revenue': 12.0, 'tenure': 3, 'contract': 'Annual'},
            {'name': 'Customer E', 'revenue': 10.0, 'tenure': 4, 'contract': 'Month-to-month'},
            {'name': 'Other', 'revenue': 120.0, 'tenure': None, 'contract': 'Various'},
        ])

        cust_start = 5
        row = 5
        for cust in customers:
            ws.cell(row=row, column=1, value=cust['name'])
            ws.cell(row=row, column=2, value=cust['revenue'])
            ws.cell(row=row, column=2).number_format = '#,##0.0'
            self._format_input_cell(ws, row, 2)
            ws.cell(row=row, column=4, value=cust['tenure'] if cust['tenure'] else "\u2014")
            ws.cell(row=row, column=5, value=cust['contract'])
            row += 1
        cust_end = row - 1

        # Concentration summary
        row += 1
        total_rev_row = row
        ws.cell(row=row, column=1, value="Total Revenue")
        ws.cell(row=row, column=2, value=f"=SUM(B{cust_start}:B{cust_end})")
        ws.cell(row=row, column=2).number_format = '#,##0.0'
        ws.cell(row=row, column=2).font = Font(bold=True)

        # Fill % of Total (col C) referencing total row
        for r in range(cust_start, cust_end + 1):
            ws.cell(row=r, column=3, value=f"=IFERROR(B{r}/B${total_rev_row},0)")
            ws.cell(row=r, column=3).number_format = '0.0%'
        row += 1

        # Cross-check vs. Operating Model (when available)
        xcheck_row = None
        if om and 'revenue_row' in om and 'ltm_col' in om:
            om_ltm_col_letter = get_column_letter(om['ltm_col'])
            om_revenue_row = om['revenue_row']
            xcheck_row = row
            ws.cell(row=row, column=1, value="Cross-check vs. Operating Model")
            ws.cell(row=row, column=1).font = Font(italic=True, color="666666")
            ws.cell(row=row, column=2,
                    value=f"='Operating Model'!{om_ltm_col_letter}{om_revenue_row}")
            ws.cell(row=row, column=2).number_format = '#,##0.0'
            ws.cell(row=row, column=2).font = Font(italic=True, color="666666")
            # Variance column
            ws.cell(row=row, column=3, value=f"=IFERROR(B{total_rev_row}-B{row},0)")
            ws.cell(row=row, column=3).number_format = '#,##0.0'
            ws.cell(row=row, column=3).font = Font(italic=True, color="FF0000")
            ws.cell(row=row, column=4, value="Variance ($M)")
            ws.cell(row=row, column=4).font = Font(italic=True, color="666666")
            row += 1

        top5_end = min(cust_start + 4, cust_end)
        ws.cell(row=row, column=1, value="Top 5 Concentration")
        ws.cell(row=row, column=2, value=f"=IFERROR(SUM(B{cust_start}:B{top5_end})/B{total_rev_row},0)")
        ws.cell(row=row, column=2).number_format = '0.0%'
        ws.cell(row=row, column=2).font = Font(bold=True)
        ws.cell(row=row, column=3, value=f'=IF(B{row}>0.5,"\u26a0 High concentration risk","")')
        ws.cell(row=row, column=3).font = Font(color="FF0000")
        row += 1

        ws.cell(row=row, column=1, value="Top 1 Concentration")
        ws.cell(row=row, column=2, value=f"=IFERROR(B{cust_start}/B{total_rev_row},0)")
        ws.cell(row=row, column=2).number_format = '0.0%'
        row += 3

        # Revenue Retention / Churn
        self._add_section_header(ws, "REVENUE RETENTION & CHURN", row, 1)
        row += 1

        headers = ["Metric", "FY-2", "FY-1", "LTM", "Trend"]
        for i, h in enumerate(headers):
            ws.cell(row=row, column=1 + i, value=h)
        self._format_header_row(ws, row, 1, len(headers))
        row += 1

        retention_data = data.get('retention_metrics', [
            {'name': 'Gross Revenue Retention', 'values': [0.92, 0.94, 0.95], 'good': 'up'},
            {'name': 'Net Revenue Retention', 'values': [1.05, 1.08, 1.12], 'good': 'up'},
            {'name': 'Logo Churn Rate', 'values': [0.08, 0.06, 0.05], 'good': 'down'},
            {'name': 'Dollar Churn Rate', 'values': [0.08, 0.06, 0.05], 'good': 'down'},
        ])

        for metric in retention_data:
            ws.cell(row=row, column=1, value=metric['name'])
            for i, v in enumerate(metric['values']):
                ws.cell(row=row, column=2 + i, value=v)
                ws.cell(row=row, column=2 + i).number_format = '0.0%'

            # Trend indicator
            trend = metric['values'][-1] - metric['values'][0]
            if (trend > 0 and metric['good'] == 'up') or (trend < 0 and metric['good'] == 'down'):
                ws.cell(row=row, column=5, value="\u2713 Improving")
                ws.cell(row=row, column=5).font = Font(color="008000")
            else:
                ws.cell(row=row, column=5, value="\u26a0 Declining")
                ws.cell(row=row, column=5).font = Font(color="FF0000")
            row += 1

        row += 2

        # Unit Economics (for recurring revenue businesses)
        self._add_section_header(ws, "UNIT ECONOMICS", row, 1)
        row += 1

        unit_econ = data.get('unit_economics', {
            'arpu': 15000,
            'cac': 8000,
            'gross_margin': 0.75,
            'churn_rate': 0.05,
        })

        # Input rows (fixed positions for formulas)
        arpu_row = row
        ws.cell(row=row, column=1, value="Average Revenue Per User (ARPU)")
        ws.cell(row=row, column=2, value=unit_econ['arpu'])
        ws.cell(row=row, column=2).number_format = '#,##0'
        self._format_input_cell(ws, row, 2)
        row += 1

        cac_row = row
        ws.cell(row=row, column=1, value="Customer Acquisition Cost (CAC)")
        ws.cell(row=row, column=2, value=unit_econ['cac'])
        ws.cell(row=row, column=2).number_format = '#,##0'
        self._format_input_cell(ws, row, 2)
        row += 1

        margin_row = row
        ws.cell(row=row, column=1, value="Gross Margin")
        ws.cell(row=row, column=2, value=unit_econ['gross_margin'])
        ws.cell(row=row, column=2).number_format = '0.0%'
        self._format_input_cell(ws, row, 2)
        row += 1

        churn_row = row
        ws.cell(row=row, column=1, value="Annual Churn Rate")
        ws.cell(row=row, column=2, value=unit_econ['churn_rate'])
        ws.cell(row=row, column=2).number_format = '0.0%'
        self._format_input_cell(ws, row, 2)
        row += 1

        ltv_row = row
        ws.cell(row=row, column=1, value="Customer Lifetime Value (LTV)")
        ws.cell(row=row, column=2, value=f"=IFERROR(B{arpu_row}*B{margin_row}/B{churn_row},0)")
        ws.cell(row=row, column=2).number_format = '#,##0'
        row += 1

        ltv_cac_row = row
        ws.cell(row=row, column=1, value="LTV/CAC Ratio")
        ws.cell(row=row, column=2, value=f"=IFERROR(B{ltv_row}/B{cac_row},0)")
        ws.cell(row=row, column=2).number_format = '0.0x'
        ws.cell(row=row, column=2).font = Font(bold=True)
        ws.cell(row=row, column=3, value=f'=IF(B{row}>=3,"\u2713 Healthy (>3x)",IF(B{row}>=1,"\u26a0 Marginal (1-3x)","\u2717 Unhealthy (<1x)"))')
        row += 1

        ws.cell(row=row, column=1, value="CAC Payback (months)")
        ws.cell(row=row, column=2, value=f"=IFERROR(B{cac_row}/(B{arpu_row}*B{margin_row}/12),0)")
        ws.cell(row=row, column=2).number_format = '0.0'
        row += 1

        row += 2

        # Revenue Quality Score
        self._add_section_header(ws, "REVENUE QUALITY ASSESSMENT", row, 1)
        row += 1

        quality_factors = data.get('quality_factors', [
            {'factor': 'Recurring vs. One-time', 'score': 4, 'notes': '85% recurring revenue'},
            {'factor': 'Contract Length', 'score': 3, 'notes': 'Mix of annual and multi-year'},
            {'factor': 'Customer Concentration', 'score': 3, 'notes': 'Top 5 = 40% of revenue'},
            {'factor': 'Retention Rate', 'score': 4, 'notes': 'NRR > 110%'},
            {'factor': 'Pricing Power', 'score': 3, 'notes': '3-5% annual price increases'},
        ])

        headers = ["Quality Factor", "Score (1-5)", "Notes"]
        for i, h in enumerate(headers):
            ws.cell(row=row, column=1 + i, value=h)
        self._format_header_row(ws, row, 1, 3)
        row += 1

        qf_start = row
        for factor in quality_factors:
            ws.cell(row=row, column=1, value=factor['factor'])
            ws.cell(row=row, column=2, value=factor['score'])
            self._format_input_cell(ws, row, 2)
            ws.cell(row=row, column=3, value=factor['notes'])
            row += 1
        qf_end = row - 1

        row += 1
        ws.cell(row=row, column=1, value="Overall Quality Score")
        ws.cell(row=row, column=2, value=f"=AVERAGE(B{qf_start}:B{qf_end})")
        ws.cell(row=row, column=2).number_format = '0.0'
        ws.cell(row=row, column=2).font = Font(bold=True)
        ws.cell(row=row, column=2).fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")

        # Column widths
        ws.column_dimensions['A'].width = 30
        ws.column_dimensions['B'].width = 15
        ws.column_dimensions['C'].width = 15
        ws.column_dimensions['D'].width = 15
        ws.column_dimensions['E'].width = 20

        cell_map_entry = {
            'total_rev_row': total_rev_row,
            'cust_start': cust_start,
            'cust_end': cust_end,
        }
        if xcheck_row is not None:
            cell_map_entry['xcheck_row'] = xcheck_row
        self.cell_map['revenue_quality'] = cell_map_entry

        return self

    # ============================================================
    # MODULE: CREDIT/DEBT SIZING
    # ============================================================

    def add_credit_debt_sizing(self, data: Dict = None) -> 'DDModulesMixin':
        """
        Add Credit/Debt Sizing analysis with live Excel formulas.

        Smart cross-ref upgrades:
        - Row 6 (Total Debt): When Sources & Uses exists, uses formula
          ='Sources & Uses'!B4+'Sources & Uses'!B5 (senior + sub debt).
          Falls back to hardcoded calculation when S&U not built.
        - Row 9 (Interest Expense): When Debt Schedule (standard mode)
          exists, references DS total interest row at the entry column.
          Falls back to hardcoded approximation when DS not built.
        """
        ws = self.wb.create_sheet("Credit Analysis")
        self.sheets_created.append("Credit Analysis")

        a = self.assumptions
        data = data or {}

        # Check upstream modules
        su = self.cell_map.get('sources_uses', {})
        ds = self.cell_map.get('debt_schedule', {})

        # Title
        self._add_title(ws, f"{self.company_name} - Credit & Debt Sizing Analysis", 1, 1)

        # ── CREDIT PROFILE (rows 3-10) ──
        self._add_section_header(ws, "CREDIT PROFILE", 3, 1)

        # Row 4: LTM EBITDA (always cross-ref Assumptions)
        ws.cell(row=4, column=1, value="LTM EBITDA")
        ws.cell(row=4, column=2, value=f"='Assumptions'!B6")
        ws.cell(row=4, column=2).number_format = '#,##0.0'

        # Row 5: LTM Revenue (always cross-ref Assumptions)
        ws.cell(row=5, column=1, value="LTM Revenue")
        ws.cell(row=5, column=2, value=f"='Assumptions'!B5")
        ws.cell(row=5, column=2).number_format = '#,##0.0'

        # Row 6: Total Debt -- cross-ref Sources & Uses if available
        ws.cell(row=6, column=1, value="Total Debt")
        if su and 'senior_debt' in su and 'sub_debt' in su:
            # Smart formula: Senior Debt + Sub Debt from Sources & Uses
            ws.cell(row=6, column=2,
                    value=f"='Sources & Uses'!{su['senior_debt']}+'Sources & Uses'!{su['sub_debt']}")
        else:
            # Fallback: hardcoded from assumptions
            ws.cell(row=6, column=2, value=a.ltm_ebitda * (a.senior_debt_multiple + a.sub_debt_multiple))
        ws.cell(row=6, column=2).number_format = '#,##0.0'

        # Row 7: Cash & Equivalents (input)
        ws.cell(row=7, column=1, value="Cash & Equivalents")
        ws.cell(row=7, column=2, value=data.get('cash', 10.0))
        ws.cell(row=7, column=2).number_format = '#,##0.0'
        self._format_input_cell(ws, 7, 2)

        # Row 8: Net Debt (formula)
        ws.cell(row=8, column=1, value="Net Debt")
        ws.cell(row=8, column=2, value="=B6-B7")
        ws.cell(row=8, column=2).number_format = '#,##0.0'

        # Row 9: Interest Expense -- cross-ref Debt Schedule if available
        ws.cell(row=9, column=1, value="Interest Expense")
        if ds and ds.get('mode') == 'standard' and 'total_interest_row' in ds:
            # In standard mode, DS uses entry_col=2 (B) for entry values.
            # Total Interest row has per-year values starting at y1_col (C).
            # For an LTM / entry-level approximation, use Year 1 (col C)
            # which is the first full-year interest figure.
            ds_ti_row = ds['total_interest_row']
            ws.cell(row=9, column=2,
                    value=f"='Debt Schedule'!C{ds_ti_row}")
        elif ds and ds.get('mode') == 'quick' and 'interest_row' in ds:
            # Quick mode: single tranche, Year 1 interest at col C
            ds_int_row = ds['interest_row']
            ws.cell(row=9, column=2,
                    value=f"='Debt Schedule'!C{ds_int_row}")
        else:
            # Fallback: hardcoded approximation
            ws.cell(row=9, column=2, value=a.ltm_ebitda * a.senior_debt_multiple * 0.08)
        ws.cell(row=9, column=2).number_format = '#,##0.0'
        if not ds:
            # Only style as input when using hardcoded value
            self._format_input_cell(ws, 9, 2)

        # Row 10: CapEx (cross-ref Assumptions)
        ws.cell(row=10, column=1, value="CapEx")
        ws.cell(row=10, column=2, value=f"=B5*'Assumptions'!B22")
        ws.cell(row=10, column=2).number_format = '#,##0.0'

        # ── KEY CREDIT RATIOS (rows 12-18) ──
        self._add_section_header(ws, "KEY CREDIT RATIOS", 12, 1)

        ws.cell(row=13, column=1, value="Ratio")
        ws.cell(row=13, column=2, value="Current")
        ws.cell(row=13, column=3, value="Threshold")
        ws.cell(row=13, column=4, value="Status")
        self._format_header_row(ws, 13, 1, 4)

        # All ratios as formulas
        ratio_rows = [
            (14, "Total Debt / EBITDA", "=IFERROR(B6/B4,0)", 5.5, "max"),
            (15, "Net Debt / EBITDA", "=IFERROR(B8/B4,0)", 5.0, "max"),
            (16, "EBITDA / Interest", "=IFERROR(B4/B9,0)", 2.0, "min"),
            (17, "(EBITDA - CapEx) / Interest", "=IFERROR((B4-B10)/B9,0)", 1.5, "min"),
            (18, "Debt / Revenue", "=IFERROR(B6/B5,0)", 0.5, "max"),
        ]

        for row, name, formula, threshold, direction in ratio_rows:
            ws.cell(row=row, column=1, value=name)
            ws.cell(row=row, column=2, value=formula)
            ws.cell(row=row, column=2).number_format = '0.00x'
            ws.cell(row=row, column=3, value=threshold)
            ws.cell(row=row, column=3).number_format = '0.00x'
            # Status with IF formula
            if direction == "max":
                ws.cell(row=row, column=4, value=f'=IF(B{row}<=C{row},"Pass","Fail")')
            else:
                ws.cell(row=row, column=4, value=f'=IF(B{row}>=C{row},"Pass","Fail")')

        # ── DEBT CAPACITY (rows 20-28) ──
        self._add_section_header(ws, "DEBT CAPACITY ANALYSIS", 20, 1)

        ws.cell(row=21, column=1, value="Constraint")
        ws.cell(row=21, column=2, value="Multiple")
        ws.cell(row=21, column=3, value="Max Debt ($M)")
        self._format_header_row(ws, 21, 1, 3)

        constraints = [
            (22, 'Leverage Ratio (5.5x EBITDA)', 5.5),
            (23, 'Interest Coverage (2.0x floor)', 4.5),
            (24, 'Fixed Charge Coverage (1.5x)', 4.0),
            (25, 'Senior Secured (4.0x)', 4.0),
        ]

        for row, name, mult in constraints:
            ws.cell(row=row, column=1, value=name)
            ws.cell(row=row, column=2, value=mult)
            ws.cell(row=row, column=2).number_format = '0.0x'
            self._format_input_cell(ws, row, 2)
            ws.cell(row=row, column=3, value=f"=B4*B{row}")
            ws.cell(row=row, column=3).number_format = '#,##0.0'

        ws.cell(row=27, column=1, value="Binding Constraint (Max Debt)")
        ws.cell(row=27, column=2, value="=MIN(C22:C25)")
        ws.cell(row=27, column=2).number_format = '#,##0.0'
        ws.cell(row=27, column=2).font = Font(bold=True)
        ws.cell(row=27, column=2).fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
        ws.cell(row=27, column=3, value="=IFERROR(B27/B4,0)")
        ws.cell(row=27, column=3).number_format = '0.0x'
        ws.cell(row=27, column=3).font = Font(bold=True)

        # ── STRESS TEST (rows 29-36) ──
        self._add_section_header(ws, "EBITDA STRESS TEST", 29, 1)

        ws.cell(row=30, column=1, value="EBITDA Decline")
        ws.cell(row=30, column=2, value="Stressed EBITDA")
        ws.cell(row=30, column=3, value="Leverage")
        ws.cell(row=30, column=4, value="Coverage")
        ws.cell(row=30, column=5, value="Status")
        self._format_header_row(ws, 30, 1, 5)

        stress_levels = [0.0, -0.10, -0.20, -0.30, -0.40]
        for i, decline in enumerate(stress_levels):
            row = 31 + i
            ws.cell(row=row, column=1, value=decline)
            ws.cell(row=row, column=1).number_format = '0%'
            # Stressed EBITDA = LTM x (1 + decline)
            ws.cell(row=row, column=2, value=f"=$B$4*(1+A{row})")
            ws.cell(row=row, column=2).number_format = '#,##0.0'
            # Leverage = Debt / Stressed EBITDA
            ws.cell(row=row, column=3, value=f"=IFERROR($B$6/B{row},999)")
            ws.cell(row=row, column=3).number_format = '0.0x'
            # Coverage = Stressed EBITDA / Interest
            ws.cell(row=row, column=4, value=f"=IFERROR(B{row}/$B$9,0)")
            ws.cell(row=row, column=4).number_format = '0.0x'
            # Status formula
            ws.cell(row=row, column=5,
                    value=f'=IF(AND(C{row}<=6,D{row}>=1.5),"Serviceable",'
                          f'IF(AND(C{row}<=7,D{row}>=1),"Tight","Distressed"))')

        # ── IMPLIED RATING (row 38) ──
        self._add_section_header(ws, "IMPLIED CREDIT RATING", 38, 1)
        ws.cell(row=39, column=1, value="Implied Rating")
        ws.cell(row=39, column=2,
                value='=IF(AND(B14<2,B16>8),"BBB / Baa2",'
                      'IF(AND(B14<3.5,B16>4),"BB / Ba2",'
                      'IF(AND(B14<5,B16>2.5),"B+ / B1",'
                      'IF(AND(B14<6,B16>2),"B / B2","B- / B3 or lower"))))')
        ws.cell(row=39, column=2).font = Font(bold=True)

        # Column widths
        ws.column_dimensions['A'].width = 30
        for col in ['B', 'C', 'D', 'E']:
            ws.column_dimensions[col].width = 15

        self.cell_map['credit_analysis'] = {
            'ebitda': 'B4',
            'revenue': 'B5',
            'total_debt': 'B6',
            'cash': 'B7',
            'net_debt': 'B8',
            'interest': 'B9',
            'capex': 'B10',
            'leverage_ratio': 'B14',
            'coverage_ratio': 'B16',
            'max_debt': 'B27',
            'implied_rating': 'B39',
        }

        return self
