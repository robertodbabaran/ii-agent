#!/usr/bin/env python3
"""
Specialized Excel Modules Mixin.

Contains: Rollup Model, Tax Analysis,
Control Premium Analysis, Source Index.
"""

from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from datetime import datetime
from .formula_builder import FormulaBuilder as FB
from ._base import (INPUT_FILL, INPUT_FONT, HEADER_FILL, HEADER_FONT,
                     BOTTOM_BORDER, DOUBLE_BORDER, ModelDepth)
from typing import Dict


class SpecializedModulesMixin:
    """Mixin providing specialized situation analysis modules."""

    def add_rollup_model(self, data: Dict = None) -> 'SpecializedModulesMixin':
        """Add Roll-up/Consolidation Model sheet with EXIT ANALYSIS section."""
        ws = self.wb.create_sheet("Roll-up Model")
        self.sheets_created.append("Roll-up Model")
        a = self.assumptions
        data = data or {}

        self._add_title(ws, f"{self.company_name} - Roll-up Strategy", 1, 1)

        ws.cell(row=3, column=1, value="Platform EBITDA")
        ws.cell(row=3, column=3, value=f"='Assumptions'!B6")
        ws.cell(row=3, column=3).number_format = '#,##0.0'

        self._add_section_header(ws, "ACQUISITION SCHEDULE", 4, 1)
        headers = ["Year", "Target", "EBITDA", "Multiple", "EV"]
        for i, h in enumerate(headers):
            ws.cell(row=5, column=1+i, value=h)
        self._format_header_row(ws, 5, 1, 5)

        acquisitions = data.get('acquisitions', [
            {'year': 1, 'name': 'Tuck-in A', 'ebitda': 3.0, 'multiple': 5.0},
            {'year': 2, 'name': 'Regional B', 'ebitda': 5.0, 'multiple': 5.5},
            {'year': 3, 'name': 'Strategic C', 'ebitda': 8.0, 'multiple': 6.0},
        ])
        acq_start = 6
        for i, acq in enumerate(acquisitions):
            row = acq_start + i
            ws.cell(row=row, column=1, value=acq['year'])
            ws.cell(row=row, column=2, value=acq['name'])
            ws.cell(row=row, column=3, value=acq['ebitda'])
            ws.cell(row=row, column=3).number_format = '#,##0.0'
            self._format_input_cell(ws, row, 3)
            ws.cell(row=row, column=4, value=acq['multiple'])
            ws.cell(row=row, column=4).number_format = '0.0x'
            self._format_input_cell(ws, row, 4)
            ws.cell(row=row, column=5, value=f"=C{row}*D{row}")
            ws.cell(row=row, column=5).number_format = '#,##0.0'
        acq_end = acq_start + len(acquisitions) - 1

        pf_row = acq_end + 2
        ws.cell(row=pf_row, column=1, value="Pro Forma EBITDA")
        ws.cell(row=pf_row, column=3, value=f"=C3+SUM(C{acq_start}:C{acq_end})")
        ws.cell(row=pf_row, column=3).number_format = '#,##0.0'
        ws.cell(row=pf_row, column=3).font = Font(bold=True)
        ws.cell(row=pf_row, column=3).fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")

        blend_row = pf_row + 1
        ws.cell(row=blend_row, column=1, value="Blended Acquisition Multiple")
        ws.cell(row=blend_row, column=4,
                value=f"=IFERROR(SUM(E{acq_start}:E{acq_end})/SUM(C{acq_start}:C{acq_end}),0)")
        ws.cell(row=blend_row, column=4).number_format = '0.0x'
        ws.cell(row=blend_row, column=4).font = Font(bold=True)

        # --- EXIT ANALYSIS SECTION ---
        exit_section_row = blend_row + 2
        self._add_section_header(ws, "EXIT ANALYSIS", exit_section_row, 1)

        exit_mult_row = exit_section_row + 1
        ws.cell(row=exit_mult_row, column=1, value="Exit Multiple")
        ws.cell(row=exit_mult_row, column=2, value="='Assumptions'!B8")
        ws.cell(row=exit_mult_row, column=2).number_format = '0.0x'

        pf_ev_row = exit_mult_row + 1
        ws.cell(row=pf_ev_row, column=1, value="Pro Forma Exit EV")
        ws.cell(row=pf_ev_row, column=2, value=f"=C{pf_row}*B{exit_mult_row}")
        ws.cell(row=pf_ev_row, column=2).number_format = '#,##0.0'
        ws.cell(row=pf_ev_row, column=2).font = Font(bold=True)
        ws.cell(row=pf_ev_row, column=2).fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")

        total_acq_row = pf_ev_row + 1
        ws.cell(row=total_acq_row, column=1, value="Total Acquisition Capital")
        ws.cell(row=total_acq_row, column=2, value=f"=SUM(E{acq_start}:E{acq_end})")
        ws.cell(row=total_acq_row, column=2).number_format = '#,##0.0'

        value_created_row = total_acq_row + 1
        ws.cell(row=value_created_row, column=1, value="Value Created (Multiple Arbitrage)")
        # PF Exit EV minus platform entry EV (platform EBITDA * entry multiple) minus total acquisition costs
        ws.cell(row=value_created_row, column=2,
                value=f"=B{pf_ev_row}-C3*'Assumptions'!B7-B{total_acq_row}")
        ws.cell(row=value_created_row, column=2).number_format = '#,##0.0'
        ws.cell(row=value_created_row, column=2).font = Font(bold=True, color="006100")
        ws.cell(row=value_created_row, column=2).fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")

        ws.column_dimensions['A'].width = 35
        for c in ['B','C','D','E']: ws.column_dimensions[c].width = 14

        # --- cell_map entry for rollup module ---
        self.cell_map['rollup_model'] = {
            'platform_ebitda_row': 3,
            'acq_start_row': acq_start,
            'acq_end_row': acq_end,
            'pf_ebitda_row': pf_row,
            'blended_mult_row': blend_row,
            'exit_mult_row': exit_mult_row,
            'pf_ev_row': pf_ev_row,
            'total_acq_row': total_acq_row,
            'value_created_row': value_created_row,
        }

        return self

    def add_tax_analysis(self, data: Dict = None) -> 'SpecializedModulesMixin':
        """Add Tax Analysis sheet (338(h)(10), step-up, NOLs).

        Smart cross-sheet references:
        - Row 4 (Purchase Price): Uses Sources & Uses EV when available, falls back to Assumptions.
        - Row 9 (Discount Rate): Uses WACC sheet when available, falls back to hardcoded 0.10.
        """
        ws = self.wb.create_sheet("Tax Analysis")
        self.sheets_created.append("Tax Analysis")
        a = self.assumptions
        data = data or {}

        self._add_title(ws, f"{self.company_name} - Tax Structure Analysis", 1, 1)

        self._add_section_header(ws, "STEP-UP BENEFIT (338(h)(10))", 3, 1)

        # Row 4: Purchase Price — prefer Sources & Uses EV if available, else Assumptions
        ws.cell(row=4, column=1, value="Purchase Price")
        su_map = self.cell_map.get('sources_uses', {})
        if su_map.get('purchase_ev'):
            ws.cell(row=4, column=2, value=f"='Sources & Uses'!{su_map['purchase_ev']}")
        else:
            ws.cell(row=4, column=2, value="='Assumptions'!B6*'Assumptions'!B7")
        ws.cell(row=4, column=2).number_format = '#,##0.0'

        ws.cell(row=5, column=1, value="Existing Tax Basis")
        ws.cell(row=5, column=2, value=data.get('tax_basis', a.ltm_ebitda * a.entry_multiple * 0.30))
        ws.cell(row=5, column=2).number_format = '#,##0.0'
        self._format_input_cell(ws, 5, 2)

        ws.cell(row=6, column=1, value="Step-Up Amount")
        ws.cell(row=6, column=2, value="=B4-B5")
        ws.cell(row=6, column=2).number_format = '#,##0.0'
        ws.cell(row=6, column=2).font = Font(bold=True)

        ws.cell(row=8, column=1, value="Amortization Period (years)")
        ws.cell(row=8, column=2, value=15)
        self._format_input_cell(ws, 8, 2)

        # Row 9: Discount Rate — prefer WACC sheet if available, else hardcoded 0.10
        ws.cell(row=9, column=1, value="Discount Rate")
        wacc_map = self.cell_map.get('wacc', {})
        if wacc_map.get('wacc'):
            ws.cell(row=9, column=2, value=f"='WACC'!{wacc_map['wacc']}")
        else:
            ws.cell(row=9, column=2, value=0.10)
            self._format_input_cell(ws, 9, 2)
        ws.cell(row=9, column=2).number_format = '0.0%'

        ws.cell(row=10, column=1, value="Tax Rate")
        ws.cell(row=10, column=2, value="='Assumptions'!B24")
        ws.cell(row=10, column=2).number_format = '0.0%'

        ws.cell(row=12, column=1, value="Annual Amortization")
        ws.cell(row=12, column=2, value="=B6/B8")
        ws.cell(row=12, column=2).number_format = '#,##0.0'

        ws.cell(row=13, column=1, value="Annual Tax Shield")
        ws.cell(row=13, column=2, value="=B12*B10")
        ws.cell(row=13, column=2).number_format = '#,##0.0'

        # NPV using PV annuity formula: Shield x (1 - (1+r)^-n) / r
        ws.cell(row=14, column=1, value="NPV of Tax Benefit")
        ws.cell(row=14, column=2, value="=B13*(1-(1+B9)^(-B8))/B9")
        ws.cell(row=14, column=2).number_format = '#,##0.0'
        ws.cell(row=14, column=2).font = Font(bold=True)
        ws.cell(row=14, column=2).fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")

        ws.column_dimensions['A'].width = 30
        ws.column_dimensions['B'].width = 15

        # --- cell_map entry for tax_analysis module ---
        self.cell_map['tax_analysis'] = {
            'purchase_price': 'B4',
            'tax_basis': 'B5',
            'step_up': 'B6',
            'amort_period': 'B8',
            'discount_rate': 'B9',
            'tax_rate': 'B10',
            'annual_amort': 'B12',
            'annual_shield': 'B13',
            'npv_benefit': 'B14',
        }

        return self

    def add_control_premium_analysis(self, data: Dict = None) -> 'SpecializedModulesMixin':
        """Add Control Premium Analysis sheet with IMPLIED VALUATION section."""
        ws = self.wb.create_sheet("Control Premium")
        self.sheets_created.append("Control Premium")
        data = data or {}

        self._add_title(ws, f"{self.company_name} - Control Premium Analysis", 1, 1)

        self._add_section_header(ws, "PRECEDENT TRANSACTION PREMIA", 3, 1)
        headers = ["Target", "1-Day", "30-Day"]
        row = 4
        for i, h in enumerate(headers):
            ws.cell(row=row, column=1+i, value=h)
        self._format_header_row(ws, row, 1, 3)
        row += 1

        transactions = data.get('transactions', [
            {'name': 'Comp A', '1day': 0.25, '30day': 0.35},
            {'name': 'Comp B', '1day': 0.30, '30day': 0.40},
            {'name': 'Comp C', '1day': 0.22, '30day': 0.32},
            {'name': 'Comp D', '1day': 0.28, '30day': 0.38},
        ])
        txn_start = row
        for txn in transactions:
            ws.cell(row=row, column=1, value=txn['name'])
            ws.cell(row=row, column=2, value=txn['1day'])
            ws.cell(row=row, column=2).number_format = '0.0%'
            self._format_input_cell(ws, row, 2)
            ws.cell(row=row, column=3, value=txn['30day'])
            ws.cell(row=row, column=3).number_format = '0.0%'
            self._format_input_cell(ws, row, 3)
            row += 1
        txn_end = row - 1

        row += 1
        ws.cell(row=row, column=1, value="Mean Premium")
        ws.cell(row=row, column=2, value=f"=AVERAGE(B{txn_start}:B{txn_end})")
        ws.cell(row=row, column=2).number_format = '0.0%'
        ws.cell(row=row, column=2).font = Font(bold=True)
        ws.cell(row=row, column=3, value=f"=AVERAGE(C{txn_start}:C{txn_end})")
        ws.cell(row=row, column=3).number_format = '0.0%'
        ws.cell(row=row, column=3).font = Font(bold=True)
        mean_row = row
        row += 3

        unaffected = data.get('unaffected_price', 50.0)
        self._add_section_header(ws, "IMPLIED OFFER PRICE", row, 1)
        row += 1
        ws.cell(row=row, column=1, value="Unaffected Share Price")
        ws.cell(row=row, column=2, value=unaffected)
        ws.cell(row=row, column=2).number_format = '$#,##0.00'
        self._format_input_cell(ws, row, 2)
        price_row = row
        row += 1
        ws.cell(row=row, column=1, value="Selected Premium (30-day)")
        ws.cell(row=row, column=2, value=f"=C{mean_row}")
        ws.cell(row=row, column=2).number_format = '0.0%'
        prem_row = row
        row += 1
        ws.cell(row=row, column=1, value="Implied Offer Price")
        ws.cell(row=row, column=2, value=f"=B{price_row}*(1+B{prem_row})")
        ws.cell(row=row, column=2).number_format = '$#,##0.00'
        ws.cell(row=row, column=2).font = Font(bold=True)
        ws.cell(row=row, column=2).fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
        offer_price_row = row

        # --- IMPLIED VALUATION SECTION ---
        row += 2
        self._add_section_header(ws, "IMPLIED VALUATION", row, 1)
        row += 1

        shares_row = row
        ws.cell(row=shares_row, column=1, value="Shares Outstanding (M)")
        shares_outstanding = data.get('shares_outstanding', 100.0)
        ws.cell(row=shares_row, column=2, value=shares_outstanding)
        ws.cell(row=shares_row, column=2).number_format = '#,##0.0'
        self._format_input_cell(ws, shares_row, 2)
        row += 1

        eq_val_row = row
        ws.cell(row=eq_val_row, column=1, value="Implied Equity Value")
        ws.cell(row=eq_val_row, column=2, value=f"=B{offer_price_row}*B{shares_row}")
        ws.cell(row=eq_val_row, column=2).number_format = '#,##0.0'
        ws.cell(row=eq_val_row, column=2).font = Font(bold=True)
        row += 1

        net_debt_row = row
        ws.cell(row=net_debt_row, column=1, value="Net Debt")
        # Use Assumptions-derived net debt if available, else data param or default
        assumptions_map = self.cell_map.get('assumptions', {})
        net_debt_default = data.get('net_debt', None)
        if net_debt_default is not None:
            ws.cell(row=net_debt_row, column=2, value=net_debt_default)
            self._format_input_cell(ws, net_debt_row, 2)
        elif assumptions_map.get('ltm_ebitda') and assumptions_map.get('senior_debt_mult') and assumptions_map.get('sub_debt_mult'):
            # Derive net debt from Assumptions: EBITDA * (senior + sub debt multiples)
            ws.cell(row=net_debt_row, column=2,
                    value=f"='Assumptions'!{assumptions_map['ltm_ebitda']}*('Assumptions'!{assumptions_map['senior_debt_mult']}+'Assumptions'!{assumptions_map['sub_debt_mult']})")
        else:
            ws.cell(row=net_debt_row, column=2, value=0.0)
            self._format_input_cell(ws, net_debt_row, 2)
        ws.cell(row=net_debt_row, column=2).number_format = '#,##0.0'
        row += 1

        implied_ev_row = row
        ws.cell(row=implied_ev_row, column=1, value="Implied Enterprise Value")
        ws.cell(row=implied_ev_row, column=2, value=f"=B{eq_val_row}+B{net_debt_row}")
        ws.cell(row=implied_ev_row, column=2).number_format = '#,##0.0'
        ws.cell(row=implied_ev_row, column=2).font = Font(bold=True)
        ws.cell(row=implied_ev_row, column=2).fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
        row += 1

        implied_ev_ebitda_row = row
        ws.cell(row=implied_ev_ebitda_row, column=1, value="Implied EV/EBITDA")
        ws.cell(row=implied_ev_ebitda_row, column=2,
                value=f"=IFERROR(B{implied_ev_row}/'Assumptions'!B6,0)")
        ws.cell(row=implied_ev_ebitda_row, column=2).number_format = '0.0x'
        ws.cell(row=implied_ev_ebitda_row, column=2).font = Font(bold=True)
        row += 1

        vs_entry_row = row
        ws.cell(row=vs_entry_row, column=1, value="vs. Entry Multiple")
        ws.cell(row=vs_entry_row, column=2,
                value=f"=B{implied_ev_ebitda_row}-'Assumptions'!B7")
        ws.cell(row=vs_entry_row, column=2).number_format = '0.0x'
        ws.cell(row=vs_entry_row, column=2).font = Font(bold=True, color="006100")

        ws.column_dimensions['A'].width = 30
        ws.column_dimensions['B'].width = 16
        ws.column_dimensions['C'].width = 12

        # --- cell_map entry for control_premium module ---
        self.cell_map['control_premium'] = {
            'mean_premium_1day': f'B{mean_row}',
            'mean_premium_30day': f'C{mean_row}',
            'unaffected_price': f'B{price_row}',
            'offer_price': f'B{offer_price_row}',
            'shares_outstanding': f'B{shares_row}',
            'implied_equity_value': f'B{eq_val_row}',
            'net_debt': f'B{net_debt_row}',
            'implied_ev': f'B{implied_ev_row}',
            'implied_ev_ebitda': f'B{implied_ev_ebitda_row}',
            'vs_entry_multiple': f'B{vs_entry_row}',
        }

        return self

    # ============================================================
    # MODULE: SOURCE INDEX (Canonical Data Repository)
    # ============================================================

    def add_source_index(self, metrics: list = None) -> 'SpecializedModulesMixin':
        """Add Source Index tab — canonical repository of all metrics flowing into slides.

        This tab is the SINGLE SOURCE OF TRUTH for every data point used in any
        PowerPoint slide. All other tabs (including Powerpoint Outputs bridge) pull
        from or validate against this index.

        Columns:
            A - # (sequential number)
            B - Metric (name of the data point)
            C - Value (the canonical value)
            D - Source Link (URL or document reference for auditability)
            E - Category (Transaction / Operational / Financial / Valuation / Industry / Sponsor)
            F - Slide Reference (which slide(s) use this metric)
            G - Data Vintage (FIXED / AT ENTRY / ACTUAL / DERIVED / MARKET / ILLUSTRATIVE)
            H - Last Verified (date of last verification)

        Args:
            metrics: Optional list of dicts with keys matching columns above.
                     If None, creates template with empty rows for manual population.
        """
        ws = self.wb.create_sheet("Source Index")
        self.sheets_created.append("Source Index")

        # --- GS Formatting Constants ---
        GS_HEADER_FILL = PatternFill(start_color="002060", end_color="002060", fill_type="solid")
        GS_HEADER_FONT = Font(name="Arial", size=10, color="FFFFFF", bold=True)
        GS_INPUT_FONT = Font(name="Arial", size=10, color="0070C0")  # Blue for hardcoded
        GS_CALC_FONT = Font(name="Arial", size=10, color="000000")
        GS_LINK_FONT = Font(name="Arial", size=10, color="FF0000", underline="single")  # Red for external
        GS_BODY_FONT = Font(name="Arial", size=10)
        GS_BORDER = Border(
            left=Side(style='thin', color="D9D9D9"),
            right=Side(style='thin', color="D9D9D9"),
            top=Side(style='thin', color="D9D9D9"),
            bottom=Side(style='thin', color="D9D9D9")
        )
        CATEGORY_FILLS = {
            "Transaction": PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid"),
            "Operational": PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid"),
            "Financial": PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid"),
            "Valuation": PatternFill(start_color="E2D9F3", end_color="E2D9F3", fill_type="solid"),
            "Industry": PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid"),
            "Sponsor": PatternFill(start_color="D6DCE4", end_color="D6DCE4", fill_type="solid"),
            "Pro Forma": PatternFill(start_color="F2DCDB", end_color="F2DCDB", fill_type="solid"),
        }

        # --- Sheet Setup (GS standards) ---
        ws.sheet_view.showGridLines = False
        ws.sheet_properties.pageSetUpPr = None
        ws.page_setup.orientation = 'landscape'
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0

        # --- Title ---
        ws.cell(row=1, column=1, value=f"{self.company_name} — Source Index")
        ws.cell(row=1, column=1).font = Font(name="Arial", size=14, bold=True)
        ws.cell(row=2, column=1, value="Canonical repository: ALL metrics flowing into slides originate here.")
        ws.cell(row=2, column=1).font = Font(name="Arial", size=10, italic=True, color="666666")
        ws.cell(row=3, column=1, value=f"Generated: {datetime.now().strftime('%B %d, %Y')}")
        ws.cell(row=3, column=1).font = Font(name="Arial", size=9, color="999999")

        # --- Column Headers (Row 5) ---
        headers = [
            ("#", 6),
            ("Metric", 40),
            ("Value", 18),
            ("Source Link", 50),
            ("Category", 16),
            ("Slide Reference", 25),
            ("Data Vintage", 16),
            ("Last Verified", 14),
        ]
        header_row = 5
        for col_idx, (header_name, width) in enumerate(headers, 1):
            cell = ws.cell(row=header_row, column=col_idx, value=header_name)
            cell.fill = GS_HEADER_FILL
            cell.font = GS_HEADER_FONT
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = GS_BORDER
            ws.column_dimensions[get_column_letter(col_idx)].width = width

        # --- Freeze panes below header ---
        ws.freeze_panes = "A6"

        # --- Auto-filter ---
        ws.auto_filter.ref = f"A{header_row}:H{header_row}"

        # --- Populate with metrics or template rows ---
        if metrics:
            for i, metric in enumerate(metrics, 1):
                row = header_row + i
                # Column A: Number
                ws.cell(row=row, column=1, value=i)
                ws.cell(row=row, column=1).font = GS_BODY_FONT
                ws.cell(row=row, column=1).alignment = Alignment(horizontal='center')
                ws.cell(row=row, column=1).border = GS_BORDER

                # Column B: Metric
                ws.cell(row=row, column=2, value=metric.get('metric', ''))
                ws.cell(row=row, column=2).font = GS_BODY_FONT
                ws.cell(row=row, column=2).border = GS_BORDER

                # Column C: Value
                val = metric.get('value', '')
                ws.cell(row=row, column=3, value=val)
                ws.cell(row=row, column=3).font = GS_INPUT_FONT  # Blue = hardcoded input
                ws.cell(row=row, column=3).alignment = Alignment(horizontal='right')
                ws.cell(row=row, column=3).border = GS_BORDER

                # Column D: Source Link
                source = metric.get('source', '')
                ws.cell(row=row, column=4, value=source)
                if source and (source.startswith('http://') or source.startswith('https://')):
                    ws.cell(row=row, column=4).font = GS_LINK_FONT
                    ws.cell(row=row, column=4).hyperlink = source
                else:
                    ws.cell(row=row, column=4).font = GS_BODY_FONT
                ws.cell(row=row, column=4).border = GS_BORDER

                # Column E: Category
                category = metric.get('category', '')
                ws.cell(row=row, column=5, value=category)
                ws.cell(row=row, column=5).font = GS_BODY_FONT
                ws.cell(row=row, column=5).alignment = Alignment(horizontal='center')
                ws.cell(row=row, column=5).border = GS_BORDER
                if category in CATEGORY_FILLS:
                    ws.cell(row=row, column=5).fill = CATEGORY_FILLS[category]

                # Column F: Slide Reference
                ws.cell(row=row, column=6, value=metric.get('slide_ref', ''))
                ws.cell(row=row, column=6).font = GS_BODY_FONT
                ws.cell(row=row, column=6).border = GS_BORDER

                # Column G: Data Vintage
                ws.cell(row=row, column=7, value=metric.get('vintage', ''))
                ws.cell(row=row, column=7).font = GS_BODY_FONT
                ws.cell(row=row, column=7).alignment = Alignment(horizontal='center')
                ws.cell(row=row, column=7).border = GS_BORDER

                # Column H: Last Verified
                ws.cell(row=row, column=8, value=metric.get('verified', datetime.now().strftime('%Y-%m-%d')))
                ws.cell(row=row, column=8).font = GS_BODY_FONT
                ws.cell(row=row, column=8).alignment = Alignment(horizontal='center')
                ws.cell(row=row, column=8).border = GS_BORDER

            # Update auto-filter range to include data
            last_row = header_row + len(metrics)
            ws.auto_filter.ref = f"A{header_row}:H{last_row}"
        else:
            # Template mode: create 50 empty placeholder rows with formatting
            for i in range(1, 51):
                row = header_row + i
                ws.cell(row=row, column=1, value=i)
                ws.cell(row=row, column=1).font = GS_BODY_FONT
                ws.cell(row=row, column=1).alignment = Alignment(horizontal='center')
                for col_idx in range(1, 9):
                    ws.cell(row=row, column=col_idx).border = GS_BORDER
                    if col_idx > 1:
                        ws.cell(row=row, column=col_idx).font = GS_BODY_FONT
            ws.auto_filter.ref = f"A{header_row}:H{header_row + 50}"

        # --- Print area ---
        last_data_row = header_row + (len(metrics) if metrics else 50)
        ws.print_area = f"A1:H{last_data_row}"

        # Track in cell_map for cross-sheet references
        self.cell_map["source_index"] = {
            "header_row": str(header_row),
            "data_start_row": str(header_row + 1),
            "metric_col": "B",
            "value_col": "C",
            "source_col": "D",
            "category_col": "E",
        }

        return self
