#!/usr/bin/env python3
"""
Transaction Structure Excel Modules Mixin.

Contains: Add-on Analysis, Synergy Model,
Carve-out Analysis, Earnout Model, Purchase Price Allocation.
"""

from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from .formula_builder import FormulaBuilder as FB
from ._base import (INPUT_FILL, INPUT_FONT, HEADER_FILL, HEADER_FONT,
                     BOTTOM_BORDER, DOUBLE_BORDER, ModelDepth)
from typing import Dict


class TransactionModulesMixin:
    """Mixin providing transaction structure analysis modules."""

    def add_addon_analysis(self, data: Dict = None) -> 'TransactionModulesMixin':
        """Add Add-on/Bolt-on Acquisition analysis with formulas."""
        ws = self.wb.create_sheet("Add-on Analysis")
        self.sheets_created.append("Add-on Analysis")
        a = self.assumptions
        data = data or {}

        self._add_title(ws, f"{self.company_name} - Add-on Acquisition Analysis", 1, 1)

        # Platform (rows 3-5)
        self._add_section_header(ws, "PLATFORM COMPANY", 3, 1)
        ws.cell(row=4, column=1, value="Platform EBITDA")
        ws.cell(row=4, column=2, value=f"='Assumptions'!B6")
        ws.cell(row=4, column=2).number_format = '#,##0.0'
        ws.cell(row=5, column=1, value="Entry Multiple")
        ws.cell(row=5, column=2, value=f"='Assumptions'!B7")
        ws.cell(row=5, column=2).number_format = '0.0x'

        # Add-ons (rows 7+)
        self._add_section_header(ws, "ADD-ON TARGETS", 7, 1)
        headers = ["Target", "EBITDA", "Multiple", "EV", "Synergies"]
        for i, h in enumerate(headers):
            ws.cell(row=8, column=1+i, value=h)
        self._format_header_row(ws, 8, 1, 5)

        addons = data.get('addons', [
            {'name': 'Target A', 'ebitda': 5.0, 'multiple': 5.5, 'synergies': 0.8},
            {'name': 'Target B', 'ebitda': 3.0, 'multiple': 5.0, 'synergies': 0.5},
        ])
        addon_start = 9
        for i, addon in enumerate(addons):
            row = addon_start + i
            ws.cell(row=row, column=1, value=addon['name'])
            ws.cell(row=row, column=2, value=addon['ebitda'])
            ws.cell(row=row, column=2).number_format = '#,##0.0'
            self._format_input_cell(ws, row, 2)
            ws.cell(row=row, column=3, value=addon['multiple'])
            ws.cell(row=row, column=3).number_format = '0.0x'
            self._format_input_cell(ws, row, 3)
            ws.cell(row=row, column=4, value=f"=B{row}*C{row}")
            ws.cell(row=row, column=4).number_format = '#,##0.0'
            ws.cell(row=row, column=5, value=addon['synergies'])
            ws.cell(row=row, column=5).number_format = '#,##0.0'
            self._format_input_cell(ws, row, 5)
        addon_end = addon_start + len(addons) - 1

        pf_row = addon_end + 2
        ws.cell(row=pf_row, column=1, value="Pro Forma EBITDA")
        ws.cell(row=pf_row, column=2,
                value=f"=B4+SUM(B{addon_start}:B{addon_end})+SUM(E{addon_start}:E{addon_end})")
        ws.cell(row=pf_row, column=2).number_format = '#,##0.0'
        ws.cell(row=pf_row, column=2).font = Font(bold=True)
        ws.cell(row=pf_row, column=2).fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")

        ws.cell(row=pf_row + 1, column=1, value="Multiple Arbitrage")
        ws.cell(row=pf_row + 1, column=2,
                value=f"=B5-IFERROR(SUM(D{addon_start}:D{addon_end})/SUM(B{addon_start}:B{addon_end}),0)")
        ws.cell(row=pf_row + 1, column=2).number_format = '0.0x'
        ws.cell(row=pf_row + 1, column=2).font = Font(bold=True, color="008000")

        # --- Pro Forma Returns section ---
        pf_returns_header_row = pf_row + 3
        self._add_section_header(ws, "PRO FORMA RETURNS", pf_returns_header_row, 1)

        pf_ev_row = pf_returns_header_row + 1
        ws.cell(row=pf_ev_row, column=1, value="Pro Forma Exit EV")
        ws.cell(row=pf_ev_row, column=2,
                value=f"=B{pf_row}*'Assumptions'!B8")
        ws.cell(row=pf_ev_row, column=2).number_format = '#,##0.0'

        acq_cost_row = pf_ev_row + 1
        ws.cell(row=acq_cost_row, column=1, value="Less: Total Acquisition Cost")
        ws.cell(row=acq_cost_row, column=2,
                value=f"=SUM(D{addon_start}:D{addon_end})")
        ws.cell(row=acq_cost_row, column=2).number_format = '#,##0.0'

        value_created_row = acq_cost_row + 1
        ws.cell(row=value_created_row, column=1, value="Pro Forma Value Created")
        ws.cell(row=value_created_row, column=2,
                value=f"=B{pf_ev_row}-B{acq_cost_row}")
        ws.cell(row=value_created_row, column=2).number_format = '#,##0.0'
        ws.cell(row=value_created_row, column=2).font = Font(bold=True)
        ws.cell(row=value_created_row, column=2).fill = PatternFill(
            start_color="90EE90", end_color="90EE90", fill_type="solid")

        # Update cell_map with new keys
        addon_map = self.cell_map.get('addon_analysis', {})
        addon_map['pf_ebitda_row'] = pf_row
        addon_map['pf_exit_ev_row'] = pf_ev_row
        addon_map['acq_cost_row'] = acq_cost_row
        addon_map['value_created_row'] = value_created_row
        self.cell_map['addon_analysis'] = addon_map

        ws.column_dimensions['A'].width = 25
        for c in ['B','C','D','E']: ws.column_dimensions[c].width = 12
        return self

    def add_synergy_model(self, data: Dict = None) -> 'TransactionModulesMixin':
        """Add Synergy Model sheet with cost/revenue synergies."""
        ws = self.wb.create_sheet("Synergy Model")
        self.sheets_created.append("Synergy Model")
        data = data or {}

        self._add_title(ws, f"{self.company_name} - Synergy Analysis", 1, 1)

        # --- Base EBITDA row referencing Operating Model when available ---
        om = self.cell_map.get('operating_model', {})
        base_ebitda_row = 3
        ws.cell(row=base_ebitda_row, column=1, value="Base EBITDA")
        if om and 'ebitda_row' in om and 'ltm_col' in om:
            # Reference Operating Model LTM EBITDA
            ltm_col_letter = get_column_letter(om['ltm_col'])
            ws.cell(row=base_ebitda_row, column=2,
                    value=f"='Operating Model'!{ltm_col_letter}{om['ebitda_row']}")
        else:
            # Fallback to Assumptions LTM EBITDA
            ws.cell(row=base_ebitda_row, column=2,
                    value=f"='Assumptions'!B6")
        ws.cell(row=base_ebitda_row, column=2).number_format = '#,##0.0'
        ws.cell(row=base_ebitda_row, column=2).font = Font(bold=True)

        self._add_section_header(ws, "COST SYNERGIES", 5, 1)

        headers = ["Category", "Yr 1", "Yr 2", "Yr 3", "Run-Rate", "Prob."]
        row = 6
        for i, h in enumerate(headers):
            ws.cell(row=row, column=1+i, value=h)
        self._format_header_row(ws, row, 1, 6)
        row += 1

        synergies = data.get('cost_synergies', [
            {'name': 'Headcount', 'y1': 2.0, 'y2': 4.0, 'y3': 5.0, 'prob': 0.90},
            {'name': 'Procurement', 'y1': 1.0, 'y2': 2.0, 'y3': 2.5, 'prob': 0.85},
            {'name': 'Facilities', 'y1': 0.5, 'y2': 1.5, 'y3': 2.0, 'prob': 0.75},
        ])
        syn_start = row
        for syn in synergies:
            ws.cell(row=row, column=1, value=syn['name'])
            ws.cell(row=row, column=2, value=syn['y1'])
            ws.cell(row=row, column=3, value=syn['y2'])
            ws.cell(row=row, column=4, value=syn['y3'])
            ws.cell(row=row, column=5, value=f"=D{row}")  # Run-rate = Year 3
            ws.cell(row=row, column=6, value=syn['prob'])
            ws.cell(row=row, column=6).number_format = '0%'
            for c in [2,3,4]: self._format_input_cell(ws, row, c)
            for c in [2,3,4,5]: ws.cell(row=row, column=c).number_format = '#,##0.0'
            row += 1
        syn_end = row - 1

        row += 1
        total_cost_syn_row = row
        ws.cell(row=row, column=1, value="Total Cost Synergies")
        for c in [2, 3, 4, 5]:
            cl = get_column_letter(c)
            ws.cell(row=row, column=c, value=f"=SUM({cl}{syn_start}:{cl}{syn_end})")
            ws.cell(row=row, column=c).number_format = '#,##0.0'
        ws.cell(row=row, column=5).font = Font(bold=True)
        ws.cell(row=row, column=5).fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")

        # --- Probability-Weighted Synergies row ---
        row += 1
        prob_weighted_row = row
        ws.cell(row=row, column=1, value="Probability-Weighted Synergies")
        ws.cell(row=row, column=1).font = Font(bold=True)
        for c in [2, 3, 4, 5]:
            cl = get_column_letter(c)
            ws.cell(row=row, column=c,
                    value=f"=SUMPRODUCT({cl}{syn_start}:{cl}{syn_end},F{syn_start}:F{syn_end})")
            ws.cell(row=row, column=c).number_format = '#,##0.0'
        ws.cell(row=row, column=5).font = Font(bold=True)

        # --- Pro Forma EBITDA row ---
        row += 1
        pf_ebitda_row = row
        ws.cell(row=row, column=1, value="Pro Forma EBITDA")
        ws.cell(row=row, column=1).font = Font(bold=True)
        for c in [2, 3, 4, 5]:
            cl = get_column_letter(c)
            ws.cell(row=row, column=c,
                    value=f"=B{base_ebitda_row}+{cl}{prob_weighted_row}")
            ws.cell(row=row, column=c).number_format = '#,##0.0'
        ws.cell(row=row, column=5).font = Font(bold=True)
        ws.cell(row=row, column=5).fill = PatternFill(
            start_color="90EE90", end_color="90EE90", fill_type="solid")

        # Update cell_map with new keys
        syn_map = self.cell_map.get('synergy_model', {})
        syn_map['base_ebitda_row'] = base_ebitda_row
        syn_map['syn_start'] = syn_start
        syn_map['syn_end'] = syn_end
        syn_map['total_cost_syn_row'] = total_cost_syn_row
        syn_map['prob_weighted_row'] = prob_weighted_row
        syn_map['pf_ebitda_row'] = pf_ebitda_row
        self.cell_map['synergy_model'] = syn_map

        ws.column_dimensions['A'].width = 30
        for c in ['B','C','D','E','F']: ws.column_dimensions[c].width = 12
        return self

    def add_carveout_analysis(self, data: Dict = None) -> 'TransactionModulesMixin':
        """Add Carve-out Analysis sheet."""
        ws = self.wb.create_sheet("Carve-out")
        self.sheets_created.append("Carve-out")
        a = self.assumptions
        data = data or {}

        self._add_title(ws, f"{self.company_name} - Carve-out Analysis", 1, 1)

        # --- Parent Revenue row referencing Operating Model when available ---
        om = self.cell_map.get('operating_model', {})
        parent_rev_row = 3
        ws.cell(row=parent_rev_row, column=1, value="Parent Revenue")
        if om and 'revenue_row' in om and 'ltm_col' in om:
            ltm_col_letter = get_column_letter(om['ltm_col'])
            ws.cell(row=parent_rev_row, column=2,
                    value=f"='Operating Model'!{ltm_col_letter}{om['revenue_row']}")
        else:
            ws.cell(row=parent_rev_row, column=2,
                    value=f"='Assumptions'!B5")
        ws.cell(row=parent_rev_row, column=2).number_format = '#,##0.0'
        ws.cell(row=parent_rev_row, column=2).font = Font(bold=True)

        self._add_section_header(ws, "STANDALONE COST ANALYSIS", 5, 1)

        headers = ["Category", "Allocated", "Standalone", "Variance"]
        row = 6
        for i, h in enumerate(headers):
            ws.cell(row=row, column=1+i, value=h)
        self._format_header_row(ws, row, 1, 4)
        row += 1

        costs = data.get('costs', [
            {'name': 'Corporate overhead', 'allocated': 5.0, 'standalone': 3.5},
            {'name': 'IT infrastructure', 'allocated': 3.0, 'standalone': 4.0},
            {'name': 'Finance/Accounting', 'allocated': 2.0, 'standalone': 2.5},
        ])
        cost_start = row
        for cost in costs:
            ws.cell(row=row, column=1, value=cost['name'])
            ws.cell(row=row, column=2, value=cost['allocated'])
            ws.cell(row=row, column=3, value=cost['standalone'])
            ws.cell(row=row, column=4, value=f"=C{row}-B{row}")  # Variance formula
            for c in [2,3]: self._format_input_cell(ws, row, c)
            for c in [2,3,4]: ws.cell(row=row, column=c).number_format = '#,##0.0'
            row += 1
        cost_end = row - 1

        row += 1
        dissynergy_row = row
        ws.cell(row=row, column=1, value="Dis-synergies")
        ws.cell(row=row, column=4, value=f"=SUM(D{cost_start}:D{cost_end})")
        ws.cell(row=row, column=4).number_format = '#,##0.0'
        ws.cell(row=row, column=4).font = Font(bold=True)

        # --- Standalone Economics section ---
        row += 2
        self._add_section_header(ws, "STANDALONE ECONOMICS", row, 1)
        row += 1

        standalone_rev_row = row
        ws.cell(row=row, column=1, value="Standalone Revenue")
        standalone_rev = data.get('standalone_revenue', 50.0)
        ws.cell(row=row, column=2, value=standalone_rev)
        ws.cell(row=row, column=2).number_format = '#,##0.0'
        self._format_input_cell(ws, row, 2)

        row += 1
        dissyn_pct_row = row
        ws.cell(row=row, column=1, value="Dis-synergy as % of Revenue")
        ws.cell(row=row, column=2,
                value=f"=IFERROR(D{dissynergy_row}/B{standalone_rev_row},0)")
        ws.cell(row=row, column=2).number_format = '0.0%'

        row += 1
        margin_input_row = row
        ws.cell(row=row, column=1, value="Standalone EBITDA Margin")
        standalone_margin = data.get('standalone_ebitda_margin', 0.20)
        ws.cell(row=row, column=2, value=standalone_margin)
        ws.cell(row=row, column=2).number_format = '0.0%'
        self._format_input_cell(ws, row, 2)

        row += 1
        standalone_ebitda_row = row
        ws.cell(row=row, column=1, value="Standalone EBITDA")
        ws.cell(row=row, column=2,
                value=f"=B{standalone_rev_row}*B{margin_input_row}")
        ws.cell(row=row, column=2).number_format = '#,##0.0'
        ws.cell(row=row, column=2).font = Font(bold=True)
        ws.cell(row=row, column=2).fill = PatternFill(
            start_color="90EE90", end_color="90EE90", fill_type="solid")

        # Update cell_map with new keys
        carveout_map = self.cell_map.get('carveout', {})
        carveout_map['parent_rev_row'] = parent_rev_row
        carveout_map['dissynergy_row'] = dissynergy_row
        carveout_map['standalone_rev_row'] = standalone_rev_row
        carveout_map['dissyn_pct_row'] = dissyn_pct_row
        carveout_map['standalone_ebitda_row'] = standalone_ebitda_row
        self.cell_map['carveout'] = carveout_map

        ws.column_dimensions['A'].width = 28
        for c in ['B','C','D']: ws.column_dimensions[c].width = 15
        return self

    def add_earnout_model(self, data: Dict = None) -> 'TransactionModulesMixin':
        """Add Earnout/Contingent Consideration modeling sheet."""
        ws = self.wb.create_sheet("Earnout Model")
        self.sheets_created.append("Earnout Model")
        data = data or {}

        self._add_title(ws, f"{self.company_name} - Earnout Analysis", 1, 1)
        self._add_section_header(ws, "EARNOUT STRUCTURE", 3, 1)

        # Inputs (rows 4-5)
        # Upfront Consideration: reference Sources & Uses if available
        su = self.cell_map.get('sources_uses', {})
        ws.cell(row=4, column=1, value="Upfront Consideration")
        if su and 'purchase_ev' in su:
            ws.cell(row=4, column=2,
                    value=f"='Sources & Uses'!{su['purchase_ev']}")
        else:
            ws.cell(row=4, column=2, value=data.get('upfront', 150.0))
            self._format_input_cell(ws, 4, 2)
        ws.cell(row=4, column=2).number_format = '#,##0.0'

        ws.cell(row=5, column=1, value="Maximum Earnout")
        ws.cell(row=5, column=2, value=data.get('max_earnout', 50.0))
        ws.cell(row=5, column=2).number_format = '#,##0.0'
        self._format_input_cell(ws, 5, 2)

        self._add_section_header(ws, "SCENARIO ANALYSIS", 7, 1)
        headers = ["Scenario", "Prob.", "Payout", "Weighted"]
        for i, h in enumerate(headers):
            ws.cell(row=8, column=1+i, value=h)
        self._format_header_row(ws, 8, 1, 4)

        scenarios = data.get('scenarios', [
            {'name': 'Exceed', 'prob': 0.20, 'payout': 50.0},
            {'name': 'Meet', 'prob': 0.45, 'payout': 35.0},
            {'name': 'Partial', 'prob': 0.25, 'payout': 15.0},
            {'name': 'Miss', 'prob': 0.10, 'payout': 0.0},
        ])
        sc_start = 9
        for i, sc in enumerate(scenarios):
            row = sc_start + i
            ws.cell(row=row, column=1, value=sc['name'])
            ws.cell(row=row, column=2, value=sc['prob'])
            ws.cell(row=row, column=2).number_format = '0%'
            self._format_input_cell(ws, row, 2)
            ws.cell(row=row, column=3, value=sc['payout'])
            ws.cell(row=row, column=3).number_format = '#,##0.0'
            self._format_input_cell(ws, row, 3)
            ws.cell(row=row, column=4, value=f"=B{row}*C{row}")  # Weighted formula
            ws.cell(row=row, column=4).number_format = '#,##0.0'
        sc_end = sc_start + len(scenarios) - 1

        exp_row = sc_end + 2
        ws.cell(row=exp_row, column=1, value="Expected Earnout")
        ws.cell(row=exp_row, column=4, value=f"=SUM(D{sc_start}:D{sc_end})")
        ws.cell(row=exp_row, column=4).number_format = '#,##0.0'
        ws.cell(row=exp_row, column=4).font = Font(bold=True)

        ws.cell(row=exp_row + 1, column=1, value="Effective Purchase Price")
        ws.cell(row=exp_row + 1, column=4, value=f"=B4+D{exp_row}")
        ws.cell(row=exp_row + 1, column=4).number_format = '#,##0.0'
        ws.cell(row=exp_row + 1, column=4).font = Font(bold=True)
        ws.cell(row=exp_row + 1, column=4).fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")

        # --- NPV Section ---
        npv_header_row = exp_row + 3
        self._add_section_header(ws, "NPV ANALYSIS", npv_header_row, 1)

        discount_rate_row = npv_header_row + 1
        ws.cell(row=discount_rate_row, column=1, value="Discount Rate")
        discount_rate = data.get('earnout_discount_rate', 0.10)
        ws.cell(row=discount_rate_row, column=2, value=discount_rate)
        ws.cell(row=discount_rate_row, column=2).number_format = '0.0%'
        self._format_input_cell(ws, discount_rate_row, 2)

        npv_earnout_row = discount_rate_row + 1
        ws.cell(row=npv_earnout_row, column=1, value="NPV of Expected Earnout")
        # Simplified: 2-year weighted average timing
        ws.cell(row=npv_earnout_row, column=4,
                value=f"=D{exp_row}/(1+B{discount_rate_row})^2")
        ws.cell(row=npv_earnout_row, column=4).number_format = '#,##0.0'

        total_eff_cost_row = npv_earnout_row + 1
        ws.cell(row=total_eff_cost_row, column=1, value="Total Effective Cost (NPV-adjusted)")
        ws.cell(row=total_eff_cost_row, column=4,
                value=f"=B4+D{npv_earnout_row}")
        ws.cell(row=total_eff_cost_row, column=4).number_format = '#,##0.0'
        ws.cell(row=total_eff_cost_row, column=4).font = Font(bold=True)
        ws.cell(row=total_eff_cost_row, column=4).fill = PatternFill(
            start_color="90EE90", end_color="90EE90", fill_type="solid")

        # Update cell_map with new keys
        earnout_map = self.cell_map.get('earnout_model', {})
        earnout_map['exp_row'] = exp_row
        earnout_map['discount_rate_row'] = discount_rate_row
        earnout_map['npv_earnout_row'] = npv_earnout_row
        earnout_map['total_eff_cost_row'] = total_eff_cost_row
        self.cell_map['earnout_model'] = earnout_map

        ws.column_dimensions['A'].width = 32
        for c in ['B','C','D']: ws.column_dimensions[c].width = 12
        return self

    def add_purchase_price_allocation(self, data: Dict = None) -> 'TransactionModulesMixin':
        """Add Purchase Price Allocation (PPA) sheet."""
        ws = self.wb.create_sheet("PPA")
        self.sheets_created.append("PPA")
        a = self.assumptions
        data = data or {}

        self._add_title(ws, f"{self.company_name} - Purchase Price Allocation", 1, 1)

        self._add_section_header(ws, "ASSETS ACQUIRED AT FAIR VALUE", 3, 1)
        headers = ["Asset", "Book", "Step-Up", "Fair Value"]
        row = 4
        for i, h in enumerate(headers):
            ws.cell(row=row, column=1+i, value=h)
        self._format_header_row(ws, row, 1, 4)

        assets = data.get('assets', [
            {'name': 'Tangible assets', 'book': 50.0, 'stepup': 10.0},
            {'name': 'Customer relationships', 'book': 0.0, 'stepup': 60.0},
            {'name': 'Technology/IP', 'book': 5.0, 'stepup': 30.0},
            {'name': 'Trade name', 'book': 0.0, 'stepup': 15.0},
        ])
        asset_start = 5
        for i, asset in enumerate(assets):
            row = asset_start + i
            ws.cell(row=row, column=1, value=asset['name'])
            ws.cell(row=row, column=2, value=asset['book'])
            ws.cell(row=row, column=3, value=asset['stepup'])
            ws.cell(row=row, column=4, value=f"=B{row}+C{row}")  # FV formula
            for c in [2,3]: self._format_input_cell(ws, row, c)
            for c in [2,3,4]: ws.cell(row=row, column=c).number_format = '#,##0.0'
        asset_end = asset_start + len(assets) - 1

        liabilities = data.get('liabilities', 30.0)
        t_row = asset_end + 2
        ws.cell(row=t_row, column=1, value="Total Identifiable Assets")
        ws.cell(row=t_row, column=4, value=f"=SUM(D{asset_start}:D{asset_end})")
        ws.cell(row=t_row, column=4).number_format = '#,##0.0'

        ws.cell(row=t_row + 1, column=1, value="Less: Liabilities")
        ws.cell(row=t_row + 1, column=4, value=-liabilities)
        ws.cell(row=t_row + 1, column=4).number_format = '(#,##0.0)'
        self._format_input_cell(ws, t_row + 1, 4)

        ws.cell(row=t_row + 2, column=1, value="Net Identifiable Assets")
        ws.cell(row=t_row + 2, column=4, value=f"=D{t_row}+D{t_row+1}")
        ws.cell(row=t_row + 2, column=4).number_format = '#,##0.0'

        pp_row = t_row + 4
        ws.cell(row=pp_row, column=1, value="Purchase Price")
        # Use Sources & Uses purchase_ev when available (more accurate, includes fees)
        su = self.cell_map.get('sources_uses', {})
        if su and 'purchase_ev' in su:
            ws.cell(row=pp_row, column=4,
                    value=f"='Sources & Uses'!{su['purchase_ev']}")
        else:
            ws.cell(row=pp_row, column=4,
                    value=f"='Assumptions'!B6*'Assumptions'!B7")
        ws.cell(row=pp_row, column=4).number_format = '#,##0.0'

        goodwill_row = pp_row + 1
        ws.cell(row=goodwill_row, column=1, value="Goodwill")
        ws.cell(row=goodwill_row, column=4, value=f"=D{pp_row}-D{t_row+2}")
        ws.cell(row=goodwill_row, column=4).number_format = '#,##0.0'
        ws.cell(row=goodwill_row, column=4).font = Font(bold=True)
        ws.cell(row=goodwill_row, column=4).fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")

        # --- Amortization Schedule section ---
        amort_header_row = goodwill_row + 2
        self._add_section_header(ws, "AMORTIZATION SCHEDULE", amort_header_row, 1)

        amort_col_headers = ["Asset", "Fair Value", "Useful Life (Yrs)", "Annual Amort"]
        # Use columns D, E, F, G offset to columns 1-4 (reuse same alignment as above)
        amort_hdr_row = amort_header_row + 1
        for i, h in enumerate(amort_col_headers):
            ws.cell(row=amort_hdr_row, column=1+i, value=h)
        self._format_header_row(ws, amort_hdr_row, 1, 4)

        # Identify intangible assets for amortization
        # Default useful lives (can be overridden via data)
        intangible_defaults = data.get('intangible_amort', [
            {'name': 'Customer relationships', 'useful_life': 15},
            {'name': 'Technology/IP', 'useful_life': 7},
            {'name': 'Trade name', 'useful_life': 10},
        ])

        amort_start = amort_hdr_row + 1
        amort_rows = []
        for i, intangible in enumerate(intangible_defaults):
            r = amort_start + i
            amort_rows.append(r)
            ws.cell(row=r, column=1, value=intangible['name'])

            # Find the corresponding asset row for Fair Value reference
            # Search assets list by name to get the correct row reference
            matching_asset_row = None
            for j, asset in enumerate(assets):
                if asset['name'] == intangible['name']:
                    matching_asset_row = asset_start + j
                    break

            if matching_asset_row is not None:
                # Reference the Fair Value column (D) from the asset row above
                ws.cell(row=r, column=2, value=f"=D{matching_asset_row}")
            else:
                # Fallback: use step-up value from data if no match
                ws.cell(row=r, column=2, value=intangible.get('fair_value', 0.0))
                self._format_input_cell(ws, r, 2)
            ws.cell(row=r, column=2).number_format = '#,##0.0'

            ws.cell(row=r, column=3, value=intangible['useful_life'])
            self._format_input_cell(ws, r, 3)
            ws.cell(row=r, column=3).number_format = '0'

            # Annual Amort = Fair Value / Useful Life
            ws.cell(row=r, column=4, value=f"=IFERROR(B{r}/C{r},0)")
            ws.cell(row=r, column=4).number_format = '#,##0.0'

        amort_end = amort_start + len(intangible_defaults) - 1

        # Total Annual Amortization
        total_amort_row = amort_end + 1
        ws.cell(row=total_amort_row, column=1, value="Total Annual Amortization")
        ws.cell(row=total_amort_row, column=1).font = Font(bold=True)
        ws.cell(row=total_amort_row, column=4,
                value=f"=SUM(D{amort_start}:D{amort_end})")
        ws.cell(row=total_amort_row, column=4).number_format = '#,##0.0'
        ws.cell(row=total_amort_row, column=4).font = Font(bold=True)

        # Annual Tax Shield = Total Amort x Tax Rate
        tax_shield_row = total_amort_row + 1
        ws.cell(row=tax_shield_row, column=1, value="Annual Tax Shield")
        ws.cell(row=tax_shield_row, column=4,
                value=f"=D{total_amort_row}*'Assumptions'!B24")
        ws.cell(row=tax_shield_row, column=4).number_format = '#,##0.0'
        ws.cell(row=tax_shield_row, column=4).font = Font(bold=True)
        ws.cell(row=tax_shield_row, column=4).fill = PatternFill(
            start_color="90EE90", end_color="90EE90", fill_type="solid")

        # Update cell_map with new keys
        ppa_map = self.cell_map.get('ppa', {})
        ppa_map['pp_row'] = pp_row
        ppa_map['goodwill_row'] = goodwill_row
        ppa_map['total_amort_row'] = total_amort_row
        ppa_map['tax_shield_row'] = tax_shield_row
        self.cell_map['ppa'] = ppa_map

        ws.column_dimensions['A'].width = 25
        for c in ['B','C','D']: ws.column_dimensions[c].width = 15
        return self
