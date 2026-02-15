#!/usr/bin/env python3
"""
Modular Excel Financial Model Generator

Creates individual model components that can be combined based on:
- Time available (24hr, 48hr, 7-day case)
- Analysis depth required
- Specific module requests

Based on PE textbook frameworks (Pignataro, Rosenbaum & Pearl, Zeisberger)

This file is a facade that assembles the full ExcelModelGenerator class
from mixin modules and re-exports all public symbols for backward compatibility.
"""

from typing import Dict

# Import base infrastructure
from ._base import (
    # Style constants
    INPUT_FILL, CALC_FILL, HEADER_FILL, SUBTOTAL_FILL, TOTAL_FILL,
    INPUT_FONT, CALC_FONT, HEADER_FONT, TITLE_FONT, SECTION_FONT,
    THIN_BORDER, BOTTOM_BORDER, DOUBLE_BORDER,
    # Enums and data
    ModelDepth, MODEL_MODULES, CASE_FRAMEWORKS, ModelAssumptions,
    # Base class
    ExcelModelGeneratorBase,
    # Helper functions
    list_model_modules, list_case_frameworks, get_framework_recommendation,
)

# Import all mixin classes
from ._core_modules import CoreModulesMixin
from ._dd_modules import DDModulesMixin
from ._capital_structure import CapitalStructureMixin
from ._transaction_modules import TransactionModulesMixin
from ._value_creation import ValueCreationMixin
from ._specialized_modules import SpecializedModulesMixin
from ._cfa_modules import CFAModulesMixin


# ============================================================
# ASSEMBLED CLASS
# ============================================================

class ExcelModelGenerator(
    CoreModulesMixin,
    DDModulesMixin,
    CapitalStructureMixin,
    TransactionModulesMixin,
    ValueCreationMixin,
    SpecializedModulesMixin,
    CFAModulesMixin,
    ExcelModelGeneratorBase,
):
    """
    Modular Excel model generator.

    Example:
        gen = ExcelModelGenerator("Acme Corp", depth=ModelDepth.STANDARD)
        gen.add_sources_uses()
        gen.add_operating_model()
        gen.add_debt_schedule()
        gen.add_returns_analysis()
        gen.save("acme_model.xlsx")
    """
    pass


# ============================================================
# QUICK GENERATION FUNCTIONS
# ============================================================

def generate_quick_lbo(company_name: str, output_path: str, assumptions: Dict = None) -> str:
    """Generate a quick 24-hour LBO model."""
    a = ModelAssumptions(company_name=company_name)
    if assumptions:
        for key, value in assumptions.items():
            if hasattr(a, key):
                setattr(a, key, value)

    gen = ExcelModelGenerator(company_name, ModelDepth.QUICK, a)
    gen.add_assumptions_sheet()
    gen.add_sources_uses()
    gen.add_operating_model()
    gen.add_returns_analysis()
    gen.add_sensitivity_tables()
    return gen.save(output_path)


def generate_standard_lbo(company_name: str, output_path: str, assumptions: Dict = None) -> str:
    """Generate a standard 48-hour to 5-day LBO model."""
    a = ModelAssumptions(company_name=company_name)
    if assumptions:
        for key, value in assumptions.items():
            if hasattr(a, key):
                setattr(a, key, value)

    gen = ExcelModelGenerator(company_name, ModelDepth.STANDARD, a)
    gen.add_assumptions_sheet()
    gen.add_sources_uses()
    gen.add_revenue_build()
    gen.add_expense_build()
    gen.add_operating_model()
    gen.add_debt_schedule()
    gen.add_working_capital()
    gen.add_returns_analysis()
    gen.add_sensitivity_tables()
    return gen.save(output_path)


def generate_comprehensive_lbo(company_name: str, output_path: str, assumptions: Dict = None) -> str:
    """Generate a comprehensive 7+ day LBO model with full institutional modules."""
    a = ModelAssumptions(company_name=company_name)
    if assumptions:
        for key, value in assumptions.items():
            if hasattr(a, key):
                setattr(a, key, value)

    gen = ExcelModelGenerator(company_name, ModelDepth.COMPREHENSIVE, a)

    # Assumptions first (central input sheet)
    gen.add_assumptions_sheet()

    # Core modules
    gen.add_sources_uses()
    gen.add_revenue_build()
    gen.add_expense_build()
    gen.add_operating_model()
    gen.add_debt_schedule()
    gen.add_working_capital()
    gen.add_wacc_calculation()
    gen.add_returns_analysis()
    gen.add_sensitivity_tables()

    # Institutional modules (7+ day case)
    gen.add_scenario_analysis()
    gen.add_management_vs_buyer()
    gen.add_dcf_valuation()
    gen.add_covenant_analysis()

    return gen.save(output_path)


def generate_debt_schedule(company_name: str, output_path: str, assumptions: Dict = None) -> str:
    """Generate standalone debt schedule."""
    a = ModelAssumptions(company_name=company_name)
    if assumptions:
        for key, value in assumptions.items():
            if hasattr(a, key):
                setattr(a, key, value)

    gen = ExcelModelGenerator(company_name, ModelDepth.STANDARD, a)
    gen.add_debt_schedule()
    return gen.save(output_path)


def generate_wacc_model(company_name: str, output_path: str, assumptions: Dict = None) -> str:
    """Generate standalone WACC calculation."""
    a = ModelAssumptions(company_name=company_name)
    if assumptions:
        for key, value in assumptions.items():
            if hasattr(a, key):
                setattr(a, key, value)

    gen = ExcelModelGenerator(company_name, ModelDepth.STANDARD, a)
    gen.add_wacc_calculation()
    return gen.save(output_path)


def generate_revenue_build(company_name: str, output_path: str, assumptions: Dict = None) -> str:
    """Generate standalone revenue build."""
    a = ModelAssumptions(company_name=company_name)
    if assumptions:
        for key, value in assumptions.items():
            if hasattr(a, key):
                setattr(a, key, value)

    gen = ExcelModelGenerator(company_name, ModelDepth.STANDARD, a)
    gen.add_revenue_build()
    return gen.save(output_path)


def generate_expense_build(company_name: str, output_path: str, assumptions: Dict = None) -> str:
    """Generate standalone expense/SG&A build."""
    a = ModelAssumptions(company_name=company_name)
    if assumptions:
        for key, value in assumptions.items():
            if hasattr(a, key):
                setattr(a, key, value)

    gen = ExcelModelGenerator(company_name, ModelDepth.STANDARD, a)
    gen.add_expense_build()
    return gen.save(output_path)


if __name__ == "__main__":
    print("=" * 60)
    print("Modular Excel Model Generator")
    print("=" * 60)

    print("\nAvailable Modules:")
    for key, info in MODEL_MODULES.items():
        print(f"  {key}: {info['name']}")

    print("\nCase Frameworks:")
    for key, info in CASE_FRAMEWORKS.items():
        print(f"\n  {info['name']}:")
        print(f"    Required: {', '.join(info['required_modules'])}")
        print(f"    Depth: {info['depth'].value}")
