# Excel Financial Model Reference Guide

## Time-Based Model Selection

### 24-Hour Sprint Case
**Focus:** Essentials only - get to a decision fast

| Module | Include? | Depth |
|--------|----------|-------|
| Sources & Uses | ✅ Required | Simple table |
| Operating Model | ✅ Required | Revenue → EBITDA (5 lines) |
| Debt Schedule | ❌ Skip | Use single debt assumption |
| Returns Analysis | ✅ Required | MOIC + IRR only |
| Sensitivity | ⚡ If time | Entry/exit matrix only |

**Tips:**
- Use management projections directly
- Single debt tranche, flat paydown
- Don't build detailed schedules
- One sensitivity table maximum

**Generate:**
```python
generate_quick_lbo("Company", "output.xlsx")
```

---

### 48-Hour Standard Case
**Focus:** Solid analysis with key supporting detail

| Module | Include? | Depth |
|--------|----------|-------|
| Sources & Uses | ✅ Required | With fees breakdown |
| Operating Model | ✅ Required | Full P&L through EBITDA |
| Revenue Build | ⚡ If time | Segment breakdown |
| Debt Schedule | ✅ Required | Multi-tranche |
| Working Capital | ⚡ If time | % of revenue |
| Returns Analysis | ✅ Required | By exit year + bridge |
| Sensitivity | ✅ Required | 2-3 matrices |

**Tips:**
- Challenge 1-2 management assumptions
- Build debt schedule with amortization
- Simple value creation bridge
- Base/upside/downside scenarios

**Generate:**
```python
gen = ExcelModelGenerator("Company", ModelDepth.STANDARD)
gen.add_sources_uses()
gen.add_operating_model()
gen.add_debt_schedule()
gen.add_returns_analysis()
gen.add_sensitivity_tables()
gen.save("output.xlsx")
```

---

### 5-Day Detailed Case
**Focus:** Comprehensive analysis with full supporting detail

| Module | Include? | Depth |
|--------|----------|-------|
| Sources & Uses | ✅ Required | Full breakdown |
| Revenue Build | ✅ Required | Segment drivers |
| Expense Build | ✅ Required | COGS + SG&A breakdown |
| Operating Model | ✅ Required | Segment-level P&L |
| Debt Schedule | ✅ Required | Cash sweep, multi-tranche |
| Working Capital | ✅ Required | Days-based calculation |
| WACC | ⚡ If DCF | Full CAPM build |
| Returns Analysis | ✅ Required | Full attribution |
| Sensitivity | ✅ Required | Multiple matrices |
| Covenant Analysis | ⚡ If relevant | Leverage + coverage |

**Tips:**
- Build revenue from segment drivers
- Detailed SG&A with categories
- Cash sweep mechanics in debt schedule
- Full value creation attribution
- Management vs. buyer case comparison

**Generate:**
```python
generate_standard_lbo("Company", "output.xlsx")
```

---

### 7+ Day Institutional Case
**Focus:** Full institutional-quality model

| Module | Include? | Depth |
|--------|----------|-------|
| All Standard modules | ✅ Required | Comprehensive |
| Revenue Build | ✅ Required | Unit economics, cohorts |
| Expense Build | ✅ Required | Headcount model |
| Debt Schedule | ✅ Required | Revolver, PIK, covenants |
| WACC | ✅ Required | Unlevering/relevering |
| DCF Model | ✅ Required | Multiple terminal methods |
| Trading Comps | ✅ Required | Full analysis |
| Transaction Comps | ✅ Required | Control premium |
| Covenant Analysis | ✅ Required | Full package |
| Synergy Model | ⚡ If M&A | Bottom-up build |
| 3-Statement | ⚡ If requested | Fully circular |

**Generate:**
```python
generate_comprehensive_lbo("Company", "output.xlsx")
```

---

## Module Quick Reference

### Sources & Uses (`add_sources_uses`)
```
Sources              Uses
────────────────     ────────────────
Senior Debt          Purchase Price
Sub Debt             Refinance Debt
Rollover Equity      Transaction Fees
Sponsor Equity       Financing Fees
```

### Revenue Build (`add_revenue_build`)
- **Quick:** Total revenue with growth rate
- **Standard:** Segment breakdown with individual growth
- **Comprehensive:** Unit economics, customer cohorts, pricing

### Expense Build (`add_expense_build`)
- **Quick:** COGS and SG&A as % of revenue
- **Standard:** Line-item breakdown (materials, labor, S&M, G&A, R&D)
- **Comprehensive:** Headcount model, vendor analysis

### Debt Schedule (`add_debt_schedule`)
- **Quick:** Single tranche, flat annual paydown
- **Standard:** Multi-tranche (Senior + Sub), mandatory amortization
- **Comprehensive:** Cash sweep, revolver, PIK toggle, covenant tracking

### WACC Calculation (`add_wacc_calculation`)
- **Quick:** Simple assumption (e.g., 10%)
- **Standard:** Full CAPM: Rf + β(ERP) + Size Premium
- **Comprehensive:** Unlevered beta, relevering, country risk

### Working Capital (`add_working_capital`)
- **Quick:** NWC as % of revenue
- **Standard:** AR/AP/Inventory days calculation
- **Comprehensive:** Seasonal patterns, normalization, CCC analysis

### Returns Analysis (`add_returns_analysis`)
- **Quick:** MOIC and IRR by exit year
- **Standard:** Value creation bridge (EBITDA growth, multiple, debt paydown)
- **Comprehensive:** Full attribution, management vs. buyer case

### Sensitivity Tables (`add_sensitivity_tables`)
- **Quick:** Entry vs. Exit multiple matrix
- **Standard:** Multiple matrices (growth, margin, leverage)
- **Comprehensive:** Monte Carlo inputs, tornado chart data

### Scenario Analysis (`add_scenario_analysis`) - Institutional
- Bull/Bear/Base case comparison
- Probability-weighted returns calculation
- Downside protection analysis
- Sensitivity by scenario

### Management vs Buyer (`add_management_vs_buyer`) - Institutional
- Side-by-side projection comparison
- Revenue and EBITDA variance analysis
- Returns comparison under each case
- Haircut rationale documentation

### DCF Valuation (`add_dcf_valuation`) - Institutional
- Unlevered free cash flow build
- Terminal value calculation (Gordon Growth / Exit Multiple)
- Present value calculation
- Implied multiples and sanity checks

### Covenant Analysis (`add_covenant_analysis`) - Institutional
- Leverage ratio tracking (Total Debt / EBITDA)
- Interest coverage tracking (EBITDA / Interest)
- Covenant compliance status
- Headroom analysis

### CFA-Derived Modules

#### Reverse DCF (`add_reverse_dcf`)
- Market-implied growth rates from current share price
- Implied vs forecast comparison table

#### DuPont Analysis (`add_dupont_analysis`)
- 3-component and 5-component ROE decomposition
- Trend analysis with driver attribution

#### ROIC Decomposition (`add_roic_decomposition`)
- ROIC vs WACC spread calculation
- Economic profit and invested capital turnover

#### DDM Valuation (`add_ddm_valuation`)
- Two-stage and H-model dividend discount
- Sensitivity grid (Cost of Equity × Terminal Growth)

#### Earnings Quality (`add_earnings_quality`)
- Beneish M-Score, Altman Z-Score, Piotroski F-Score
- Conditional formatting dashboard

#### Tornado Sensitivity (`add_tornado_sensitivity`)
- Single-variable sensitivity with IRR impact ranking
- Auto-sorted horizontal bar data for slide generation

#### Football Field (`add_football_field`)
- Valuation range comparison across methodologies
- Low/Mid/High per method with weighted blend

#### SOTP Valuation (`add_sotp_valuation`)
- Sum-of-the-parts by business segment
- Conglomerate discount sensitivity

#### Enhanced WACC (`add_enhanced_wacc`)
- 3-method cost of debt, multi-method beta
- Country risk premium, size premium, illiquidity premium

#### Geographic Terminal Growth (`add_geographic_terminal_growth`)
- GDP-weighted terminal growth by geography
- Reasonability check vs long-term inflation

#### Multi-Stage DCF (`add_multi_stage_dcf`)
- 3-stage DCF: explicit → transition → terminal
- Linear growth decline with EV bridge

#### Source Index (`add_source_index`)
- Canonical repository of all metrics flowing to slides
- Numbered metrics with source links and data vintage

### Due Diligence Modules

#### Quality of Earnings (`add_quality_of_earnings`)
- EBITDA normalization with adjustments bridge
- Run-rate analysis with pro forma earnings

#### Working Capital Normalization (`add_working_capital_normalization`)
- Component-level days analysis
- NWC peg mechanism with seasonal patterns

#### Customer & Revenue Quality (`add_customer_revenue_quality`)
- Customer concentration analysis (top 10)
- Retention cohorts, LTV, unit economics

#### Credit & Debt Sizing (`add_credit_debt_sizing`)
- Credit ratios and agency benchmarks
- Debt capacity with stress test scenarios

### Capital Structure Modules

#### Dividend Recap (`add_dividend_recap`)
- Mid-hold dividend sizing with leverage constraints
- Returns impact analysis (before/after recap)

#### Refinancing Analysis (`add_refinancing_analysis`)
- Rate comparison with NPV of savings
- Prepayment penalty cost-benefit

#### Cap Table Waterfall (`add_cap_table_waterfall`)
- Equity distribution with preferred return and catch-up
- Multi-class equity with MIP and co-invest

#### Sponsor Economics (`add_sponsor_economics`)
- GP carry and management fee calculation
- Fund-level returns with clawback analysis

### Transaction Structure Modules

#### Add-on Analysis (`add_addon_analysis`)
- Platform + bolt-on combined model
- Synergy quantification and accretion analysis

#### Synergy Model (`add_synergy_model`)
- Revenue and cost synergy phasing
- Implementation costs and net NPV

#### Carve-out Analysis (`add_carveout_analysis`)
- Standalone P&L with TSA costs
- Stranded cost identification and phase-out

#### Earnout Model (`add_earnout_model`)
- Performance milestone tracking
- Probability-weighted earnout valuation

#### Purchase Price Allocation (`add_purchase_price_allocation`)
- Intangibles identification with useful lives
- Goodwill residual and deferred tax impact

### Value Creation Modules

#### Value Creation Bridge (`add_value_creation_bridge`)
- EBITDA growth, multiple expansion, deleveraging attribution
- Organic vs. inorganic decomposition

#### 100-Day Plan (`add_hundred_day_plan`)
- Post-close priorities with phased milestones
- Owner assignment and KPI tracking

#### Exit Readiness (`add_exit_readiness`)
- Exit option comparison and timing analysis
- Readiness scorecard with gap identification

#### Management Incentive Plan (`add_management_incentive_plan`)
- MIP structure with vesting schedule
- Payout scenarios with ratchets and sweet equity

### Specialized Modules

#### Rollup Model (`add_rollup_model`)
- Multi-acquisition pipeline tracking
- Multiple-arbitrage and combined metrics

#### Tax Analysis (`add_tax_analysis`)
- NOL utilization schedule
- Step-up analysis with tax shield quantification

#### Control Premium Analysis (`add_control_premium_analysis`)
- Premium to unaffected price (30-day VWAP)
- Historical precedent premium comparison

---

## Prompt-to-Module Mapping

| User Request | Module | Quick Function |
|--------------|--------|----------------|
| "debt schedule", "debt paydown", "amortization" | `debt_schedule` | `generate_debt_schedule()` |
| "WACC", "cost of capital", "discount rate" | `wacc_calculation` | `generate_wacc_model()` |
| "revenue build", "segment analysis", "top line" | `revenue_build` | `generate_revenue_build()` |
| "SG&A", "expense build", "cost structure" | `expense_build` | `generate_expense_build()` |
| "working capital", "NWC", "AR/AP" | `working_capital` | via ExcelModelGenerator |
| "sources and uses", "S&U", "funding" | `sources_uses` | via ExcelModelGenerator |
| "returns", "MOIC", "IRR" | `returns_analysis` | via ExcelModelGenerator |
| "sensitivity", "matrix", "scenarios" | `sensitivity_tables` | via ExcelModelGenerator |
| "quick LBO", "24-hour model" | all essentials | `generate_quick_lbo()` |
| "full LBO model", "detailed model" | all modules | `generate_comprehensive_lbo()` |

### Institutional Modules (7+ Day Case)

| User Request | Module | Quick Function |
|--------------|--------|----------------|
| "scenario analysis", "bull/bear/base", "probability weighted" | `scenario_analysis` | via ExcelModelGenerator |
| "management vs buyer", "case comparison", "haircut analysis" | `management_vs_buyer` | via ExcelModelGenerator |
| "DCF model", "discounted cash flow", "terminal value" | `dcf_valuation` | via ExcelModelGenerator |
| "covenant analysis", "leverage covenant", "coverage covenant" | `covenant_analysis` | via ExcelModelGenerator |
| "institutional model", "7-day model", "comprehensive LBO" | all + institutional | `generate_comprehensive_lbo()` |

### CFA Enrichment Modules

| User Request | Module | Function |
|--------------|--------|----------|
| "reverse DCF", "implied growth" | `reverse_dcf` | `add_reverse_dcf()` |
| "DuPont", "ROE decomposition" | `dupont_analysis` | `add_dupont_analysis()` |
| "ROIC", "economic profit", "invested capital" | `roic_decomposition` | `add_roic_decomposition()` |
| "DDM", "dividend discount" | `ddm_valuation` | `add_ddm_valuation()` |
| "earnings quality", "M-Score", "Z-Score", "F-Score" | `earnings_quality` | `add_earnings_quality()` |
| "tornado", "single variable sensitivity" | `tornado_sensitivity` | `add_tornado_sensitivity()` |
| "football field", "valuation range" | `football_field` | `add_football_field()` |
| "SOTP", "sum of the parts" | `sotp_valuation` | `add_sotp_valuation()` |
| "enhanced WACC", "multi-method beta" | `enhanced_wacc` | `add_enhanced_wacc()` |
| "geographic terminal growth", "GDP-weighted" | `geographic_terminal_growth` | `add_geographic_terminal_growth()` |
| "multi-stage DCF", "3-stage DCF" | `multi_stage_dcf` | `add_multi_stage_dcf()` |

### Due Diligence Modules

| User Request | Module | Function |
|--------------|--------|----------|
| "quality of earnings", "QoE", "EBITDA adjustments" | `quality_of_earnings` | `add_quality_of_earnings()` |
| "NWC normalization", "NWC peg", "target NWC" | `working_capital_normalization` | `add_working_capital_normalization()` |
| "customer quality", "concentration", "churn", "LTV" | `customer_revenue_quality` | `add_customer_revenue_quality()` |
| "credit analysis", "debt capacity", "stress test" | `credit_debt_sizing` | `add_credit_debt_sizing()` |

### Capital Structure Modules

| User Request | Module | Function |
|--------------|--------|----------|
| "dividend recap", "recapitalization" | `dividend_recap` | `add_dividend_recap()` |
| "refinancing", "rate savings" | `refinancing_analysis` | `add_refinancing_analysis()` |
| "cap table", "waterfall", "LP/GP" | `cap_table_waterfall` | `add_cap_table_waterfall()` |
| "sponsor economics", "carry", "GP economics" | `sponsor_economics` | `add_sponsor_economics()` |

### Transaction Structure Modules

| User Request | Module | Function |
|--------------|--------|----------|
| "add-on", "bolt-on" | `addon_analysis` | `add_addon_analysis()` |
| "synergies", "cost savings" | `synergy_model` | `add_synergy_model()` |
| "carve-out", "spin-off" | `carveout_analysis` | `add_carveout_analysis()` |
| "earnout", "contingent consideration" | `earnout_model` | `add_earnout_model()` |
| "PPA", "goodwill", "purchase price allocation" | `purchase_price_allocation` | `add_purchase_price_allocation()` |

### Value Creation Modules

| User Request | Module | Function |
|--------------|--------|----------|
| "value bridge", "value creation" | `value_creation_bridge` | `add_value_creation_bridge()` |
| "100-day plan", "post-close" | `hundred_day_plan` | `add_hundred_day_plan()` |
| "exit readiness", "exit planning" | `exit_readiness` | `add_exit_readiness()` |
| "MIP", "management equity" | `management_incentive_plan` | `add_management_incentive_plan()` |

### Specialized Modules

| User Request | Module | Function |
|--------------|--------|----------|
| "rollup", "platform build" | `rollup_model` | `add_rollup_model()` |
| "tax", "NOL", "step-up" | `tax_analysis` | `add_tax_analysis()` |
| "control premium", "takeover premium" | `control_premium_analysis` | `add_control_premium_analysis()` |

---

## Usage Examples

### Single Module
```python
from excel_modules import generate_debt_schedule

generate_debt_schedule(
    company_name="Acme Corp",
    output_path="debt_schedule.xlsx",
    assumptions={
        "ltm_ebitda": 50,
        "senior_debt_multiple": 4.0,
        "sub_debt_multiple": 1.5,
    }
)
```

### Composable Model
```python
from excel_modules import ExcelModelGenerator, ModelDepth, ModelAssumptions

# Custom assumptions
assumptions = ModelAssumptions(
    company_name="Target Co",
    ltm_revenue=200,
    ltm_ebitda=40,
    entry_multiple=9.0,
    exit_multiple=8.5,
)

# Build custom model
gen = ExcelModelGenerator("Target Co", ModelDepth.STANDARD, assumptions)
gen.add_sources_uses()
gen.add_revenue_build()
gen.add_debt_schedule()
gen.add_returns_analysis()
gen.save("custom_model.xlsx")
```

### Time-Based Generation
```python
from excel_modules import get_framework_recommendation

# Get recommendation based on hours
framework = get_framework_recommendation(hours_available=36)
print(f"Recommended: {framework['name']}")
print(f"Required modules: {framework['required_modules']}")
print(f"Tips: {framework['tips']}")
```

---

## Key PE Textbook Principles Applied

### From Rosenbaum & Pearl
- Sources & Uses always balances
- Entry multiple = EV / LTM EBITDA
- Senior debt sized on coverage and leverage

### From Pignataro
- EBITDA adjustments clearly separated
- Working capital = operating items only
- Capex = maintenance + growth

### From Zeisberger
- Value creation attribution: EBITDA growth, multiple expansion, debt paydown
- IRR sensitive to timing of cash flows
- MOIC more stable than IRR for comparison

---

## File Locations

```
ii-agent/src/ii_skills/ib_toolkit/templates/excel_models/
├── excel_modules.py           # ExcelModelGenerator assembly (USE THIS)
├── _base.py                   # Base class, MODEL_MODULES registry, CASE_FRAMEWORKS
├── _core_modules.py           # 14 core modules (S&U, OpModel, Debt, Returns, Scenario, DCF, etc.)
├── _cfa_modules.py            # 11 CFA modules (Reverse DCF, DuPont, ROIC, DDM, Earnings Quality, etc.)
├── _dd_modules.py             # 4 DD modules (QoE, NWC Normalization, Customer Quality, Credit)
├── _capital_structure.py      # 4 capital structure modules (Dividend Recap, Refi, Waterfall, Sponsor)
├── _transaction_modules.py    # 5 transaction modules (Add-on, Synergy, Carve-out, Earnout, PPA)
├── _value_creation.py         # 4 value creation modules (Bridge, 100-Day, Exit, MIP)
├── _specialized_modules.py    # 4 specialized modules (Rollup, Tax, Control Premium, Source Index)
├── formula_builder.py         # FormulaBuilder helper class
├── MODEL_REFERENCE.md         # This reference document
└── [generated models]         # Output files
```
