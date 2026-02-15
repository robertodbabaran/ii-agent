# Valuation Playbook
## Every Technique from 10 Years of CFA Research Challenge Winners (2016-2025)
*Cross-referenced against IB Toolkit — NEW items flagged for development*

---

## Master Valuation Method Frequency

| Method | Years Used | Frequency |
|--------|-----------|-----------|
| DCF (FCFF/FCFE) | ALL 10 | 100% |
| Relative / Comps (EV/EBITDA) | ALL 10 | 100% |
| Sensitivity (WACC × TGR grid) | ALL 10 | 100% |
| Scenario Analysis (Bull/Base/Bear) | ALL 10 | 100% |
| Monte Carlo Simulation | 8/10 | 80% |
| Football Field | 7/10 | 70% |
| DDM / Dividend Discount | 6/10 | 60% |
| SOTP (Sum of the Parts) | 4/10 | 40% |
| Reverse DCF | 4/10 | 40% |
| Relative Valuation Regression (ROE vs P/B) | 3/10 | 30% |
| PEG Multiple | 2/10 | 20% |
| Residual Income Model | 1/10 | 10% |
| Brownian Motion Price Path | 1/10 | 10% |

---

## 1. Discounted Cash Flow (DCF) — REQUIRED
**Status:** EXISTS in IB Toolkit (DCF Valuation module + Multi-Stage DCF + Geographic Terminal Growth)
**Enhancements BUILT (Sprint 4):** Multi-stage DCF (`add_multi_stage_dcf()`), Geographic terminal growth (`add_geographic_terminal_growth()`), Enhanced WACC (`add_enhanced_wacc()`)

### Execution Steps
1. Choose FCFF (enterprise) or FCFE (equity) — FCFF is standard for PE; FCFE for equity research
2. Set forecast horizon (5-7 years typical; 6-year most common in CFA reports)
3. Build revenue model bottom-up (segment × geography × driver)
4. Project operating costs, working capital, capex
5. Calculate unlevered free cash flow (UFCF) per year
6. Determine WACC (see WACC section below)
7. Calculate terminal value via perpetuity growth OR exit multiple (many teams use BOTH as cross-check)
8. Discount to present, bridge from EV to equity value per share

### CFA Best Practices (What Winners Do)
- **Multi-stage DCF** (2018 Lausanne): 3 stages — explicit forecast → linear growth decline → terminal
- **Geographic-weighted terminal growth** (2018, 2021): Weight country GDP growth by revenue exposure
- **Revenue-weighted risk-free rate** (2021 BI Norwegian): Weight government bond yields by revenue geography
- **Dual WACC** (2023 Sydney, 2025 Kozminski): Separate forecast WACC and terminal WACC (lower terminal reflects maturity)
- **Terminal value as % of EV disclosure** (2021: 83%, 2023: 87.8%) — transparency on terminal dependency
- **Always cross-check**: Implied exit multiple from terminal growth, and vice versa

### WACC Build-Up — Advanced Techniques

| Component | Standard | CFA Winner Enhancement |
|-----------|----------|----------------------|
| Risk-free rate | Single country | Revenue-weighted across geographies (2017, 2021) |
| Beta | 5Y monthly regression | Bottom-up unlevered peer average, re-levered (ALL years) |
| ERP | Single estimate | Blended: historical + survey + regression (2023 Sydney) |
| Cost of Debt | Weighted avg interest rate | Three-method: Bond YTM, synthetic credit rating, market rate (2025) |
| Size premium | Often omitted | Ten-decile analysis (2016 Waterloo) |
| Illiquidity premium | Often omitted | Bid-ask spread calibrated to academic literature (2025 Kozminski: 40bps) |
| Country risk premium | Often omitted | Revenue-weighted across operating geographies (2017 Copa) |

**BUILT (Sprint 4):**
- [x] **DONE: Multi-stage DCF template** → `add_multi_stage_dcf()` (3-stage with linear growth decline)
- [x] **DONE: Geographic-weighted terminal growth calculator** → `add_geographic_terminal_growth()`
- [x] **DONE: Enhanced WACC module** → `add_enhanced_wacc()` (3-method Kd, multi-method beta, CRP, size premium)

**REMAINING:**
- [ ] **NEW: Add revenue-weighted risk-free rate calculator**
- [ ] **NEW: Add dual WACC (forecast vs terminal) option**
- [ ] **NEW: Add illiquidity premium calculator** (bid-ask spread method)

---

## 2. Relative Valuation / Comparable Companies — REQUIRED
**Status:** EXISTS in IB Toolkit (Valuation slides)
**Gap:** Missing deep-dive comps analysis, premium/discount justification, regression-based fair value

### Execution Steps
1. Select peer universe (3-15 companies, geography-segmented)
2. Calculate forward multiples: EV/EBITDA (most common), P/E, EV/Sales, P/BV, PEG
3. Determine appropriate premium/discount to peers
4. Apply multiple to company's projected metric
5. Bridge from EV to equity value per share

### CFA Best Practices
- **Multiple peer groups by geography** (2023 Qantas): Europe 30%, North America 30%, APAC 40%
- **Historical premium/discount analysis** (2017 Copa, 2023 Qantas): 10-year tracking of company's premium to peers, then justify current position
- **ROE vs P/BV regression** (2020 CBA, 2025 Kozminski): Cross-sectional regression with R-squared; company plotted above/below line = over/undervalued
- **PEG ratio comparability** (2018 Lausanne): Compare company PEG to peer median — discount = undervalued
- **ESG premium justification** (2021 Vestas): ~10% premium to peer median EV/EBITDA for ESG leadership
- **Sector-appropriate multiples**: EV/EBITDAR for airlines (2017), EV/ARR for SaaS, P/FFO for REITs (2016)

**NEXT STEPS for IB Toolkit:**
- [ ] **NEW: Standalone comparable company analysis deep-dive slide** (currently only in valuation section)
- [ ] **NEW: Premium/discount justification framework** (historical tracking + fundamental drivers)
- [ ] **NEW: ROE vs P/BV regression slide template** (scatter plot with regression line)
- [ ] **NEW: Add PEG-based valuation module**

---

## 3. Monte Carlo Simulation — HIGHLY RECOMMENDED (80% frequency)
**Status:** NOT IN IB Toolkit
**Gap:** No Monte Carlo capability whatsoever

### Execution Steps
1. Identify 5-10 key variables to stress (revenue growth, margins, WACC, terminal growth, volumes, FX)
2. Assign probability distributions (Normal for most; Lognormal for impairments/skewed data)
3. Run 5,000-1,000,000 iterations
4. Output: histogram, mean, median, standard deviation, percentiles
5. Calculate probability of BUY/HOLD/SELL recommendations
6. Report: "X% probability of minimum Y% upside"

### CFA Best Practices
- **1,000,000 iterations** (2017 Copa, 2022 MSI) — highest confidence
- **10,000 iterations** (2020 CBA, 2025 Kozminski) — standard
- **5,000 iterations** (2023 Qantas) — minimum
- **Report kurtosis and skewness** (2021 Vestas) — shows distribution shape
- **On income statement** (2018 Lausanne) — stress revenue/costs, not just valuation
- **Full statistical summary** (mean, median, SD, percentiles, CoV)
- **Visual**: Histogram with color gradient (red=downside, green=upside), current price and target marked

**NEXT STEPS for IB Toolkit:**
- [ ] **NEW: Monte Carlo simulation module** (Python script using numpy)
- [ ] **NEW: Monte Carlo histogram slide template**
- [ ] **NEW: Monte Carlo statistical summary table template**
- This is the SINGLE HIGHEST-PRIORITY gap — 80% of winners use it

---

## 4. Sum-of-the-Parts (SOTP) — WHEN APPLICABLE
**Status:** BUILT (Sprint 3) — `add_sotp_valuation()` (Excel) + `add_sotp_waterfall_slide()` (Slide)

### When to Use
- Conglomerates / multi-segment businesses (2016 Canadian Tire: 5 business units)
- Mixed business models (2019 D&L: domestic + export split)
- Real estate developers with project-by-project visibility (2025 Dom: 60+ project NPVs)
- Companies with distinct business units warranting different multiples/methodologies

### Execution Steps
1. Identify separable business units/segments
2. Value each segment using the most appropriate method:
   - DCF for growing/mature segments
   - Multiples for segments with clean comps
   - P/FFO for REIT segments (2016)
   - Project NPV for real estate pipelines (2025)
   - Precedent transaction for unique segments (2016: CTFS at Scotiabank acquisition multiple)
3. Sum segment values → Enterprise Value
4. Apply conglomerate discount if appropriate (2016: analyzed but not applied)
5. Add cash, subtract debt → Equity Value

**BUILT (Sprint 3):**
- [x] **DONE: SOTP valuation slide template** → `add_sotp_waterfall_slide()` (waterfall showing segment contributions)
- [x] **DONE: SOTP Excel module** → `add_sotp_valuation()` (segment-by-segment with conglomerate discount sensitivity)

**REMAINING:**
- [ ] **NEW: Conglomerate discount analysis framework** (deeper standalone module)

---

## 5. Reverse DCF — POWERFUL CROSS-CHECK
**Status:** NOT IN IB Toolkit
**Gap:** No reverse DCF capability

### What It Does
Solves BACKWARDS: "What growth/margins does the market price imply?" Then compare to your forecast.

### Execution Steps
1. Take current market price as given
2. Use your WACC and terminal growth assumptions
3. Solve for the implied revenue growth rate (or margin, or ROIC) that equates DCF value to market price
4. Compare implied assumptions to your forecast
5. If market implies X% growth and you forecast Y% (where Y > X), stock is undervalued

### CFA Best Practices
- **2018 Lausanne**: Required 12% FCF growth to justify price; team forecasted 13% → buy
- **2023 Qantas**: Market pricing in load factor 2.3-4.6% BELOW minimum historical → mispricing
- **2024 Cargojet**: Market pricing 6.0% revenue CAGR; team forecasted 7.5% → undervalued
- **2025 Kozminski**: Market pricing 3% revenue CAGR; team forecasted 8% → undervalued
- **2020 CBA**: Solved for cash rate at which ROE = Ke (Effective Lower Bound) → novel variant

**NEXT STEPS for IB Toolkit:**
- [ ] **NEW: Reverse DCF module** (solve for implied growth rate at market price)
- [ ] **NEW: Reverse DCF slide template** (market-implied vs. forecast side-by-side)

---

## 6. Dividend Discount Model (DDM)
**Status:** NOT IN IB Toolkit (identified in gap analysis)
**Gap:** No DDM template

### When to Use
- Companies with established dividend history
- Financial services (banks, REITs, utilities)
- As cross-check even when assigned 0% weight (2018 Lausanne, 2025 Kozminski)

### Variants Seen
- **Two-stage DDM** (2024 Cargojet): High growth → stable growth
- **Multi-stage DDM** (standard): Explicit + transition + terminal
- **Reverse DDM** (2020 CBA): Solve for implied cost of equity → if market-implied Ke < fundamental Ke, stock is overvalued

**NEXT STEPS for IB Toolkit:**
- [ ] **NEW: DDM Excel module** (two-stage and multi-stage)
- [ ] **NEW: DDM slide template with sensitivity grid**

---

## 7. Residual Income Model (RIM)
**Status:** NOT IN IB Toolkit
**Gap:** No RIM capability

### When to Use
- Banks and financial institutions (2020 CBA used as PRIMARY method, 60% weight)
- Book value is meaningful anchor
- Avoids FCF estimation problems unique to banks

### Why for Banks
- Loan growth is both investment AND revenue — FCF definition breaks down
- ROE vs Ke spread determines franchise value
- Book value is regulatory capital, directly meaningful

**NEXT STEPS for IB Toolkit:**
- [ ] **NEW: Residual Income Model for financial services** (add to sector playbooks)

---

## 8. Scenario Analysis (Bull/Base/Bear) — REQUIRED
**Status:** EXISTS in IB Toolkit (Scenario Analysis module)
**Enhancement needed:** Structured thesis-linked scenarios, cascading impact chains

### CFA Best Practices
- **Each scenario linked to investment thesis pillars** (2021 Vestas): Bear = tech moat fails; Bull = tech moat widens
- **Cascading economic impact chains** (2016 Waterloo): O&G downturn → Alberta retail → segment revenue
- **Probability assignment** not always done but powerful when present
- **Per-thesis assumption table** (most years): Show exactly which assumptions change per scenario

**NEXT STEPS for IB Toolkit:**
- [ ] **UPDATE: Add thesis-linked scenario structure to existing module**
- [ ] **NEW: Cascading impact chain diagram template**

---

## 9. Football Field Chart — STANDARD
**Status:** BUILT (Sprint 3) — `add_football_field()` (Excel) + `add_football_field_slide()` (Slide)

### What It Shows
Horizontal bar chart: each row = valuation method, showing range (low to high) with target marked.

### Standard Rows
1. 52-week trading range
2. Broker consensus range
3. DCF range (bear to bull)
4. Relative valuation range (low to high peer multiple)
5. SOTP (if applicable)
6. DDM (if applicable)
7. Blended target price

**BUILT (Sprint 3):**
- [x] **DONE: Football field slide template** → `add_football_field_slide()` (scaled range bars with mid-markers)
- [x] **DONE: Football field Excel helper** → `add_football_field()` (Low/Mid/High per methodology, weighted blend)

---

## 10. Blended Valuation Weighting — BEST PRACTICE
**Status:** BUILT (Sprint 1) — `add_valuation_blending_slide()` (Slide)

### Weighting Patterns from Winners

| Year | Company | DCF Weight | Relative Weight | Other Weights |
|------|---------|-----------|----------------|---------------|
| 2017 | Copa | 30% FCFE | 70% EV/EBITDAR | — |
| 2018 | VAT Group | 75% DCF | 25% PEG | DDM 0% (cross-check) |
| 2019 | D&L | Primary DCF | Supporting comps | SOTP cross-check |
| 2020 | CBA | 60% RIM | 30% Relative | 10% DDM |
| 2023 | Qantas | 70% DCF | 30% EV/EBITDA | — |
| 2024 | Cargojet | 40% DCF-TG | 20% Comps | 40% DCF-Exit Multiple |
| 2025 | Dom Dev | 50% DCF | 50% Relative | SOTP, DDM as checks |

**Key Insight:** No single method is trusted alone. Always blend 2-3 methods with explicit weights.

### Weighting Formula
```
Target Price = w₁ × DCF + w₂ × Relative + w₃ × Other
Where: w₁ + w₂ + w₃ = 100%
```

**BUILT (Sprint 1):**
- [x] **DONE: Valuation blending framework** → `add_valuation_blending_slide()` (method table + equation bar + weight justification)

---

## Priority Development Queue (Ranked by Impact)

| Priority | Module | Frequency | Status |
|----------|--------|-----------|--------|
| 1 | Monte Carlo Simulation | 80% of winners | TODO |
| 2 | Football Field Chart | 70% of winners | **DONE** (Sprint 3) |
| 3 | Reverse DCF | 40% of winners | TODO |
| 4 | SOTP Framework | 40% of winners | **DONE** (Sprint 3) |
| 5 | DDM Template | 60% of winners | TODO |
| 6 | ROE vs P/BV Regression | 30% of winners | TODO |
| 7 | Multi-stage DCF Enhancement | 20% of winners | **DONE** (Sprint 4) |
| 8 | Residual Income Model | 10% (banks only) | TODO |
| 9 | Blended Valuation Framework | Best practice | **DONE** (Sprint 1) |
| 10 | Geographic-weighted WACC | Best practice | **DONE** (Sprint 4) |
