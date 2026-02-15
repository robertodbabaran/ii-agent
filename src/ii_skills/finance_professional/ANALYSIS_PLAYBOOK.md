# Financial Analysis Playbook
## Step-by-Step Techniques from 10 Years of CFA Research Challenge Winners (2016-2025)
*Cross-referenced against IB Toolkit — NEW items flagged for development*

---

## Master Analysis Frequency

| Analysis Type | Years Used | Status in IB Toolkit |
|---------------|-----------|---------------------|
| Revenue Decomposition (segment/geography) | 10/10 | EXISTS (Revenue Build) |
| DuPont Analysis | 7/10 | **NEW** |
| Peer Financial Benchmarking | 10/10 | Partial |
| Margin Trajectory Analysis | 10/10 | EXISTS (Operating Model) |
| Leverage / Debt Analysis | 10/10 | EXISTS (Credit Analysis) |
| Working Capital / Cash Cycle | 8/10 | EXISTS (Working Capital) |
| CapEx Analysis (growth vs maintenance) | 7/10 | Partial |
| EPS / Dividend Analysis | 8/10 | **NEW** |
| ROIC Analysis | 6/10 | **NEW** |
| Beneish M-Score | 2/10 | **NEW** |
| Altman Z-Score | 2/10 | **NEW** |
| Piotroski F-Score | 1/10 | **NEW** |
| Regression-based Revenue Forecasting | 2/10 | **NEW** |
| Correlation Analysis | 2/10 | **NEW** |

---

## 1. Revenue Decomposition — ALWAYS DO THIS FIRST
**Status:** EXISTS — Revenue Build module

### Step-by-Step
1. **Segment split**: Break revenue into business units (minimum 3 segments)
2. **Geography split**: Map revenue by country/region
3. **Driver identification**: For each segment, identify the specific growth driver:
   - Retail: same-store sales × store count (2016)
   - Airlines: ASKs × Load Factor × Yield (2017, 2023)
   - Industrial: Units delivered × ASP (2018, 2021, 2024)
   - Real estate: Units delivered × Transaction price per unit (2025)
   - Financial services: Interest earning assets × NIM (2020)
4. **Historical CAGR**: Calculate by segment (minimum 5 years)
5. **Bottom-up forecast**: Project each driver separately, then multiply to get segment revenue
6. **Cross-check with top-down**: Ensure bottom-up total aligns with market growth estimates

### CFA Innovation: Dual Revenue Build (2024 Cargojet)
Build BOTH top-down (segment CAGR) AND bottom-up (capacity × utilization × rate), verify convergence.

**NEXT STEPS:**
- [ ] **UPDATE: Add dual-build cross-check to Revenue Build module**
- [ ] **NEW: Sector-specific revenue driver templates** (airlines, industrials, real estate, financial services)

---

## 2. DuPont Analysis — HIGH VALUE, MISSING FROM TOOLKIT
**Status:** NOT IN IB Toolkit

### Step-by-Step (3-Component)
```
ROE = Net Profit Margin × Asset Turnover × Equity Multiplier
ROE = (NI/Revenue) × (Revenue/Assets) × (Assets/Equity)
```

### Step-by-Step (5-Component — Extended)
```
ROE = Tax Burden × Interest Burden × EBIT Margin × Asset Turnover × Leverage
ROE = (NI/EBT) × (EBT/EBIT) × (EBIT/Revenue) × (Revenue/Assets) × (Assets/Equity)
```

### When to Use
- Comparing profitability drivers across peers (2018 Lausanne: VAT vs 3 competitors)
- Tracking ROE trend over time and identifying WHICH component drives changes (2019 D&L: 7-year)
- Diagnosing declining returns (2023 Qantas: margin compression identified)

### Presentation: DuPont Decomposition Visual
```
┌─────────────────────────────────────────────────┐
│ DuPont ROE Decomposition (FY19A-FY28E)          │
├─────────────────────────────────────────────────┤
│                                                  │
│  Net Profit     Asset          Equity            │
│  Margin    ×    Turnover  ×    Multiplier = ROE  │
│  ┌────┐         ┌────┐         ┌────┐    ┌────┐ │
│  │17% │    ×    │0.6x│    ×    │3.1x│ =  │32% │ │
│  └─┬──┘         └─┬──┘         └─┬──┘    └────┘ │
│    │               │               │              │
│  Declining       Stable         Declining         │
│  (cost pressure) (asset-light)  (deleveraging)    │
└─────────────────────────────────────────────────┘
```

**NEXT STEPS:**
- [ ] **NEW: DuPont Analysis Excel module** (3-component and 5-component, 5-year historical + forecast)
- [ ] **NEW: DuPont Decomposition slide template** (visual with arrows showing which lever moves)

---

## 3. ROIC Analysis — CRITICAL FOR VALUE CREATION
**Status:** NOT IN IB Toolkit (separately from returns analysis)

### Step-by-Step
```
ROIC = NOPAT / Invested Capital
NOPAT = EBIT × (1 - Tax Rate)
Invested Capital = NWC + Net PP&E + Goodwill + Intangibles
```

### Key Insight (2024 Cargojet)
When invested capital is stable (PP&E base no longer growing), margin improvement flows DIRECTLY to ROIC. This is the operating leverage inflection point.

### When to Use
- Always compare ROIC to WACC (ROIC > WACC = value creation)
- Track ROIC decomposition: NOPAT margin × Invested Capital turnover
- Identify inflection points where capex cycle matures

**NEXT STEPS:**
- [ ] **NEW: ROIC decomposition module** (with invested capital turnover analysis)
- [ ] **NEW: ROIC vs WACC value creation spread chart**

---

## 4. Earnings Quality & Fraud Detection Scores
**Status:** NOT IN IB Toolkit

### Beneish M-Score (2017 Copa, 2022 MSI)
```
M-Score = -4.84 + 0.920(DSRI) + 0.528(GMI) + 0.404(AQI) + 0.892(SGI)
          + 0.115(DEPI) - 0.172(SGAI) + 4.679(TATA) - 0.327(LVGI)

If M-Score > -1.78 → probable earnings manipulation
```
**8 variables, 5 years of data minimum**

### Altman Z-Score (2017 Copa, 2022 MSI)
```
Z = 1.2(WC/TA) + 1.4(RE/TA) + 3.3(EBIT/TA) + 0.6(MV Equity/BV Debt) + 1.0(Sales/TA)

Z > 2.99 → Safe zone
1.81 < Z < 2.99 → Grey zone
Z < 1.81 → Distress zone
```

### Piotroski F-Score (2022 MSI)
9-point scoring system: profitability (4 points), leverage/liquidity (3 points), operating efficiency (2 points).
Score 7-9 = strong, 0-3 = weak.

**NEXT STEPS:**
- [ ] **NEW: Earnings Quality module** (M-Score, Z-Score, F-Score calculated automatically)
- [ ] **NEW: Earnings Quality dashboard slide** (traffic light for each score)

---

## 5. Sector-Specific KPI Analysis

### Airlines (2017 Copa, 2023 Qantas)
| KPI | Definition | How to Use |
|-----|------------|------------|
| RPM/RPK | Revenue Passenger Miles/Km | Volume metric |
| ASM/ASK | Available Seat Miles/Km | Capacity metric |
| Load Factor | RPM/ASM | Utilization |
| Yield | Passenger Rev / RPM | Pricing power |
| RASK | Revenue / ASK | Unit revenue |
| CASK | Cost / ASK | Unit cost (ex-fuel and total) |
| BELF | Break-even Load Factor | Minimum to cover costs |
| EBITDAR | EBITDA + Aircraft Rent | Normalizes lease vs own |

### Industrials / Manufacturing (2018 VAT, 2021 Vestas)
| KPI | Definition | How to Use |
|-----|------------|------------|
| Backlog | Unfilled orders | Revenue visibility |
| Book-to-bill | Orders / Revenue | Growth indicator |
| EBITDA/unit | Efficiency metric (e.g., EBITDA/GW) | Peer comparison |
| Capacity utilization | Output / Max capacity | Pricing power indicator |
| R&D intensity | R&D / Revenue | Innovation investment |
| Operating cycle | DIO + DSO - DPO | Cash efficiency |

### Financial Services (2020 CBA)
| KPI | Definition | How to Use |
|-----|------------|------------|
| NIM | Net Interest Margin | Core profitability |
| CTI | Cost-to-Income Ratio | Operating efficiency |
| CET1 | Common Equity Tier 1 | Capital adequacy |
| Impairment Rate | Loan losses / Gross loans | Credit quality |
| Lending Beta | Pass-through of rate changes | Rate sensitivity |
| Jaws | Revenue growth - Cost growth | Operating leverage |

### Air Freight / Logistics (2024 Cargojet)
| KPI | Definition | How to Use |
|-----|------------|------------|
| Revenue/lb capacity | Revenue / Total capacity | Asset productivity |
| Utilization rate | Actual volume / Total capacity | Demand indicator |
| Cost per block hour | Direct costs / Flight hours | Operating efficiency |
| Fixed cost ratio | Fixed / Total costs | Operating leverage |

**NEXT STEPS:**
- [ ] **NEW: Sector KPI reference sheets** for Airlines, Industrials, Financial Services, Logistics, Real Estate
- [ ] **UPDATE: Sector Playbooks** to include CFA-derived KPIs alongside PE-focused metrics

---

## 6. Capital Allocation / OCF Framework (2022 MSI)
**Status:** NOT IN IB Toolkit

### Framework
```
Operating Cash Flow = 100%
├── CapEx: 20% (maintenance + growth)
├── Acquisitions: 37.5% (M&A pipeline)
├── Dividends: 30% (shareholder return)
└── Share Repurchases: 12.5% (residual)
```

### When to Use
Track how management allocates cash over time. Compare disclosed priorities to actual spending.

**NEXT STEPS:**
- [ ] **NEW: Capital allocation analysis slide** (pie chart of OCF deployment + historical tracking)
- [ ] **NEW: Capital allocation analysis in IB Toolkit**

---

## 7. NIM Decomposition (Financial Services Specific)
**Status:** NOT IN IB Toolkit (no Financial Services sector)

### Step-by-Step (2020 CBA)
1. Build Average Interest Earning Assets (AIEA) by product
2. Model lending beta (rate pass-through to borrowers: typically 0.6-0.7x)
3. Model deposit pricing floor (% of deposits already at zero)
4. Quantify replicating portfolio drag (hedging book roll-off at lower rates)
5. Layer in competitive dynamics (non-bank lender pressure)
6. Output: NIM forecast with waterfall showing each driver

**NEXT STEPS:**
- [ ] **NEW: Financial Services sector playbook** (NIM decomposition, CTI analysis, capital adequacy)

---

## 8. Primary Research Integration
**Status:** NOT formalized in IB Toolkit

### What CFA Winners Do
| Year | Research Type | Scale |
|------|--------------|-------|
| 2016 | Price basket across 6 retailers, store visits | Custom fieldwork |
| 2017 | Consumer survey (385 respondents, 6 regions) | Statistical survey |
| 2019 | Client relationship mapping | Competitive intelligence |
| 2022 | First responder survey | Customer insights |
| 2024 | 15 expert interviews + consumer survey (N=97) | Expert network |
| 2025 | 30+ expert interviews + site visit | Institutional-grade |

**Key Insight:** Top teams differentiate through PRIMARY research that sell-side cannot replicate.

**NEXT STEPS:**
- [ ] **NEW: Primary Research Framework** (template for planning expert interviews, surveys, site visits)
- [ ] **NEW: Expert interview quote integration guide** (how to weave primary research into analysis)

---

## Priority Development Queue

| Priority | New Module | Impact | Effort |
|----------|-----------|--------|--------|
| 1 | DuPont Analysis (Excel + Slide) | High — used 70% | Low |
| 2 | Earnings Quality Scores (M-Score, Z-Score, F-Score) | Medium — differentiation | Low |
| 3 | ROIC Decomposition | High — value creation | Low |
| 4 | Sector KPI Reference Sheets | High — reusable | Medium |
| 5 | Capital Allocation Framework | Medium | Low |
| 6 | NIM Decomposition (Financial Services) | Sector-specific | Medium |
| 7 | Primary Research Framework | High — differentiation | Low |
