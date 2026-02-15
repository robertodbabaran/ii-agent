# Risk & Sensitivity Playbook
## Frameworks from 10 Years of CFA Research Challenge Winners (2016-2025)

---

## Risk Framework Frequency

| Framework | Years Used | Status in IB Toolkit |
|-----------|-----------|---------------------|
| Risk Matrix (Probability × Impact) | 10/10 | **BUILT** — `add_risk_matrix_slide()` (Sprint 2) |
| WACC × TGR Sensitivity Grid | 10/10 | EXISTS (Sensitivity module) |
| Scenario Analysis (Bull/Base/Bear) | 10/10 | EXISTS (Scenario module) |
| Monte Carlo Simulation | 8/10 | **NEW** |
| Tornado / Spider Sensitivity | 4/10 | **BUILT** — `add_tornado_sensitivity()` (Excel) + `add_tornado_sensitivity_slide()` (Sprint 3) |
| Risk-Mitigant Pairing Table | 8/10 | **BUILT** — `add_risk_mitigant_slide()` (Sprint 1) |
| Value-at-Risk (VaR) | 1/10 | **NEW** |
| Yield Breakeven Analysis | 1/10 | **NEW** |
| Quantified Risk Impact ($ per bp) | 3/10 | **NEW** |

---

## 1. Risk Matrix (Probability × Impact) — MUST HAVE
**Status:** BUILT (Sprint 2) — `add_risk_matrix_slide()` (3×3 heat-mapped bubble chart with category legend)
**Frequency:** 100% of winners use this

### Grid Options
- **2×2** (2016): Simple Likelihood vs Impact with 6 risks
- **3×3** (2017, 2023, 2025): Low/Med/High on each axis, color-coded
- **5×5** (2020): Full granularity, used for financial services

### Step-by-Step
1. Identify 6-12 risks across categories: Market, Operational, Political/Regulatory, Financial
2. Assess probability (Low/Medium/High) and impact (Low/Medium/High)
3. Plot on grid with color coding by category
4. For each risk, provide:
   - Description (1 sentence)
   - Valuation impact (quantified: "$X per share" or "X% downside")
   - Mitigant (specific, not generic)
5. Present as scatter plot or bubble chart

### CFA Best Practice: Quantified Risk Impact (2021 Vestas, 2024 Cargojet, 2025 Kozminski)
Every risk should have a QUANTIFIED valuation impact:
- "15% drop in deliveries → target drops to DKK 1,225 (-17%)" (2021)
- "10% decrease in government contracts → 8.52% downside" (2022)
- "Delta -2.5% revenues → -22% value impact" (2025)

### Slide Template
```
┌─────────────────────────────────────────────────┐
│ RISK ASSESSMENT                                  │
├─────────────────────────────────────────────────┤
│                                                  │
│  Impact ↑                                        │
│  HIGH  │  [P2]     [M1]●    [M2]●               │
│        │                                         │
│  MED   │  [O1]     [O2]●                        │
│        │                                         │
│  LOW   │  [P1]                                   │
│        └────────────────────────────→            │
│          LOW      MED       HIGH   Probability   │
│                                                  │
│  ● Market  ■ Operational  ▲ Political            │
│                                                  │
│  Risk Details:                                   │
│  M1: [Risk] → [X]% valuation impact             │
│      Mitigant: [Specific action]                 │
│  M2: [Risk] → [X]% valuation impact             │
│      Mitigant: [Specific action]                 │
└─────────────────────────────────────────────────┘
```

**BUILT:**
- [x] **DONE: Risk Matrix slide template** → `add_risk_matrix_slide()` (bubble chart with quantified impacts, Sprint 2)
- [x] **DONE: Risk-Mitigant pairing table template** → `add_risk_mitigant_slide()` (quantified impact footer, Sprint 1)

**REMAINING:**
- [ ] **NEW: Risk Register Excel module** (auto-calculates expected value impact)

---

## 2. Sensitivity Analysis — ENHANCED APPROACHES

### Standard: WACC × Terminal Growth Rate (5×5 Grid) — EXISTS
Every winner does this. Standard format:
- 5 WACC values (±1% from base in 0.5% steps)
- 5 Terminal Growth values (±1% from base in 0.5% steps)
- Color-coded: green (BUY), yellow (HOLD), red (SELL)

### Advanced: Multi-Variable Two-Way Tables
Beyond WACC × TGR, top teams add additional sensitivity grids:
| Grid | Year | Purpose |
|------|------|---------|
| WACC × MW Delivered | 2021 | Volume sensitivity |
| WACC × Average Selling Price | 2021 | Pricing sensitivity |
| WACC × COGS/Revenue | 2021 | Margin sensitivity |
| Brent Price × Refining Margin | 2023 | Input cost sensitivity |
| Revenue Change × Cost Change | 2025 | Combined operating sensitivity |
| Market Share × Marketing Cost | 2023 | Competitive sensitivity |
| Payout Ratio × Cost of Equity | 2025 | DDM sensitivity |

### Single-Variable Sensitivity (Tornado Chart)
Test each variable independently ±10-15%:
1. Rank variables by price impact (largest first)
2. Present as horizontal tornado chart
3. Identify top 2-3 risk factors

### CFA Best Practice (2022 MSI)
"Each assumption stressed ±10% independently, then all simultaneously via Monte Carlo"

**BUILT (Sprint 3):**
- [x] **DONE: Tornado chart slide template** → `add_tornado_sensitivity_slide()` (horizontal bars ranked by impact)
- [x] **DONE: Single-variable sensitivity table** → `add_tornado_sensitivity()` (Excel, auto-sorted by swing)

**REMAINING:**
- [ ] **UPDATE: Add multi-variable sensitivity grids** beyond WACC × TGR

---

## 3. Scenario Analysis — THESIS-LINKED
**Status:** EXISTS in IB Toolkit (Scenario Analysis module)

### CFA Best Practice: Link Scenarios to Investment Thesis
Don't just change numbers — tie each scenario to thesis pillar:

| Scenario | Thesis 1 (Tech Moat) | Thesis 2 (Offshore) | Thesis 3 (Green) | Target |
|----------|---------------------|---------------------|-------------------|--------|
| **Bear** | Moat fails, share 20% | Synergies fail, share 20% | Wind can't compete | DKK 1,087 |
| **Base** | Moat maintained, share 31% | Partial synergies, share 27% | Wind + solar complement | DKK 1,480 |
| **Bull** | Moat widens, share 40% | Full offshore success, share 40% | Wind preferred | DKK 1,848 |

### Cascading Impact Chain (2016 Waterloo)
```
Oil price drops → Alberta unemployment → Reduced retail foot traffic →
CTC segment revenue decline → Lower EBITDA → Target price impact
```

**NEXT STEPS:**
- [ ] **UPDATE: Scenario Analysis module** to include thesis-pillar mapping
- [ ] **NEW: Cascading impact chain diagram slide template**

---

## 4. Natural Disaster / Facility Risk Mapping (2018 Lausanne)
**Status:** NOT IN IB Toolkit

### When to Use
Companies with concentrated production facilities.

### Step-by-Step
1. Map all production facilities
2. Assess seismic risk, flood risk, political stability per location
3. Quantify revenue concentration per facility
4. Present as risk heat map overlay on facility map

**NEXT STEPS:**
- [ ] **NEW: Facility risk mapping template** (for industrial/manufacturing cases)

---

## 5. Specific Risk Quantification Formulas

### Per-Basis-Point Impact (2020 CBA)
```
-1bp NIM = ~$75M NPAT impact = ~$0.04 per share
+10bp impairment rate = ~$760M pre-tax = ~$0.30 per share
```

### Operational Threshold Analysis
```
100bps reduction in BELF = +12cps share price (2023 Qantas)
10bp annual fuel efficiency gain = 1.8% share price impact (2023 Qantas)
```

### Yield Breakeven (2017 Copa)
"At what yield does recommendation change from BUY to HOLD? From HOLD to SELL?"

**NEXT STEPS:**
- [ ] **NEW: Risk quantification formula templates** (per-bp and threshold analysis)

---

## Priority Development Queue

| Priority | Item | Impact | Status |
|----------|------|--------|--------|
| 1 | Risk Matrix slide template (quantified) | High — 100% frequency | **DONE** (Sprint 2) |
| 2 | Monte Carlo module (see Valuation Playbook) | High — 80% frequency | TODO |
| 3 | Tornado sensitivity chart | Medium — visual impact | **DONE** (Sprint 3) |
| 4 | Multi-variable sensitivity grids | Medium — enhancement | TODO |
| 5 | Risk-Mitigant pairing table | Medium — best practice | **DONE** (Sprint 1) |
| 6 | Cascading impact chain diagram | Medium — differentiator | TODO |
