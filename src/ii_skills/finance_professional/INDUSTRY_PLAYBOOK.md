# Industry Analysis Playbook
## Frameworks from 10 Years of CFA Research Challenge Winners (2016-2025)

---

## Framework Frequency

| Framework | Years Used | Status in IB Toolkit |
|-----------|-----------|---------------------|
| Porter's Five Forces | 8/10 | **ENHANCED** — `add_porters_radar_slide()` quantified pentagon (Sprint 2) |
| SWOT Analysis | 8/10 | EXISTS (Competitive Analysis slides) |
| Market Share Analysis | 10/10 | Partial |
| TAM/SAM/SOM Sizing | 6/10 | **ENHANCED** — `add_tam_funnel_slide()` cascade slide (Sprint 2) |
| Competitive Scoring Matrix | 5/10 | **NEW** |
| BCG Growth-Share Matrix | 2/10 | **NEW** |
| McKinsey/GE Matrix | 1/10 | **NEW** |
| PESTEL Analysis | 3/10 | **BUILT** — `add_pestel_slide()` 6-panel slide (Sprint 2) |
| Value Chain Analysis | 4/10 | **BUILT** — `add_value_chain_slide()` horizontal flow (Sprint 1) |
| Industry Life Cycle | 2/10 | **NEW** |
| HHI (Market Concentration) | 1/10 | **NEW** |
| LCOE / Unit Economics | 2/10 | Sector-specific |
| Patent / IP Analysis | 1/10 | **NEW** |

---

## 1. Porter's Five Forces — ENHANCED VERSION

### Standard Deployment (do every time)
Score each force 1-5 and present as radar/pentagon chart (2017 Copa style).

### CFA Enhancement: Quantified Forces
Don't just say "moderate" — quantify:
- **Supplier power**: "No single supplier >5% of COGS" (2021 Vestas)
- **Buyer power**: "Mortgage brokers originate >60% of home loans" (2020 CBA)
- **New entrants**: "Customers need ~25-30% of national volume to break even on insourcing ($413-535M)" (2024 Cargojet)
- **Substitutes**: "Income elasticity of airfare demand: Australia 1.0 vs Malaysia 5.0" (2023 Qantas)

**BUILT (Sprint 2):**
- [x] **DONE: Porter's radar chart slide** → `add_porters_radar_slide()` (quantified pentagon with star ratings, trigonometric layout)
- [x] **DONE: Porter's Five Forces** enhanced — now uses radar chart instead of text-only framework slide

**REMAINING:**
- [ ] **UPDATE: Add quantitative evidence prompts** to each force

---

## 2. Competitive Scoring Matrix — HIGH VALUE
**Status:** NOT IN IB Toolkit

### Step-by-Step (2017 Copa)
1. Identify 6-8 Key Success Factors for the industry
2. Assign weights to each factor (total = 100%)
3. Score company and each competitor (1-5 scale)
4. Calculate weighted score → competitive position ranking

### Variants
- **Technology Scoring Matrix** (2021 Vestas): 5 tech categories, 3-point scale, company vs 2 peers
- **Customer Survey Scoring** (2022 MSI): Survey-based competitive perception
- **Full weighting table** (2017 Copa): 8 KSFs with explicit weights

**NEXT STEPS:**
- [ ] **NEW: Competitive scoring matrix slide template** (weighted comparison table)
- [ ] **NEW: Competitive scoring matrix Excel module**

---

## 3. Market Sizing — FROM TAM TO ADDRESSABLE

### Framework
```
Total Addressable Market (TAM)
├── Serviceable Addressable Market (SAM)
│   └── Serviceable Obtainable Market (SOM)
│       └── Company Current Share
│           └── Forecast Share Growth
```

### CFA Best Practices
- **Sub-market cascading** (2018 Lausanne): End-market → Equipment → Vacuum → Valves (narrowing funnel)
- **Geographic supply/demand gap** (2018 Lausanne): China = 14% supply vs 62% demand → opportunity
- **Housing gap analysis** (2025 Kozminski): 1.5-2M unit deficit, overcrowding 41% (highest EU)
- **E-commerce penetration comparison** (2024 Cargojet): Canada 12% vs China 47% → growth runway
- **TAM doubling tracking** (2022 MSI): TAM doubled from 2015-2021 through adjacency acquisitions

**BUILT (Sprint 2):**
- [x] **DONE: Market sizing cascade slide** → `add_tam_funnel_slide()` (descending-width bars, navy→green gradient)

**REMAINING:**
- [ ] **NEW: Geographic supply/demand gap analysis template**

---

## 4. Barriers to Entry / Moat Analysis

### CFA Moat Frameworks

| Moat Type | Example | Year |
|-----------|---------|------|
| Switching costs (sunk infrastructure) | MSI LMR networks | 2022 |
| Regulatory protection (cabotage laws) | Cargojet domestic monopoly | 2024 |
| Scale advantage (production efficiency) | Vestas 2x peer EBITDA/GW | 2021 |
| Network effects (dealer model) | Canadian Tire associate dealers | 2016 |
| Mission-critical c-parts | VAT valves (2-3% of cost, zero-failure) | 2018 |
| Customer alignment (warrants/equity) | Amazon/DHL warrants in Cargojet | 2024 |
| Technology flywheel | R&D → tech → premium pricing → more R&D | 2021 |
| Brand + convenience | Canadian Tire 15-min drive coverage | 2016 |

### Quantified Moat (Best Practice)
Don't just describe the moat — QUANTIFY it:
- "Insourcing break-even: $413M / ~25% of national volume" (2024)
- "84% of Tocumen airport traffic" (2017)
- "90%+ domestic market share" (2024)

**NEXT STEPS:**
- [ ] **NEW: Economic moat analysis slide** (moat type + quantification)
- [ ] **NEW: Insourcing break-even calculation framework**

---

## 5. PESTEL Analysis
**Status:** BUILT (Sprint 2) — `add_pestel_slide()` (3×2 grid with colored circle badges)

### Step-by-Step
| Factor | What to Analyze | CFA Example |
|--------|----------------|-------------|
| **P**olitical | Government spending, policy | ARPA funding for public safety (2022) |
| **E**conomic | GDP, rates, FX, inflation | RBA rate cuts and NIM impact (2020) |
| **S**ocial | Demographics, consumption | Ukrainian migration impact (2025) |
| **T**echnological | Disruption, R&D trends | LTE vs LMR transition (2022) |
| **E**nvironmental | Emissions, regulation | Renewable transition, LCOE (2021) |
| **L**egal | Regulation, compliance | Spatial Planning Reform (2025) |

**BUILT (Sprint 2):**
- [x] **DONE: PESTEL analysis slide template** → `add_pestel_slide()` (6-panel with P/E/S/T/E/L badges)

**REMAINING:**
- [ ] **NEW: PESTEL checklist for case intake** (structured questionnaire)

---

## 6. Value Chain Analysis

### CFA Examples
- **Horizontal flow** (2019 D&L): Back-end B2B → Front-end Consumer
- **Competitive position** (2018 Lausanne): Fragmented suppliers → VAT → 3 customers
- **Middle-mile focus** (2024 Cargojet): Manufacturing → First Mile → Middle Mile → Last Mile → Customer
- **Vertical integration** (2021 Vestas): Project Development → Manufacturing → Transport → Construction → Service

**BUILT (Sprint 1):**
- [x] **DONE: Value chain slide template** → `add_value_chain_slide()` (horizontal boxes + arrows + highlighted company position, replaces old bullet list in `add_industry_analysis()`)

---

## 7. HHI (Market Concentration)
**Status:** NOT IN IB Toolkit

### When to Use
Quantify market concentration: HHI = sum of squared market shares.
- HHI < 1,500: Competitive market
- 1,500-2,500: Moderately concentrated
- HHI > 2,500: Highly concentrated

### CFA Example (2018 Lausanne)
VAT Group's vacuum valve market: HHI 1,850 → moderately concentrated, dominant position.

**NEXT STEPS:**
- [ ] **NEW: HHI calculation in competitive analysis module**

---

## 8. Regulatory Impact Analysis

### CFA Best Practices
- **Pending legislation modeling** (2019 D&L): TRABAHO Bill — modeled House vs Senate versions
- **Tax reform scenarios** (2019): Three scenarios for different legislative outcomes
- **Regulatory compliance cost** (2025): SG&A +10bps/year for ESG compliance
- **Royal Commission impact** (2020): 76 recommendations mapped to compliance cost and revenue impact

**NEXT STEPS:**
- [ ] **NEW: Regulatory impact assessment template** (legislation → financial impact mapping)

---

## Priority Development Queue

| Priority | Item | Impact | Status |
|----------|------|--------|--------|
| 1 | Competitive Scoring Matrix (slide + Excel) | High — reusable across all cases | TODO |
| 2 | PESTEL Analysis template | High — gap in toolkit | **DONE** (Sprint 2) |
| 3 | Value Chain slide template | Medium | **DONE** (Sprint 1) |
| 4 | Economic Moat quantification framework | High — differentiator | TODO |
| 5 | Market sizing cascade slide | Medium | **DONE** (Sprint 2) |
| 6 | HHI calculation module | Low — niche | TODO |
