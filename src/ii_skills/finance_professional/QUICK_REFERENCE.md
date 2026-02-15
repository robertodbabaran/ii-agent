# Finance Professional — Shared Skill Quick Reference

*Shared resources for IB Toolkit, IR Toolkit, and all finance case work.*

---

## What's Here

| Resource | File | Purpose |
|----------|------|---------|
| **Slide Library** | `SLIDE_LIBRARY.md` | 39+ ASCII slide templates (IC decks + equity research) |
| **Slide Design Guide** | `SLIDE_DESIGN_GUIDE.md` | 27 professional templates with python-pptx code |
| **Excel Conventions** | `EXCEL_CONVENTIONS.md` | Color coding, number formats, model architecture |
| **Case Study Standards** | `CASE_STUDY_STANDARDS.md` | 10 universal standards for all case work |
| **Analysis Playbook** | `ANALYSIS_PLAYBOOK.md` | 8 analysis types (DuPont, ROIC, Earnings Quality, etc.) |
| **Valuation Playbook** | `VALUATION_PLAYBOOK.md` | 10 valuation methods with WACC enhancements |
| **Industry Playbook** | `INDUSTRY_PLAYBOOK.md` | 13 frameworks (Porter's, PESTEL, HHI, Value Chain) |
| **Risk Playbook** | `RISK_PLAYBOOK.md` | Risk matrix, Monte Carlo, tornado, scenario analysis |
| **ESG Playbook** | `ESG_PLAYBOOK.md` | Governance scorecard, ESG-in-valuation, SDG alignment |
| **Wow Factor Playbook** | `WOW_FACTOR_PLAYBOOK.md` | 20 differentiation techniques in 4 tiers |
| **Development Roadmap** | `ROADMAP.md` | 14 prioritized items (all DONE) |
| **Reference Templates** | `templates/reference_outputs/` | 30 PPTX/XLSX institutional templates |

---

## How IB Toolkit Uses This

| IB Task | Start Here |
|---------|------------|
| Build slide deck | `SLIDE_LIBRARY.md` → pick template → `SLIDE_DESIGN_GUIDE.md` for design specs |
| Create Excel model | `EXCEL_CONVENTIONS.md` → color coding, tab architecture, formula rules |
| Start a case | `CASE_STUDY_STANDARDS.md` → folder structure, audit protocol |
| Valuation work | `VALUATION_PLAYBOOK.md` → method selection, WACC build, blending |
| Risk analysis | `RISK_PLAYBOOK.md` → risk matrix, sensitivity, Monte Carlo |
| Industry analysis | `INDUSTRY_PLAYBOOK.md` → Porter's, competitive scoring, PESTEL |
| ESG section | `ESG_PLAYBOOK.md` → governance scorecard, ESG-in-valuation |
| Differentiate work | `WOW_FACTOR_PLAYBOOK.md` → primary research, reverse DCF, proprietary metrics |
| Reference templates | `templates/reference_outputs/` → Institutional LBO/Deck templates |

## How IR Toolkit Uses This

| IR Task | Start Here |
|---------|------------|
| Build LP presentation | `SLIDE_LIBRARY.md` + `SLIDE_DESIGN_GUIDE.md` for layout |
| Create fund workbook | `EXCEL_CONVENTIONS.md` → GS formatting, dynamic linking |
| Start a case | `CASE_STUDY_STANDARDS.md` → Standards 5,6,9,10 minimum |
| ESG reporting | `ESG_PLAYBOOK.md` → environmental metrics, governance scorecard |
| Risk register | `RISK_PLAYBOOK.md` → risk matrix with quantified impacts |

---

## Cross-Reference Map

| Old Location | New Location |
|-------------|-------------|
| `docs/skills/ib_toolkit/SLIDE_TEMPLATES.md` | `finance_professional/SLIDE_LIBRARY.md` |
| `docs/skills/SLIDE_DESIGN_GUIDE.md` | `finance_professional/SLIDE_DESIGN_GUIDE.md` |
| `docs/skills/ib_toolkit/EXCEL_CONVENTIONS.md` | `finance_professional/EXCEL_CONVENTIONS.md` |
| `docs/skills/CASE_STUDY_STANDARDS.md` | `finance_professional/CASE_STUDY_STANDARDS.md` |
| `src/ii_skills/ib_toolkit/templates/reference_outputs/` | `finance_professional/templates/reference_outputs/` (copy) |

---

## CFA Research Challenge Patterns (Summary)

Extracted from 10 years of Global Winners (2016-2025). Key patterns:

1. **Always blend 2-3 valuation methods** with explicit weights
2. **Always include WACC x TGR sensitivity grid** (5x5)
3. **Always include scenario analysis** (Bull/Base/Bear) linked to thesis pillars
4. **Revenue decomposition is bottom-up** by segment AND geography
5. **Risk matrix with QUANTIFIED valuation impact** for each risk
6. **Primary research** differentiates the top tier
7. **Massive appendix** (median 150 slides) signals depth

### IB Toolkit Gaps (Ranked by CFA Frequency)

| Item | Frequency | Status |
|------|-----------|--------|
| Risk Matrix (quantified) | 100% | **BUILT** — `add_risk_matrix()` (Sprint 2) |
| Monte Carlo Simulation | 80% | **TODO** |
| Football Field Chart | 70% | **BUILT** — `add_football_field()` (Sprint 3) |
| DuPont Analysis | 70% | **TODO** |
| DDM Template | 60% | **TODO** |
| SOTP Valuation | 40% | **BUILT** — `add_sotp_valuation()` (Sprint 3) |
| Reverse DCF | 40% | **TODO** |
| ESG Scorecard | Post-2020: 100% | **TODO** |

### CFA Modules Built (14 items, all DONE)

| Sprint | Modules | Type |
|--------|---------|------|
| **Sprint 1** | What Must Be True, Risk-Mitigant, Valuation Blend, Value Chain | Slides |
| **Sprint 2** | Risk Matrix, TAM Funnel, PESTEL, Porter's Radar | Slides |
| **Sprint 3** | Tornado Sensitivity, Football Field, SOTP Waterfall | Excel + Slides |
| **Sprint 4** | Enhanced WACC, Geographic Terminal Growth, Multi-Stage DCF | Excel |

Full details: `ROADMAP.md`

---

*Location: `src/ii_skills/finance_professional/`*
*Referenced by: IB Toolkit, IR Toolkit, Career Toolkit*
*Version: 1.0 | Created: February 2026*
