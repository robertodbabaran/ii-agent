# Finance Professional Skill — Development Roadmap
## Prioritized Build Plan (14 Selected Items)

*Curated from CFA Research Challenge Analysis (2016-2025). Ordered by implementation sprint.*

---

## Sprint 1: Quick Wins — Pure Slide Templates
*Standalone slide templates, no Excel dependency. Highest ROI per effort.*

| Rank | Old # | Item | Category | Function(s) | Status |
|------|-------|------|----------|-------------|--------|
| 1 | 4 | "What Must Be True" slide | Wow Factor | `add_wmbt_slide()` | **DONE** |
| 2 | 8 | Risk-Mitigant pairing table template | Risk | `add_risk_mitigant_slide()` | **DONE** |
| 3 | 9 | Valuation blending framework | Valuation | `add_valuation_blending_slide()` | **DONE** |
| 4 | 32 | Value Chain slide template | Industry | `add_value_chain_slide()` | **DONE** |

---

## Sprint 2: Core Visual Slides
*Chart-based slides with moderate complexity. Visual differentiation for IC decks.*

| Rank | Old # | Item | Category | Function(s) | Status |
|------|-------|------|----------|-------------|--------|
| 5 | 2 | Risk Matrix (bubble chart) | Risk | `add_risk_matrix_slide()` | **DONE** |
| 6 | 34 | TAM Funnel (market sizing cascade) | Industry | `add_tam_funnel_slide()` | **DONE** |
| 7 | 31 | PESTEL analysis slide | Industry | `add_pestel_slide()` | **DONE** |
| 8 | 59 | Porter's radar chart (quantified pentagon) | Industry | `add_porters_radar_slide()` | **DONE** |

---

## Sprint 3: Excel + Slide Pairs
*Items requiring both Excel computation and slide visualization.*

| Rank | Old # | Item | Category | Function(s) | Status |
|------|-------|------|----------|-------------|--------|
| 9 | 5 | Tornado sensitivity chart | Risk | `add_tornado_sensitivity()` + `add_tornado_sensitivity_slide()` | **DONE** |
| 10 | 3 | Football Field + Excel helper | Valuation | `add_football_field()` + `add_football_field_slide()` | **DONE** |
| 11 | 13 | SOTP valuation waterfall + Excel | Valuation | `add_sotp_valuation()` + `add_sotp_waterfall_slide()` | **DONE** |

---

## Sprint 4: Valuation Engine Enhancements
*Deep Excel module work enhancing existing code. Interdependent builds.*

| Rank | Old # | Item | Category | Function(s) | Status |
|------|-------|------|----------|-------------|--------|
| 12 | 27 | Enhanced WACC (3-method Kd, multi-method beta) | Valuation | `add_enhanced_wacc()` | **DONE** |
| 13 | 23 | Geographic-weighted terminal growth | Valuation | `add_geographic_terminal_growth()` | **DONE** |
| 14 | 22 | Multi-stage DCF (3-stage, linear decline) | Valuation | `add_multi_stage_dcf()` | **DONE** |

---

## Prioritization Rationale

**Sprint 1 first** because these are standalone slide templates with ASCII specs already designed. Immediately usable in any IC deck or case study with zero Excel dependency.

**Sprint 2 second** because chart-based slides add visual differentiation. Risk Matrix (100% CFA frequency) is the single most common template across all winners.

**Sprint 3 third** because Excel + Slide pairs require dual implementation (computation + visualization). Football Field (70% frequency) is the most impactful synthesis slide.

**Sprint 4 last** because valuation engine enhancements are deep, interdependent module upgrades. Enhanced WACC feeds into DCF; Geographic terminal growth feeds into multi-stage DCF.

---

## Relevant Prompts (P-CFA series)

| Prompt ID | Name | Roadmap Item |
|-----------|------|--------------|
| P-CFA-03 | "Create risk matrix with quantified impacts" | Sprint 2, Rank 5 |
| P-CFA-06 | "Create football field chart" | Sprint 3, Rank 10 |
| P-CFA-08 | "Create SOTP valuation" | Sprint 3, Rank 11 |

---

## File Location

**This roadmap is saved at:** `src/ii_skills/finance_professional/ROADMAP.md`

**Related playbooks (detail for each item):**
- `VALUATION_PLAYBOOK.md` — Items 3, 9, 13, 22, 23, 27
- `INDUSTRY_PLAYBOOK.md` — Items 31, 32, 34, 59
- `RISK_PLAYBOOK.md` — Items 2, 5, 8
- `WOW_FACTOR_PLAYBOOK.md` — Item 4

---

*Version: 2.0 | Updated: February 2026*
*Source: 10 years of CFA Research Challenge Global Winners (2016-2025)*
*Curated to 14 items from original 60 per user prioritization.*
