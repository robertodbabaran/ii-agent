# Universal Case Study Standards

*Cross-cutting standards for all case work: IB Toolkit, IR Toolkit, Interview Prep, and future skills.*

*Source: Originally developed during the Brookfield Triton case study (Feb 2026), then generalized for universal application.*

---

## When to Use

| Case Complexity | Apply Standards | Notes |
|-----------------|----------------|-------|
| **Full case study** (4+ hours) | All 10 standards | LP Transaction Summary, Fundraising Deck, IC Package |
| **Standard deliverable** (2-4 hours) | Standards 1, 4, 5, 6, 8, 9, 10 | LP Update, DDQ, Asset Deep Dive |
| **Quick turnaround** (<2 hours) | Standards 5, 6, 9, 10 | Screening, ad-hoc analysis, rapid response |
| **Interview case study** | All 10 standards + Q&A depth | Case study + presentation prep |

---

## Standard 1: Source Material Manager

*IR Toolkit: P31 | IB Toolkit: apply to any case with 3+ source documents*

Organize source materials into a numbered markdown library with full citation tracking.

### Setup

```
<case_folder>/resources/
├── 00_SOURCE_INDEX.md     # Master citation catalog
├── 01_<Source_Name>.md    # First source document
├── 02_<Source_Name>.md    # Second source document
└── ...
```

### Per-Source File Template

```markdown
# [Source Title]

**URL:** [source URL]
**Publisher:** [publisher name]
**Date:** [publication date]
**Informs:** [what this source contributes to the case]

## Key Data Extracted

[Tables, metrics, quotes — formatted for easy reference]
```

### Source Index (`00_SOURCE_INDEX.md`)

```markdown
## Source Inventory

| # | Title | Type | Key Data |
|---|-------|------|----------|
| 01 | [title] | [press release / filing / report] | [key numbers] |
| 02 | [title] | [type] | [key numbers] |

## Data Vintage Classification

| Classification | Definition | Example |
|----------------|------------|---------|
| FIXED | Transaction terms (unchangeable) | Enterprise Value ($X.XB) |
| AT ENTRY | Historical baseline at deal time | FY20XX Revenue ($X.XM) |
| ACTUAL | Post-close reported results | FY20XX EBITDA ($XM) |
| DERIVED | Calculated from other data | EV/EBITDA (X.Xx) |
| MARKET | Third-party industry estimate | Market size ($XB) |
| ILLUSTRATIVE | Scenario/sensitivity | Stress-case revenue |

## Quick Audit Checklist

- [ ] Key claim 1 → Source ##
- [ ] Key claim 2 → Source ##
```

---

## Standard 2: Case Study Retrospective

*IR Toolkit: P32 | IB Toolkit: apply after any multi-day case*

After completing a case study, extract reusable patterns, templates, and lessons learned.

### Retrospective Template

Analyze these dimensions:

1. **Workflow efficiency** — what phases worked, what should change
2. **Excel patterns** — formatting, formulas, tab architecture worth reusing
3. **Slide patterns** — layout helpers, visual elements, speaker notes approach
4. **Script patterns** — reusable Python functions and helpers
5. **Audit patterns** — what the consistency audit caught
6. **Q&A preparation** — which questions proved most valuable
7. **Data management** — how source materials were organized
8. **Versioning discipline** — how file versions were managed

### Output

- `memory/<case_name>_retrospective.md` — full retrospective
- Updated `memory/case_study_workflow_patterns.md` — new patterns added
- New memory entries for lessons learned
- Recommendations for toolkit improvements

---

## Standard 3: Consistency Audit Generator

*IR Toolkit: P33 | IB Toolkit: apply to any case with 2+ deliverables*

Cross-deliverable consistency verification. Every number that appears in more than one file must match exactly.

### Audit Matrix Template

| Data Point | File 1 | File 2 | File 3 | ... | Match? |
|------------|--------|--------|--------|-----|--------|
| [metric] | [value] | [value] | [value] | | YES / NOTE / UPDATE |

### Check Categories

1. **Transaction metrics** — EV, equity, premium, timing
2. **Operational metrics** — fleet/capacity, utilization, contracts
3. **Financial metrics** — revenue, EBITDA, margins, leverage
4. **Valuation metrics** — entry multiple, comps
5. **Industry metrics** — market size, growth, share
6. **Sponsor metrics** — AUM, track record, distributions

### Domain-Specific Quality Checks

Append sector-appropriate checks based on the case:

| Sector | Additional Checks |
|--------|-------------------|
| **Infrastructure** | Contracted/merchant split, inflation linkage, DSCR, availability, WACL |
| **SaaS / Tech** | ARR, NDR, LTV/CAC, Rule of 40, gross margin, churn rate |
| **Healthcare** | Payor mix, reimbursement rates, bed occupancy, EBITDAR coverage |
| **Industrial / Manufacturing** | Capacity utilization, backlog, book-to-bill, maintenance capex % |
| **Business Services** | Revenue per employee, contract renewal rate, customer concentration |
| **Consumer** | Same-store sales, unit economics, brand metrics, channel mix |

### Format & Compliance Checks

- [ ] Formatting compliance (CONFIDENTIAL footer, slide numbers, footnotes)
- [ ] Number formatting consistency ($XB vs $X.XB, X% vs X.X%)
- [ ] Terminology consistency across all files

### Flags

- **YES** — values match exactly
- **NOTE** — acceptable variation (rounding, different precision)
- **UPDATE** — needs fixing (mismatch)

---

## Standard 4: Excel-PowerPoint Bridge

*IR Toolkit: P34 | IB Toolkit: apply to any workbook feeding a presentation*

Every workbook that feeds a presentation MUST include a "Powerpoint Outputs" bridge tab. This tab is the permanent bidirectional sync point between Excel data and PowerPoint content.

### Tab Structure

| Columns | Purpose |
|---------|---------|
| A-D | Slide content (Slide Name, Metric/Criterion, Display Value, Commentary) |
| E | Empty separator |
| F-I | Link Proof for Audit (source references) |
| J | Empty separator |
| K-N | Numeric inputs (feed Column C formulas) |

### Audit Color Coding (Columns F-I)

- **Green text** — internal tab reference (same workbook)
- **Red text** — external source (URL/document)
- **Gray text** — N/A (no corresponding input in K-N)

### Sync Protocol

1. **PPT to Excel**: Any PPT text/data change → update "Powerpoint Outputs" tab first
2. **Excel to PPT**: Any model change → update "Powerpoint Outputs" tab → then update PPT
3. **Direct edit**: If user edits the bridge tab → reflect in both PPT and upstream data tabs

### Creation Rule

Create the "Powerpoint Outputs" tab FIRST when starting any workbook that feeds a presentation. All other tabs reference or feed into it.

---

## Standard 4a: Source Index — Canonical Data Repository

*Applies to: EVERY workbook that feeds a presentation (IB, IR, Interview Prep)*

Every `[Data_Backup]_v[N].xlsx` MUST include a **"Source Index"** tab. This tab is the **canonical upstream repository** for ALL metrics that flow into any PowerPoint slide. The "Powerpoint Outputs" bridge tab (Standard 4) pulls from or validates against this index.

### Data Flow Hierarchy

```
Source Index (canonical truth — Standard 4a)
    ↓ feeds
Powerpoint Outputs bridge tab (sync point — Standard 4)
    ↓ feeds
PowerPoint slides
```

### Tab Structure

| Column | Header | Purpose | Font Color |
|--------|--------|---------|------------|
| A | # | Sequential number | Black |
| B | Metric | Name of the data point | Black |
| C | Value | The canonical value | Blue `#0070C0` (hardcoded input) |
| D | Source Link | URL or document reference for auditability | Red `#FF0000` + hyperlink if URL |
| E | Category | Transaction / Operational / Financial / Valuation / Industry / Sponsor / Pro Forma | Black (color-coded fill per category) |
| F | Slide Reference | Which slide(s) use this metric | Black |
| G | Data Vintage | FIXED / AT ENTRY / ACTUAL / DERIVED / MARKET / ILLUSTRATIVE | Black |
| H | Last Verified | Date of last verification (YYYY-MM-DD) | Black |

### Category Color Coding (Column E fill)

| Category | Fill Color | Hex |
|----------|-----------|-----|
| Transaction | Light green | `#E2EFDA` |
| Operational | Light blue | `#DDEBF7` |
| Financial | Light orange | `#FCE4D6` |
| Valuation | Light purple | `#E2D9F3` |
| Industry | Light yellow | `#FFF2CC` |
| Sponsor | Light gray | `#D6DCE4` |
| Pro Forma | Light pink | `#F2DCDB` |

### Rules

1. **Create this tab FIRST** — before any model tabs or the bridge tab
2. **Every metric in any slide** MUST have an entry in the Source Index
3. **Column D (Source Link) is mandatory** — no metric without a traceable source
4. **Update Source Index FIRST** when any number changes, then propagate to bridge tab and slides
5. **Auto-filter enabled** on all columns for quick lookup by category or slide
6. **Freeze panes** below header row for scrolling
7. **GS formatting** — Arial 10pt, no gridlines, navy headers, print area set

### Relationship to Other Standards

- **Standard 4 (Bridge Tab)**: The bridge tab's Columns F-I audit links should reference Source Index row numbers
- **Standard 6 (Key Metrics Table)**: The markdown Key Metrics Table is the human-readable companion; Source Index is the Excel-native canonical copy
- **Standard 3 (Consistency Audit)**: Audit should verify every bridge tab metric traces back to a Source Index row

### Excel Module

Auto-generate with: `gen.add_source_index(metrics=[...])` from `excel_modules.py`

---

## Standard 5: File Organization

*IR Toolkit: Section 11 of BEST_PRACTICES.md | IB Toolkit: apply to all case folders*

### Case Folder Template

```
<case_folder>/
├── !Archive/              # Old versions — NEVER delete files (R1)
├── !Scripts/              # Generated Python scripts (R4)
├── resources/             # Numbered source documents (01_..., 02_...)
│   └── 00_SOURCE_INDEX.md # Master citation catalog
├── charts/                # Generated chart images
├── pro_forma_analysis/    # Standalone analysis workbooks
├── 01_Research_Brief.md
├── 02_Key_Metrics_Table.md  # Source of truth for all numbers
├── 03_Slide_Content.md      # Write BEFORE touching PPTX
├── 04_[Lens]_Critique.md    # Domain-specific quality critique
├── 05_QA_Preparation.md     # 50+ questions with model answers
├── 06_Cheat_Sheet.md        # 1-page rapid reference
├── 07_Consistency_Audit.md  # Cross-deliverable audit
├── 09_Final_Audit_v3.md     # Multi-pass final audit
├── 10_Numbers_You_Must_Memorize.md  # Key metrics with mnemonics
├── [Presentation]_v[N].pptx
├── [Data_Backup]_v[N].xlsx
└── Submitted/             # Final PDF
```

### Adaptation by Case Type

- **IB cases**: `04_Investment_Lens_Critique.md` (thesis, returns, risks)
- **IR cases**: `04_Infrastructure_Lens_Critique.md` (firm language, criteria mapping)
- **Interview prep cases**: Add `11_Interview_Brief.md` (compiled prep document)
- **Quick turnaround**: Skip numbered files 04-10; use `!Archive/` and `resources/` only

---

## Standard 6: Key Metrics Table

*Source of truth for all numbers used in any deliverable.*

### Template

```markdown
# Key Metrics Table — [Case Name]

*Last updated: [date]*
*All deliverables MUST reference this table. Update here first, then propagate.*

## Transaction Metrics

| Metric | Value | Source |
|--------|-------|--------|
| Enterprise Value | $X.XB | [source] |
| Equity Value | $X.XB | [source] |
| ...

## Operational Metrics

| Metric | Value | Source |
|--------|-------|--------|
| ...

## Financial Metrics

| Metric | Value | Source |
|--------|-------|--------|
| ...

## Valuation Metrics

| Metric | Value | Source |
|--------|-------|--------|
| ...

## Industry Context

| Metric | Value | Source |
|--------|-------|--------|
| ...

## Acquirer / Sponsor Context

| Metric | Value | Source |
|--------|-------|--------|
| ...

## Pro Forma Impact

| Metric | Before | After | % Change | Source |
|--------|--------|-------|----------|--------|
| ...
```

### Rules

- Every number used in any deliverable MUST appear here with a SOURCE citation
- Update this table FIRST when any number changes — then update all downstream files
- Paste this table into every Claude prompt that needs numbers (Key Metrics Table concept)

---

## Standard 7: Q&A Preparation

*50+ questions across 5 categories with model answers for any interview-linked case study.*

### Category Structure

| Category | # Questions | Focus |
|----------|-------------|-------|
| About Your Deliverables | 10-12 | Walk-throughs, rationale, what you'd add |
| Technical Knowledge | 10-12 | Domain mechanics, metrics, calculations |
| Target/Industry | 10-12 | Business model, competitors, risks, cycles |
| Role Fit & Motivation | 8-10 | Why this role, career path, skills |
| Curveballs & Judgment | 6-8 | "What if...", largest risk, exit strategy |

### Per-Question Format

For EACH question provide:
- **Model answer** (2-4 sentences with real data from the case)
- **Key numbers to cite** (specific metrics)
- **Common mistakes to avoid**

### Supporting Documents

- **60-second verbal pitch** — scripted, memorized, uses 5-7 key numbers
- **Numbers memorization doc** — 90+ metrics organized by category with context sentences, rapid-fire Q&A, and memory anchors/mnemonics
- **Cheat sheet** — 1-page rapid reference for all critical numbers

---

## Standard 8: Multi-Pass Audit Protocol

*5 audit passes before submission. Each pass catches different categories of issues.*

### Audit Passes

| Pass | Timing | Focus | Deliverable |
|------|--------|-------|-------------|
| 1 | After Phase 2 (narrative) | Terminology, framework, and lens critique | `04_[Lens]_Critique.md` |
| 2 | After Phase 3 (build) | Excel formula integrity, formatting compliance | Fix in workbook |
| 3 | After Phase 4 (Q&A) | Cross-deliverable consistency (31+ metrics) | `07_Consistency_Audit.md` |
| 4 | Before submission | Financial data corrections + number formatting | `09_Final_Audit_v3.md` |
| 5 | Final review | Visual polish, speaker notes, print areas | Final files |

### Audit Versioning

- `07_Consistency_Audit.md` → `09_Final_Audit_v2.md` → `09_Final_Audit_v3.md`
- Each version identifies new issues AND confirms prior fixes
- Old versions go to `!Archive/`

---

## Standard 9: Versioning Discipline

*Never overwrite, always version up. Applies to all case deliverables.*

### Rules (from Operational Rules R1, R3)

1. **Never delete files** — move superseded versions to `!Archive/`
2. **Version numbering** — start at `_v2`, increment on each iteration (`_v3`, `_v4`, ...)
3. **Active copy** — the highest version number is always the working copy
4. **Archive protocol** — after creating `_v(N+1)`, move `_v(N)` to `!Archive/`

### Applies To

- PowerPoint source files (`[Presentation]_v2.pptx` → `_v3.pptx`)
- Excel backup/data workbooks (`[Data_Backup]_v2.xlsx` → `_v3.xlsx`)
- Audit documents (multi-pass: `v1` → `v2` → `v3`)

### Does NOT Apply To

- Numbered markdown deliverables (`01_Research_Brief.md`) — these are updated in-place
- Source material files in `resources/` — these are append-only
- Memory files — these are updated in-place

---

## Standard 10: GS Excel Formatting

*Apply-as-you-go. Full specification: `memory/gs_excel_formatting.md`*

### Quick Reference

| Element | Standard |
|---------|----------|
| Font | Arial 10pt — no exceptions |
| Gridlines | OFF on every sheet |
| Merged cells | NEVER — use Center Across Selection |
| Headers | Dark blue fill `#002060` + white font `#FFFFFF` |
| Hardcoded values | Blue font `#0070C0` |
| Same-sheet formulas | Purple font `#7030A0` |
| Cross-sheet formulas | Green font `#00B050` |
| External links | Red font `#FF0000` |
| Negatives | Parentheses — NEVER minus signs |
| Print area | Set to content boundaries |
| Orientation | Landscape, fit to 1 page wide |

### Application Scope

Every Excel workbook created or edited across ALL skills:
- IB Toolkit models (LBO, DCF, operating models, comps)
- IR Toolkit workbooks (LP updates, DDQ, fundraising)
- Portfolio Tracker workbooks
- Ad-hoc analysis workbooks
- Case workspace Excel files (`output/cases/`)

**Full spec with openpyxl code:** `memory/gs_excel_formatting.md`
**Operational rule:** R7 in `memory/operational_rules.md`

---

## Key Documents

| Document | Location | Relationship |
|----------|----------|-------------|
| Case Study Workflow Patterns | `memory/case_study_workflow_patterns.md` | 10-pattern implementation reference |
| GS Excel Formatting | `memory/gs_excel_formatting.md` | Full formatting spec with openpyxl code |
| Operational Rules | `memory/operational_rules.md` | R1-R7 enforcement rules |
| IB Best Practices | `docs/skills/ib_toolkit/BEST_PRACTICES.md` | IB-specific slide and model standards |
| IR Best Practices | `src/ii_skills/ir_toolkit/BEST_PRACTICES.md` | IR-specific storytelling and risk frameworks |
| IR Prompt Library | `src/ii_skills/ir_toolkit/PROMPT_LIBRARY.md` | P31-P34 prompt implementations |

---

*Version: 1.0 | Created: February 2026*
*10 universal standards for all case work across IB Toolkit, IR Toolkit, Interview Prep, and future skills*
