# Interview Prep Skill

Structured interview preparation materials generator. Produces company research briefs, Q&A frameworks, cheat sheets, and Anki flashcards — all saved to a Tier 3 case workspace.

## Architecture

The skill provides **structure, file I/O, templates, Anki export, and state tracking**. Claude Code's intelligence handles the actual research and content generation via WebSearch/WebFetch.

## 10 Actions

| # | Action | Description |
|---|--------|-------------|
| 1 | `start_prep` | Create case workspace, initialize context and section checklist |
| 2 | `load_template` | Get a prompt template (T01-T09) with placeholders filled from context |
| 3 | `save_research` | Store raw web research content for a URL |
| 4 | `write_section` | Save a completed section to the correct subfolder |
| 5 | `export_flashcards` | Generate Anki TSV + readable Markdown from card data |
| 6 | `compile_brief` | Merge all completed sections into a single 10-page brief |
| 7 | `get_prep_status` | Show progress on all 9 sections |
| 8 | `list_preps` | List all interview prep case workspaces |
| 9 | `get_research_context` | Retrieve stored research notes and completed sections |
| 10 | `update_context` | Update prep metadata (interviewer names, date, etc.) |

## Prompt Templates (T01-T09)

| ID | Template | Target Output | Words |
|----|----------|--------------|-------|
| T01 | Company Deep Dive | Company overview, leadership, culture, positioning | 3,000-5,000 |
| T02 | Deal/Transaction Database | Recent deals, exits, themes in table format | 2,000-3,000 |
| T03 | Industry Context | Market overview, trends, competitive landscape | 1,500-2,500 |
| T04 | Stakeholder Analysis | Investor base, key relationships, priorities | 1,000-1,500 |
| T05 | Role-Specific Intelligence | Team structure, responsibilities, success metrics | 1,000-1,500 |
| T06 | Interview Q&A Framework | Top questions ranked by probability + STAR behavioral | 2,000-3,000 |
| T07 | Quick Reference Cheat Sheet | Key stats, names, deals in bullet format | ~1,000 |
| T08 | Flashcard Generation | 30-50 Anki-ready Q&A cards in JSON | N/A |
| T09 | Why This Company? | 400-word, 90-second spoken response | ~400 |

## Case Workspace Layout

```
output/cases/YYYY-MM-DD_<company>_interview_prep/
├── case_config.json
├── research_notes.json
├── 01_company_research/
│   ├── company_deep_dive.md      (T01)
│   ├── industry_context.md       (T03)
│   └── stakeholder_analysis.md   (T04)
├── 02_deals_transactions/
│   └── deal_database.md          (T02)
├── 03_interview_qa/
│   ├── role_intelligence.md      (T05)
│   ├── interview_qa_framework.md (T06)
│   └── why_this_company.md       (T09)
├── 04_cheat_sheets/
│   └── quick_reference.md        (T07)
├── 05_flashcards/
│   ├── interview_prep_cards.tsv  (Anki import)
│   └── interview_prep_cards.md   (Human-readable)
└── 06_brief/
    └── full_brief.md             (Compiled document)
```

## Typical Workflow

```
User: "Prepare me for an interview at Brookfield for an IR Associate role"

Claude Code:
  1. start_prep(company_name="Brookfield", role_title="IR Associate", ...)
  2. WebSearch/WebFetch research → save_research() for each URL
  3. For each T01-T09:
     a. load_template(section_id, prep_folder) → filled prompt
     b. Generate content using intelligence + research
     c. write_section(prep_folder, section_id, content)
  4. export_flashcards(prep_folder, cards=[...])
  5. compile_brief(prep_folder)
  6. Report completion with file paths
```

## Data Models

### PrepContext
Core interview context: company name, role title, target function, URLs, job description, interview date, interviewer names, interview format, background points.

### SectionStatus
Per-section tracking: section_id, name, status (pending/in_progress/complete), file path, word count.

### InterviewBrief
Full state object wrapping PrepContext + list of SectionStatus + research notes dict.

## Integration Points

| Integration | How |
|-------------|-----|
| **Spaced Repetition** | TSV format matches spaced_rep's exporters — same Anki import format |
| **Deal Memory** | After prep, company facts can be stored via deal_memory for future reference |
| **Case Workspace** | Custom subfolder template under `output/cases/` |

## Action Parameter Reference

### start_prep
| Param | Type | Required | Description |
|-------|------|----------|-------------|
| company_name | str | Yes | Target company |
| role_title | str | Yes | Role you're interviewing for |
| target_function | str | No | e.g., "Investor Relations" |
| urls | List[str] | No | URLs to scan for research |
| job_description | str | No | Full job description text |
| interview_date | str | No | Date of interview |
| interviewer_names | List[str] | No | Known interviewer names |
| interview_format | str | No | Phone, video, in-person, case study |
| background_points | List[str] | No | Your key experiences to highlight |

### load_template
| Param | Type | Required | Description |
|-------|------|----------|-------------|
| section_id | str | Yes | T01-T09 |
| prep_folder | str | Yes | Case workspace path |

### save_research
| Param | Type | Required | Description |
|-------|------|----------|-------------|
| prep_folder | str | Yes | Case workspace path |
| url | str | Yes | Source URL |
| content | str | Yes | Extracted content |

### write_section
| Param | Type | Required | Description |
|-------|------|----------|-------------|
| prep_folder | str | Yes | Case workspace path |
| section_id | str | Yes | T01-T09 |
| content | str | Yes | Section markdown |

### export_flashcards
| Param | Type | Required | Description |
|-------|------|----------|-------------|
| prep_folder | str | Yes | Case workspace path |
| cards | List[Dict] | Yes | Cards with question, answer, tags keys |

### compile_brief
| Param | Type | Required | Description |
|-------|------|----------|-------------|
| prep_folder | str | Yes | Case workspace path |

### get_prep_status / get_research_context / update_context
| Param | Type | Required | Description |
|-------|------|----------|-------------|
| prep_folder | str | Yes | Case workspace path |
| section_id | str | No | Filter to specific section (get_research_context) |
| updates | Dict | Yes (update_context) | Fields to update |
