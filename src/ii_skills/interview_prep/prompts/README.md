# Interview Prep — Prompt Templates

9 templates (T01-T09) for structured interview preparation. Each template uses `{{PLACEHOLDER}}` syntax filled at runtime by `template_loader.py`.

## Available Placeholders

| Placeholder | Source |
|-------------|--------|
| `{{COMPANY_NAME}}` | `PrepContext.company_name` |
| `{{ROLE_TITLE}}` | `PrepContext.role_title` |
| `{{TARGET_FUNCTION}}` | `PrepContext.target_function` |
| `{{JOB_DESCRIPTION}}` | `PrepContext.job_description` |
| `{{INTERVIEW_DATE}}` | `PrepContext.interview_date` |
| `{{INTERVIEWER_NAMES}}` | `PrepContext.interviewer_names` (joined) |
| `{{BACKGROUND_POINTS}}` | `PrepContext.background_points` (bullet list) |
| `{{URL_LIST}}` | `PrepContext.urls` (bullet list) |
| `{{RESEARCH_NOTES}}` | Aggregated from `research_notes.json` |

## Template Index

| ID | Name | Target Words | Output Location |
|----|------|-------------|-----------------|
| T01 | Company Deep Dive | 3,000-5,000 | `01_company_research/` |
| T02 | Deal/Transaction Database | 2,000-3,000 | `02_deals_transactions/` |
| T03 | Industry Context | 1,500-2,500 | `01_company_research/` |
| T04 | Stakeholder Analysis | 1,000-1,500 | `01_company_research/` |
| T05 | Role-Specific Intelligence | 1,000-1,500 | `03_interview_qa/` |
| T06 | Interview Q&A Framework | 2,000-3,000 | `03_interview_qa/` |
| T07 | Quick Reference Cheat Sheet | ~1,000 | `04_cheat_sheets/` |
| T08 | Flashcard Generation | 30-50 cards | `05_flashcards/` |
| T09 | Why This Company? | ~400 | `03_interview_qa/` |

## Workflow

1. Claude Code calls `load_template(section_id, prep_folder)` to get filled prompt
2. Uses the prompt as a research/generation guide with WebSearch + its own reasoning
3. Calls `write_section(prep_folder, section_id, content)` with generated output
4. For T08: generates card JSON, then calls `export_flashcards(prep_folder, cards)`
