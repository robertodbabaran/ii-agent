# Interview Prep — Quick Reference

## User Says → Action

| User Request | Action | Key Params |
|-------------|--------|------------|
| "Prepare me for an interview at X" | `start_prep` | company_name, role_title, urls |
| "Load the company deep dive template" | `load_template` | section_id="T01" |
| "Save this research" | `save_research` | url, content |
| "Write the industry analysis section" | `write_section` | section_id="T03", content |
| "Generate flashcards" | `export_flashcards` | cards (list of Q/A dicts) |
| "Compile the full brief" | `compile_brief` | prep_folder |
| "How far along is the prep?" | `get_prep_status` | prep_folder |
| "List my interview preps" | `list_preps` | (none) |
| "Show me the research so far" | `get_research_context` | prep_folder |
| "Add interviewer name: John Smith" | `update_context` | updates={"interviewer_names": [...]} |
| "Change interview date to March 15" | `update_context` | updates={"interview_date": "2026-03-15"} |

## Section ID → Content Type

| ID | Section | Output File | Folder |
|----|---------|-------------|--------|
| T01 | Company Deep Dive | `company_deep_dive.md` | `01_company_research/` |
| T02 | Deal Database | `deal_database.md` | `02_deals_transactions/` |
| T03 | Industry Context | `industry_context.md` | `01_company_research/` |
| T04 | Stakeholder Analysis | `stakeholder_analysis.md` | `01_company_research/` |
| T05 | Role Intelligence | `role_intelligence.md` | `03_interview_qa/` |
| T06 | Interview Q&A | `interview_qa_framework.md` | `03_interview_qa/` |
| T07 | Cheat Sheet | `quick_reference.md` | `04_cheat_sheets/` |
| T08 | Flashcards | `interview_prep_cards.tsv` | `05_flashcards/` |
| T09 | Why This Company? | `why_this_company.md` | `03_interview_qa/` |

## Recommended Workflow Order

1. **start_prep** — Create workspace with company + role + URLs
2. **Research phase** — WebSearch/WebFetch → `save_research` for each source
3. **T01** Company Deep Dive — Foundation for all other sections
4. **T02** Deal Database — Transaction intelligence
5. **T03** Industry Context — Market backdrop
6. **T04** Stakeholder Analysis — Who matters and why
7. **T05** Role Intelligence — What the job really is
8. **T06** Interview Q&A — Core prep document
9. **T09** Why This Company — Must-nail answer
10. **T07** Cheat Sheet — Last-minute review card
11. **T08** Flashcards → `export_flashcards` — Spaced repetition study
12. **compile_brief** — Single document for final review

## Integration with Other Skills

| Want to... | Use... |
|------------|--------|
| Create a persistent flashcard deck | `spaced_repetition.create_deck` + `add_cards` |
| Store company facts for later | `deal_memory.remember_company_fact` |
| Build a study plan for the interview | `spaced_repetition.get_study_plan` |
