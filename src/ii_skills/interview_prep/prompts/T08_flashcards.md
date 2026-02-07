# T08 — Flashcard Generation Guide: {{COMPANY_NAME}}

## Research Target
**Company:** {{COMPANY_NAME}}
**Role:** {{ROLE_TITLE}} ({{TARGET_FUNCTION}})

## Saved Research
{{RESEARCH_NOTES}}

---

## Instructions

Generate **30-50 flashcards** in Q&A format for Anki import. Each card should test one specific fact or concept relevant to the {{ROLE_TITLE}} interview at {{COMPANY_NAME}}.

### Card Format

Each card must be a JSON object with these fields:
```json
{
  "question": "What is [specific question]?",
  "answer": "Concise answer (1-3 sentences max)",
  "tags": "category_tag"
}
```

### Required Categories (aim for 5-8 cards per category)

#### Company Facts
- Founding date, HQ location
- AUM / revenue / market cap
- Number of employees
- Fund names and sizes
- Business segments
- Key leadership names and titles

#### Deal Statistics
- Recent transaction details (target, size, date)
- Notable exits and returns
- Deal count and average deal size
- Sector/geographic focus areas

#### Industry Metrics
- Market size and growth rate
- Key industry trends with data
- Competitor rankings or market share
- Regulatory milestones

#### Key Terminology
- Firm-specific terms and acronyms
- Industry jargon relevant to {{TARGET_FUNCTION}}
- Fund structure terminology
- Role-specific technical terms

#### Leadership & People
- CEO and senior leadership backgrounds
- Interviewer backgrounds: {{INTERVIEWER_NAMES}}
- Notable board members
- Recent hires or departures

#### Role Knowledge
- Key responsibilities from JD
- Success metrics for the role
- Team structure
- Tools and platforms used

---

## Output Format

Return cards as a JSON array:
```json
[
  {"question": "...", "answer": "...", "tags": "company_facts"},
  {"question": "...", "answer": "...", "tags": "deal_statistics"}
]
```

Keep answers concise (under 100 words). Questions should be specific enough to have one clear answer.
