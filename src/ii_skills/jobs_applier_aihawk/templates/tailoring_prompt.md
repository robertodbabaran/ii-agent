# Resume Tailoring Prompt Template

Use this prompt structure when Claude Code tailors a resume or cover letter.

---

## Resume Tailoring

Given the following:

**Base Resume:**
{base_resume_text}

**Job Description:**
{job_raw}

**Extracted Requirements:**
{job_requirements}

**Key Skills Sought:**
{job_skills}

**Keywords to Emphasize:**
{job_keywords}

### Instructions

Rewrite each section of the resume to:
1. **Lead with relevant experience** — reorder bullet points so the most JD-relevant ones come first.
2. **Mirror JD language** — use the same terminology the employer uses (e.g., if JD says "stakeholder management", use that phrase).
3. **Quantify achievements** — ensure every bullet has a metric where possible.
4. **Remove irrelevant details** — deprioritize experience that doesn't map to this role.
5. **Keep to {max_pages} pages** — be concise.
6. **Preserve factual accuracy** — never fabricate experience, only reframe existing content.

---

## Cover Letter Tailoring

Given the following:

**Base Cover Letter:**
{base_cover_letter_text}

**Target Company:** {job_company}
**Target Role:** {job_role}

### Instructions

Write a tailored cover letter (3 paragraphs, under {max_cover_letter_words} words):
1. **Opening** — State the role, express genuine interest, mention one specific thing about the company.
2. **Body** — Highlight 2-3 experiences from the resume that directly match JD requirements. Use specific metrics.
3. **Closing** — Tie your background to the company's mission/needs. Express enthusiasm for contributing.

Do NOT use generic phrases like "I am writing to express my interest." Be specific and direct.
