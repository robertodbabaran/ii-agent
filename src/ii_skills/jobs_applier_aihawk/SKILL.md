---
name: jobs_applier_aihawk
description: Resume and cover letter tailoring engine powered by LLM. Takes a job description and generates tailored resume + cover letter from base templates.
source_repo: https://github.com/feder-cr/Jobs_Applier_AI_Agent_AIHawk
---

# Jobs Applier AIHawk

Skill for automatically tailoring resumes and cover letters to specific job descriptions. Uses the user's base resume and cover letter as source material, analyzes the target job description, and produces tailored DOCX outputs that emphasize relevant experience and skills.

## Actions (5)

| Action | Description |
|--------|-------------|
| `tailor_resume` | Generate a job-tailored resume from base resume + job description |
| `tailor_cover_letter` | Generate a job-tailored cover letter from base cover letter + job description |
| `full_application_package` | Generate both tailored resume + cover letter in one pass |
| `analyze_job` | Extract and summarize key requirements from a job description |
| `list_resources` | List available base resume and cover letter files |

## Workflow

1. User provides a job description (paste text or URL).
2. Skill analyzes the JD — extracts role, company, key skills, qualifications, and requirements.
3. Skill reads the base resume and cover letter from `resources/`.
4. LLM tailors each section of the resume to emphasize matching experience and keywords.
5. LLM generates a targeted cover letter referencing the specific role and company.
6. Outputs saved as DOCX files in `output/` with timestamped filenames.

## Resources

Base documents in `resources/`:
- `User Resume (Recruiter).docx` — Master resume
- `Sample Cover Letter (Generalist PE).docx` — Master cover letter template

## Outputs

Generated in `output/`:
- `{company}_{role}_Resume_{date}.docx`
- `{company}_{role}_Cover_Letter_{date}.docx`
- `{company}_{role}_JD_Analysis_{date}.md` (optional)

## Upstream Reference

This skill is inspired by [AIHawk Jobs Applier](https://github.com/feder-cr/Jobs_Applier_AI_Agent_AIHawk).
The upstream repo is cloned at `external/Jobs_Applier_AI_Agent_AIHawk/` for reference.

## Quick Commands

```
"Tailor my resume for this job: [paste JD]"
"Generate a cover letter for [Company] [Role]"
"Full application package for this posting: [paste JD]"
"Analyze this job description: [paste JD]"
```
