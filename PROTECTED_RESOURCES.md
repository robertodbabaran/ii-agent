# Protected Resources Manifest

This file defines the protection tiers for all files in the ii-agent system.
**The agent MUST check this manifest before writing to any file.**

---

## Tier 1 — Personal Vault (NEVER touch without explicit user request)

These files live OUTSIDE the git repository on the user's OneDrive Desktop.
They contain credentials and personal financial data that must never be committed to git.

**The agent must STOP and ask for explicit permission before modifying any Tier 1 file.**

### Credential Configs (source of truth — runner scripts auto-restore from here)
```
%USERPROFILE%\OneDrive\Desktop\RJ\II-Agent Credentials\credentials_backup.txt
%USERPROFILE%\OneDrive\Desktop\RJ\II-Agent Credentials\configs\networth_config.py
%USERPROFILE%\OneDrive\Desktop\RJ\II-Agent Credentials\configs\market_config.py
%USERPROFILE%\OneDrive\Desktop\RJ\II-Agent Credentials\configs\whoop_config.py
```

### Personal Documents (master copies — never committed to git)
```
%USERPROFILE%\OneDrive\Desktop\RJ\II-Agent Credentials\documents\Resume (Recruiter).docx
%USERPROFILE%\OneDrive\Desktop\RJ\II-Agent Credentials\documents\Sample Cover Letter (Generalist PE).docx
```
Working copies live in `src/ii_skills/career_toolkit/resources/` (untracked by git).
Vault copies are the OneDrive-synced backup.

### Repo Config Mirrors (auto-restored by runner .bat scripts — treat as read-only)
```
src/ii_skills/networth_newsletter/config.py
src/ii_skills/market_newsletter/config.py
src/ii_skills/health_dashboard/config.py
src/ii_skills/ib_toolkit/config.py
Config/                                       # Entire directory
```

### Rules
- Credentials are only stored in the Desktop vault, never committed to git
- Net worth holdings and personal financial inputs live in the vault only
- Repo config.py files are disposable mirrors — runner scripts copy from vault before each run
- Personal DOCX files (resume, cover letter) are NEVER committed — .gitignore blocks *.docx
- If vault files need updating, the agent must ask the user first and update BOTH the vault master AND the repo mirror

---

## Tier 2 — Canonical Reference Templates (READ-ONLY unless user explicitly requests edit)

These files represent the institutional knowledge, formatting standards, and reference
templates built over time. They are the "gold standard" that case outputs are generated from.

**The agent must STOP and ask for explicit permission before modifying any Tier 2 file.**
**When generating case output, COPY patterns from these files — never modify originals.**

### Reference Output Templates
```
src/ii_skills/ib_toolkit/templates/reference_outputs/    # 30 template files
src/ii_skills/ib_toolkit/templates/reference_outputs/Institutional_LBO_Template.xlsx
src/ii_skills/ib_toolkit/templates/reference_outputs/Institutional_Deck_Template.pptx
```

### Excel Model Infrastructure
```
templates/excel_models/excel_modules.py          # 33 module functions (8 smart formula)
templates/excel_models/formula_builder.py         # FormulaBuilder helper class
src/ii_skills/shared/excel_templates/             # 19 professionally formatted templates
```

### Slide & Presentation Infrastructure
```
src/ii_skills/ib_toolkit/templates/case_study/    # Slide generation code
```

### IB Toolkit Documentation
```
docs/skills/ib_toolkit/PROMPT_LIBRARY.md
docs/skills/ib_toolkit/ORCHESTRATION_FRAMEWORK.md
docs/skills/ib_toolkit/MODEL_AUDIT_CHECKLIST.md
docs/skills/ib_toolkit/EXCEL_CONVENTIONS.md
docs/skills/ib_toolkit/SECTOR_PLAYBOOKS.md
docs/skills/ib_toolkit/IC_PRESENTATION_PLAYBOOK.md
docs/skills/ib_toolkit/BEST_PRACTICES.md
docs/skills/ib_toolkit/CASE_INTAKE_PROTOCOL.md
docs/skills/ib_toolkit/SLIDE_TEMPLATES.md
docs/skills/SLIDE_DESIGN_GUIDE.md
```

### IB Toolkit Capability Docs
```
src/ii_skills/ib_toolkit/CAPABILITIES.md
src/ii_skills/ib_toolkit/LBO_CASE_GUIDE.md
src/ii_skills/ib_toolkit/QUICK_REFERENCE.md
```

### IR Toolkit Documentation
```
docs/skills/ir_toolkit/                           # All IR documentation
src/ii_skills/ir_toolkit/CAPABILITIES.md
src/ii_skills/ir_toolkit/CASE_GUIDE.md
src/ii_skills/ir_toolkit/PROMPT_LIBRARY.md
src/ii_skills/ir_toolkit/SLIDE_LAYOUTS.md
src/ii_skills/ir_toolkit/QUICK_REFERENCE.md
src/ii_skills/ir_toolkit/references/              # Ontology, schemas, jargon guide
src/ii_skills/ir_toolkit/modules/                 # Module specs
```

### Historical Work Samples
```
%USERPROFILE%\OneDrive\Desktop\RJ\Historical Work\   # TD Securities presentations, etc.
```

### Rules
- When generating case output, READ these files for patterns, formats, and standards
- NEVER write to these files during case work
- To ADD a new reference template: user must explicitly say "add this to the reference templates"
- To MODIFY an existing template: user must explicitly say "update the [template name]"

---

## Tier 3 — Case Workspaces (Active working area — free to create and edit)

Case workspaces are ephemeral, per-engagement folders where all live case output goes.

**The agent can freely create and modify files in active case workspaces.**

### Location
```
output/cases/<date>_<case-name>/
```

### Standard Case Folder Structure
```
output/cases/YYYY-MM-DD_case-name/
├── case_config.yml          # Case metadata (company, deal type, stage, date)
├── excel/                   # Generated Excel models
├── slides/                  # Generated PowerPoint decks
├── notes/                   # Deal notes, assumptions, research
├── data/                    # Input data (CIMs, financials, etc.)
└── deliverables/            # Final packaged outputs
```

### Rules
- When user says "let's work on [case]" → create a new case folder
- ALL generated output goes in the case folder, never in Tier 2 template directories
- Reference Tier 2 templates for formatting and structure, but write output to Tier 3
- Case folders are self-contained and exportable
- Old case folders are preserved (never auto-deleted)

---

## Decision Matrix

Before writing ANY file, the agent checks:

| File Location | Action |
|---------------|--------|
| Tier 1 (vault / credentials / personal data) | **STOP** — ask user for explicit permission |
| Tier 2 (canonical templates / docs) | **STOP** — ask user for explicit permission |
| Tier 3 (active case workspace) | **PROCEED** — this is the working area |
| New file in repo (not in any tier) | **PROCEED with caution** — confirm if ambiguous |
| System files (CLAUDE.md, PROTECTED_RESOURCES.md) | **STOP** — ask user for explicit permission |

---

## Git Protection System

Three layers prevent personal data from reaching public GitHub:

### Layer 1: .gitignore
Blocks sensitive file types from being tracked:
- `**/config.py` — credential/holdings configs
- `**/*.docx` — personal documents (resume, cover letter)
- `**/*.log`, `**/networth_history.json` — runtime data
- `output/`, `outputs/` — generated workbooks and reports

### Layer 2: Pre-commit Hook (`.git/hooks/pre-commit`)
Scans all staged files before every commit and **blocks** if it finds:
- Personal email addresses
- Full name
- Hardcoded user paths (e.g. literal `C:\Users\<username>\`)
- `.docx` binary files
- Known real financial values

To bypass in emergencies: `git commit --no-verify`

### Layer 3: Sanitized Tracked Files
Files that are tracked by git contain ONLY placeholder/generic values:

| Tracked File | What's Sanitized | Real Data Lives In |
|---|---|---|
| `networth_newsletter/SKILL.md` | Example holdings (generic tickers, round numbers) | Vault `networth_config.py` |
| `market_newsletter/SKILL.md` | Generic asset class examples | Vault `market_config.py` |
| `daily_macro_metals_newsletter/settings.py` | Empty default recipient (uses env var) | `RECIPIENT_EMAIL` env var |
| All `.bat` runner files | `%USERPROFILE%` + `%~dp0` (no hardcoded paths) | Paths resolve at runtime |
| `CLAUDE.md`, `INTEGRATION_ANALYSIS.md` | No personal name or email | N/A |
| `analysis_slides_template.py` | No personal name in comments | N/A |

### History Rewrite (2026-02-12)
Git history was rewritten with `git-filter-repo` to remove:
- Resume and cover letter DOCX binary files from all commits
- All instances of personal email addresses from file content and commit messages
- All instances of full name from file content

After rewriting, **force push is required** to sync the remote:
```bash
git push origin develop --force-with-lease
```

### Vault ↔ Repo File Map

| Purpose | Vault (real data) | Repo (sanitized) |
|---|---|---|
| Net worth config | `Credentials\configs\networth_config.py` | `networth_newsletter/config.py` (untracked) |
| Market config | `Credentials\configs\market_config.py` | `market_newsletter/config.py` (untracked) |
| WHOOP config | `Credentials\configs\whoop_config.py` | `health_dashboard/config.py` (untracked) |
| Resume | `Credentials\documents\*.docx` | `career_toolkit/resources/` (untracked) |
| Cover letter | `Credentials\documents\*.docx` | `career_toolkit/resources/` (untracked) |

**Workflow**: Vault files are auto-copied to repo at runtime by `.bat` runner scripts.
The repo copies are disposable — they can be deleted and re-restored from the vault.

---

## Updating This Manifest

This file itself is protected. To add or remove paths:
- User must explicitly request the change
- Agent updates this manifest AND the corresponding CLAUDE.md section
