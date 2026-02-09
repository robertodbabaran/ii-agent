---
name: jobs-applier-aihawk
description: Workflow wrapper for AIHawk (Jobs_Applier_AI_Agent_AIHawk) to automate job-search and application steps with guarded execution and resume/profile inputs.
source_repo: https://github.com/feder-cr/Jobs_Applier_AI_Agent_AIHawk
---

# Jobs Applier AIHawk

This skill integrates the public AIHawk project into this repository's skills catalog so agents can follow a consistent workflow for setup, configuration, and safe execution.

## When to use

Use this skill when a user asks to:

- automate job discovery and application workflows,
- run AIHawk against a prepared resume/profile,
- update AIHawk settings and run a monitored application session.

## Workflow

1. Confirm intent, target roles, locations, and constraints (remote/hybrid, salary range, visa, seniority).
2. Prepare required local inputs (resume/profile, credentials, environment variables).
3. Install AIHawk from the upstream repository listed in `references/source.md`.
4. Configure AIHawk parameters conservatively first (dry-run/limited-run if available).
5. Execute and log outcomes (jobs scanned, jobs applied, skips/rejections).
6. Produce a concise post-run summary with next-step tuning suggestions.

## Guardrails

- Never commit credentials, cookies, or session tokens.
- Require explicit user confirmation before large-scale or repeated application runs.
- Start with narrow filters to avoid spammy or low-quality applications.
- Keep an auditable run log in workspace outputs.

## Outputs

Expected outputs from an execution session:

- run configuration snapshot,
- execution log,
- summary report of applications attempted/succeeded/failed.

For upstream project details and installation entry point, see `references/source.md`.
