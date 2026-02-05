# Workflow Orchestration (Standalone Draft)

This document captures a draft workflow for disciplined task execution. It is intentionally standalone and **not** wired into runtime prompts or agent logic yet.

## Plan Mode Default
- Enter plan mode for any non-trivial task (3+ steps or architectural decisions).
- Write detailed specs upfront to reduce ambiguity.
- If something goes sideways, stop and re-plan immediately.
- Use plan mode for verification steps, not just building.

## Subagent Strategy
- Offload research, exploration, and parallel analysis to subagents.
- One task per subagent for focused execution.

## Self-Improvement Loop
- After any correction from the user, update `tasks/lessons.md` with the pattern and prevention rules.
- Review `tasks/lessons.md` at session start when it exists.

## Verification Before Done
- Do not mark a task complete without proving it works.
- Diff behavior between main and changes when relevant.
- Ask: "Would a staff engineer approve this?"
- Run tests, check logs, and demonstrate corrections.

## Demand Elegance (Balanced)
- For non-trivial changes: pause and ask "is there a more elegant way?"
- If a fix feels hacky, implement the elegant solution.
- Skip this for simple, obvious fixes.

## Autonomous Bug Fixing
- When given a bug report: fix it without requesting hand-holding.
- Use logs/errors/failing tests to locate and resolve issues.
- Resolve failing CI without prompting.

## Task Management (Repo Files)
- Plan first: write the plan to `tasks/todo.md` with checkable items.
- Verify plan: check in before starting implementation.
- Track progress: mark items complete as you go.
- Explain changes: provide high-level summary at each step.
- Document results: add review notes to `tasks/todo.md`.
- Capture lessons: update `tasks/lessons.md` after corrections.

## Core Principles
- Simplicity first: minimal, safe changes.
- No laziness: fix root causes, no temporary workarounds.
- Minimal impact: touch only what is necessary.
