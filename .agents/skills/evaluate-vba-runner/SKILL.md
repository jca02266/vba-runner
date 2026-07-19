---
name: evaluate-vba-runner
description: Evaluate vba-runner from a fresh user perspective, investigate a verified defect, implement its fix, add regression coverage, update the evaluation log, and commit the result. Use when asked to evaluate vba-runner usability or compatibility end-to-end, especially when the task should proceed from bug discovery through a tested fix and commit.
---

# Evaluate vba-runner

Use the project-owned Claude command and log as the single source of truth; do not duplicate their content in this skill.

1. Read `.claude/commands/evaluate-vba-runner.md` in full, then read `EVAL_LOG.md` in full before selecting a theme.
2. Follow the command's evaluation procedure. It explicitly calls for an independent evaluation agent, so use one. Keep its files outside the repository and do not let it read TODO files or git history.
3. Pick an untested, high-risk branch from `EVAL_LOG.md`, not a previously confirmed happy path. Give the evaluator a concrete domain and focus area. If the first pass finds no defect, make one further targeted pass against a different untested boundary; never invent a defect merely to satisfy the workflow.
4. Independently reproduce every reported defect with the smallest practical command or scratch program. Do not change tracked files for unverified reports. Identify the responsible parser, evaluator, builtin, or LSP code before editing.
5. For a verified defect, implement the smallest compatible fix. Add a regression test in the suite named by the Claude command, run that focused test and relevant typecheck/lint checks, then update `EVAL_LOG.md` as required by the command.
6. Review the diff, ensure only task-scoped files changed, and commit the fix, test, and log entry together. Use a concise conventional commit message such as `Fix: ...`.
7. Report the evaluated scenario, independently verified behavior, root cause, tests run, and commit hash. If no defect was verified, say so plainly and do not commit an invented fix.

Preserve existing user changes. Do not publish packages, tag releases, or broaden the task unless separately requested.
