---
name: evaluate-vba-runner
description: Evaluate vba-runner from a fresh user perspective, investigate a verified defect, implement its fix, add regression coverage, update the evaluation log, and commit the result. Use when asked to evaluate vba-runner usability or compatibility end-to-end, especially when the task should proceed from bug discovery through a tested fix and commit.
---

# Evaluate vba-runner

Use the project-owned Claude command and the evaluation-storage design as the
single source of truth; do not duplicate their content in this skill.

The storage migration and steady-state rules are defined in
`docs/internals/EVALUATION_STORAGE.md`. Read it before changing the evaluation
state. During migration, preserve `EVAL_LOG.md` as a read-only legacy source;
after migration, use the structured records and the generated Markdown view.

1. Read `.claude/commands/evaluate-vba-runner.md` and
   `docs/internals/EVALUATION_STORAGE.md` in full. Run the storage audit and
   obtain one compact candidate context before selecting a theme; do not load
   the entire historical log into every evaluator prompt.
2. Follow the command's evaluation procedure. It explicitly calls for an independent evaluation agent, so use one. Keep its files outside the repository and do not let it read TODO files or git history.
3. Claim one queued, high-risk branch selected by the evaluation-state CLI,
   using coverage gaps, prior cause keys, and unresolved boundaries. Give the
   evaluator only the candidate and its linked findings/audit context. If the
   first pass finds no defect, make one further targeted pass against a
   different queued boundary; never invent a defect merely to satisfy the
   workflow. A single fuzzer or mutation run never marks the method complete.
4. Independently reproduce every reported defect with the smallest practical command or scratch program. Do not change tracked files for unverified reports. Identify the responsible parser, evaluator, builtin, or LSP code before editing.
5. After a defect is independently reproduced, run a horizontal-expansion investigation before editing. Delegate a bounded subtask to a second independent sub-agent: inspect the source for every analogous dispatch/evaluation path and report suspicious sites, without changing tracked files or reading TODO files/git history. Have the sub-agent add a scratch driver outside the repository only when source inspection identifies a plausible analogue; use focused tests for those candidates rather than broad test-suite runs. Record each confirmed analogue, ruled-out path, or unresolved real-Excel semantic question in the evaluation notes, and do not treat an unverified suspicion as a product bug.
6. After implementing a verified defect, perform a root-cause design review
   before declaring the evaluation complete. Ask whether the defect is caused
   by duplicated conversion, dispatch, state, or validation logic rather than
   only by a bad branch or missing condition. If a common abstraction is
   missing because equivalent behavior is not shared, treat that lack of
   commonization as the root cause: fix the abstraction in the same work item
   (or create a justified refactoring TODO when it cannot be safely completed),
   then repeat the horizontal expansion against every caller of the
   abstraction. Add regression coverage for the original path and at least one
   analogous path. Record the original symptom, the root cause, the
   root-cause remediation, and the expanded paths in the Finding and
   evaluation record. A local patch that leaves the identified common cause in
   place is not a completed bug fix.
7. For every evaluation run, write the structured evaluation record before
   staging a commit. Record the behavior, horizontal-expansion scope, cause
   key, confirmed/rule-out/unresolved paths, coverage reference, and next
   action. For a verified defect, implement the smallest compatible fix, add
   a regression test, and run focused checks. Also assess structural
   weaknesses and update `TODO.md` or `TODO_SPEC.md` only when justified.
   When a regression test or the full regression suite exposes a defect,
   mark the linked Finding with `discoveryType: regression`; use a different
   explicit type for fuzzing, mutation, coverage, Excel comparison, or direct
   evaluation findings. Do not infer a historical type without evidence.
8. Treat `validate` and deterministic `render` as commit gates. Do not commit
   a fix, test, or state transition until the structured record is valid and
   the generated root `EVAL_LOG.md` is up to date. Do not create or commit an
   `evaluation/EVAL_LOG.generated.md` duplicate. Commit the implementation,
   regression test, structured record, and generated view together. When a
   candidate changes from `needs-excel`, `blocked`, or `in-progress`, use the
   `transition` command rather than editing its result snapshot; this appends
   the state event and preserves the previous state for later reporting.
9. After each evaluation commit, refresh the local HTML report with
   `node scripts/eval-report.mjs --html evaluation/EVAL_REPORT.html`.
   The HTML is generated output and is ignored by Git; do not stage or commit
   it. Use the refreshed report for the next loop's status and convergence
   review.
10. Report the evaluated scenario, independently verified behavior,
   horizontal-expansion investigation (including ruled-out paths), root cause,
   tests run, state transition, and commit hash. If no defect was verified,
   say so plainly and record the appropriate no-bug or limitation state.

Preserve existing user changes. Do not publish packages, tag releases, or broaden the task unless separately requested.
