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
6. After reproduction and horizontal expansion, follow the command's
   root-cause analysis gate before editing. Delegate an independent sub-agent
   to identify the direct cause, deeper design/state/flow cause, confirmed and
   ruled-out analogues, confidence, and remediation options. Update the Finding
   and evaluation record with that analysis before starting implementation.
   Once the cause is confirmed, launch a separate bounded implementation task,
   then have root verify that it addresses the recorded cause and adds
   regression coverage for the original and analogous paths. Commonization
   gaps are only one possible cause; also consider conversion, lifetime,
   boundaries, evaluation order, error propagation, and name resolution. If
   the cause remains uncertain, do not fix speculatively; record the unresolved
   analysis and run the smallest additional check.
7. For every evaluation run, write the structured evaluation record before
   staging a commit. Record the behavior, horizontal-expansion scope,
   `directCauseKey` (immediate mechanism), `causeKey` (deeper design/root
   cause), separate direct/root remediation statuses,
   confirmed/rule-out/unresolved paths, coverage reference, and next
   action. For a verified defect, implement the smallest compatible fix, add
   a regression test, and run focused checks. Also assess structural
   weaknesses and update `TODO.md` or `TODO_SPEC.md` only when justified.
   When a regression test or the full regression suite exposes a defect,
   mark the linked Finding with `discoveryType: regression`; use a different
   explicit type for fuzzing, mutation, coverage, Excel comparison, or direct
   evaluation findings. Do not infer a historical type without evidence.
   `priorCauseKey` is only a pre-evaluation hypothesis; it is not a confirmed
   cause. Keep `directFixStatus` and `rootFixStatus` separate. If root
   remediation is a follow-up candidate, record its candidate ID and commit
   when fixed; do not mark the root fixed merely because the symptom
   disappeared.
   For every new Excel-comparison candidate, reserve the next unique `XL-xxx`
   in the evaluation record's `excelProbeIds`. Then add that probe to
   `tests/excel/queue/ExcelQueueVerification.bas` and call it from the queue
   runner. Add any required `.cls`/`.frm` fixture beside that module, and
   reference the source files in the record's `tests`. A scratch-only probe is
   not a valid Excel queue entry.
   Before classifying a difference as a bug, add an `expectation` block with
   `kind: spec`, `excel`, or `hypothesis`, a concrete `statement`, references,
   and `verification`. Specification and Excel expectations require a cited
   source and completed verification. A hypothesis may remain pending only
   while the evaluation is non-terminal; verify it against the specification,
   a minimal runtime check, or Excel before using `bug-found` or `fixed`. If
   verification disproves the expectation, reclassify the evaluation and do
   not implement a speculative fix.
   Track Excel work with the two statuses `needs-excel-probe` (one or more
   required probes are absent from `ExcelQueueVerification.bas`) and
   `needs-excel` (all required probes are present, but a synchronized Excel
   result is pending). Record every required ID in `excelProbeIds`; `tests` is
   descriptive and is not used for state inference. After editing the queue or
   receiving a result, run `npm run eval -- excel-sync <evaluation-id>`.
   Advance to `needs-excel` only when it reports that state, and classify the
   result only when it reports `result-ready`. The queue result is ready only
   when it contains all required IDs, the end-of-run completion marker, and a
   SHA-256 matching the current normalized VBA source. `eval audit` never
   changes status. `validate` rejects a pending status that disagrees with the
   derived phase and rejects a ready result left in a pending state. Excel
   verification is complete only when the evaluation leaves these statuses
   for `verified-no-bug`, `known-limit`,
   `bug-found`, or `fixed`; never add a separate completion sub-status.
8. Treat `validate` and deterministic `render` as commit gates. Do not commit
   a fix, test, or state transition until the structured record is valid and
   the generated root `EVAL_LOG.md` is up to date. Do not create or commit an
   `evaluation/EVAL_LOG.generated.md` duplicate. Commit the implementation,
   regression test, structured record, and generated view together. When a
   candidate changes from `needs-excel`, `blocked`, or `in-progress`, use the
   `transition` command rather than editing its result snapshot; this appends
   the state event and preserves the previous state for later reporting.
   When Excel output is supplied, run `excel-sync` before interpretation, map
   every output ID to its evaluation record, resolve or retain each
   `unresolved` boundary, then transition the candidate. A stale, partial, or
   source-mismatched result is rejected and never changes the pending count.
9. After each evaluation commit, refresh the local HTML report with
   `node scripts/eval-report.mjs --html evaluation/EVAL_REPORT.html`.
   The HTML is generated output and is ignored by Git; do not stage or commit
   it. Use the refreshed report for the next loop's status and convergence
   review.
10. Select verification by changed-surface before running tests. For changes
   limited to `scripts/eval.mjs`, `scripts/eval-report.mjs`, evaluation
   records, or documentation, run the focused evaluation-state test,
   `npm run eval -- validate`, report generation, and documentation checks;
   do not run `scripts/bg-test.sh`. Run `bg-test.sh` only when engine,
   builtin, parser, extension, shared test harness, or runtime-facing code
   changed, or when a focused check exposes a cross-surface failure. Record
   the selected verification tier in the evaluation record and use a full
   background run before declaring such code changes regression-safe.
11. Report the evaluated scenario, independently verified behavior,
   horizontal-expansion investigation (including ruled-out paths), root cause,
   tests run, state transition, and commit hash. If no defect was verified,
   say so plainly and record the appropriate no-bug or limitation state.

Preserve existing user changes. Do not publish packages, tag releases, or broaden the task unless separately requested.
