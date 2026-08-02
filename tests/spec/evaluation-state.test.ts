import { strict as assert } from 'node:assert';
import { spawnSync } from 'node:child_process';
import { existsSync, readFileSync, unlinkSync, writeFileSync } from 'node:fs';
import * as yaml from 'js-yaml';

const root = process.cwd();
const cli = 'scripts/eval.mjs';

function run(...args: string[]) {
    return spawnSync(process.execPath, [cli, ...args], {
        cwd: root,
        encoding: 'utf8',
    });
}

const validated = run('validate');
assert.equal(validated.status, 0, validated.stderr);
assert.match(validated.stdout, /validated \d+ evaluation records/);

const targetCandidate = 'FZ-GRAMMAR-003';
const persistedResult = `${root}/evaluation/states/${targetCandidate}.result.yml`;
const persistedEvents = `${root}/evaluation/states/${targetCandidate}.events.yml`;
const persistedResultBody = existsSync(persistedResult) ? readFileSync(persistedResult, 'utf8') : null;
const persistedEventsBody = existsSync(persistedEvents) ? readFileSync(persistedEvents, 'utf8') : null;
if (persistedResultBody !== null) unlinkSync(persistedResult);
if (persistedEventsBody !== null) unlinkSync(persistedEvents);
// All migrated campaign candidates may already have a fixed evaluation. Temporarily
// hide this candidate's record so the state-machine claim assertion exercises a
// genuinely queued path; restore it immediately after claiming and at cleanup.
const targetEvaluation = `${root}/evaluation/evaluations/EV-00194.md`;
const targetEvaluationBody = readFileSync(targetEvaluation, 'utf8');
unlinkSync(targetEvaluation);

const first = run('next');
assert.equal(first.status, 0, first.stderr);
const candidate = JSON.parse(first.stdout);
assert.equal(candidate.status, 'queued');

const context = run('context', 'FZ-GRAMMAR-001');
assert.equal(context.status, 0, context.stderr);
assert.equal(JSON.parse(context.stdout).candidate.effectiveStatus, 'fixed');

const rerecorded = run('record', 'evaluation/evaluations/EV-00191.md');
assert.equal(rerecorded.status, 0, rerecorded.stderr);

const claimed = run('claim', targetCandidate);
assert.equal(claimed.status, 0, claimed.stderr);
const claimState = JSON.parse(claimed.stdout);
assert.equal(claimState.id, targetCandidate);
assert.match(claimState.token, /^[0-9a-f]+$/);
writeFileSync(targetEvaluation, targetEvaluationBody);

try {
    const whileClaimed = run('next');
    if (whileClaimed.status === 0) {
        assert.notEqual(JSON.parse(whileClaimed.stdout).id, targetCandidate,
            'claimed candidate is excluded while other candidates remain');
    } else {
        assert.match(whileClaimed.stderr, /no queued candidate/);
    }

    const duplicate = run('claim', targetCandidate);
    assert.notEqual(duplicate.status, 0);
    assert.match(duplicate.stderr, /already claimed/);

    const unauthorized = run('complete', targetCandidate, 'EV-00192', 'fixed', 'wrong-token');
    assert.notEqual(unauthorized.status, 0);
    assert.match(unauthorized.stderr, /token is invalid/);

    const mismatched = run('complete', targetCandidate, 'EV-00192', 'fixed', claimState.token);
    assert.notEqual(mismatched.status, 0);
    assert.match(mismatched.stderr, /belongs to .* not /);

    const completed = run('complete', targetCandidate, 'EV-00194', 'fixed', claimState.token);
    assert.equal(completed.status, 0, completed.stderr);

    const resultFile = `${root}/evaluation/states/${targetCandidate}.result.yml`;
    const validResult = readFileSync(resultFile, 'utf8');
    writeFileSync(resultFile, validResult.replace('evaluationId: EV-00194', 'evaluationId: EV-NOT-FOUND'));
    const invalidResult = run('validate');
    assert.notEqual(invalidResult.status, 0);
    assert.match(invalidResult.stderr, /unknown evaluation EV-NOT-FOUND/);
    writeFileSync(resultFile, validResult);

    const afterComplete = run('next');
    if (afterComplete.status === 0) {
        assert.notEqual(JSON.parse(afterComplete.stdout).id, targetCandidate,
            'completed candidate is excluded from the queue');
    } else {
        assert.match(afterComplete.stderr, /no queued candidate/);
    }
} finally {
    const result = `${root}/evaluation/states/${targetCandidate}.result.yml`;
    if (existsSync(result)) unlinkSync(result);
    if (existsSync(persistedEvents)) unlinkSync(persistedEvents);
    if (persistedResultBody !== null) writeFileSync(persistedResult, persistedResultBody);
    if (persistedEventsBody !== null) writeFileSync(persistedEvents, persistedEventsBody);
    writeFileSync(targetEvaluation, targetEvaluationBody);
    const released = run('release', targetCandidate, claimState.token);
    if (released.status !== 0 && !/not claimed/.test(released.stderr)) {
        assert.equal(released.status, 0, released.stderr);
    }
}

console.log('[PASS] evaluation state CLI claim/next/release');

// A resumable external check must preserve its status transitions separately
// from the latest result snapshot.
const transitionCandidate = 'FZ-BUILTIN-037';
const transitionResult = `${root}/evaluation/states/${transitionCandidate}.result.yml`;
const transitionEvents = `${root}/evaluation/states/${transitionCandidate}.events.yml`;
const transitionEvaluation = `${root}/evaluation/evaluations/EV-TEST-EVENTS.md`;
const oldTransitionResult = existsSync(transitionResult) ? readFileSync(transitionResult, 'utf8') : null;
const oldTransitionEvents = existsSync(transitionEvents) ? readFileSync(transitionEvents, 'utf8') : null;
const transitionRecord = `---
id: EV-TEST-EVENTS
candidateId: ${transitionCandidate}
campaign: FZ-BUILTIN
status: verified-no-bug
priority: low
focus: state event transition test
findings: []
tests:
  - tests/spec/evaluation-state.test.ts
horizontalAudit:
  confirmed: []
  ruledOut: []
  unresolved: []
---

# state event transition test
`;
if (oldTransitionResult !== null) unlinkSync(transitionResult);
if (oldTransitionEvents !== null) unlinkSync(transitionEvents);
writeFileSync(transitionEvaluation, transitionRecord);
let transitionToken: string | undefined;
try {
    const recordedTransition = run('record', transitionEvaluation);
    assert.equal(recordedTransition.status, 0, recordedTransition.stderr);
    const firstClaim = run('claim', transitionCandidate);
    assert.equal(firstClaim.status, 0, firstClaim.stderr);
    transitionToken = JSON.parse(firstClaim.stdout).token;
    const pending = run('complete', transitionCandidate, 'EV-00246', 'needs-excel', transitionToken);
    assert.equal(pending.status, 0, pending.stderr);
    const secondClaim = run('claim', transitionCandidate);
    assert.equal(secondClaim.status, 0, secondClaim.stderr);
    transitionToken = JSON.parse(secondClaim.stdout).token;
    const resolved = run('transition', transitionCandidate, 'EV-TEST-EVENTS', 'verified-no-bug', transitionToken);
    assert.equal(resolved.status, 0, resolved.stderr);
    const events = yaml.load(readFileSync(transitionEvents, 'utf8')) as Array<Record<string, string>>;
    assert.equal(events.length, 2);
    assert.equal(events[1].fromStatus, 'needs-excel');
    assert.equal(events[1].status, 'verified-no-bug');
} finally {
    const claim = `${root}/evaluation/states/${transitionCandidate}.claim.yml`;
    if (existsSync(claim)) unlinkSync(claim);
    if (existsSync(transitionResult)) unlinkSync(transitionResult);
    if (existsSync(transitionEvents)) unlinkSync(transitionEvents);
    if (oldTransitionResult !== null) writeFileSync(transitionResult, oldTransitionResult);
    if (oldTransitionEvents !== null) writeFileSync(transitionEvents, oldTransitionEvents);
    if (existsSync(transitionEvaluation)) unlinkSync(transitionEvaluation);
}

console.log('[PASS] evaluation state event history/resume');

// Cause analysis keeps the immediate implementation mechanism separate from
// the reusable design/root-cause grouping key.
const causeProbe = `${root}/evaluation/evaluations/EV-TEST-CAUSE.md`;
const causeProbeBody = `---
id: EV-TEST-CAUSE
candidateId: FZ-GRAMMAR-001
campaign: FZ-GRAMMAR
status: bug-found
priority: low
focus: cause schema test
directCauseKey: immediate-dispatch-gap
findings: []
tests:
  - tests/spec/evaluation-state.test.ts
horizontalAudit:
  confirmed: []
  ruledOut: []
  unresolved: []
---
`;
writeFileSync(causeProbe, causeProbeBody);
try {
    const invalidCause = run('record', causeProbe);
    assert.notEqual(invalidCause.status, 0);
    assert.match(invalidCause.stderr, /directCauseKey requires root causeKey/);
} finally {
    if (existsSync(causeProbe)) unlinkSync(causeProbe);
}

console.log('[PASS] evaluation cause-layer validation');
