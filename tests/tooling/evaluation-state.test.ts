import { strict as assert } from 'node:assert';
import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { existsSync, mkdirSync, readFileSync, readdirSync, unlinkSync, writeFileSync } from 'node:fs';
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

// Retired v2 findings must explicitly mark all follow-up work as unnecessary.
const retiredFinding = `${root}/evaluation/findings/BUG-00508.md`;
const retiredFindingBody = readFileSync(retiredFinding, 'utf8');
writeFileSync(retiredFinding, retiredFindingBody.replace(
    'horizontalAudit:\n  status: not-required',
    'horizontalAudit:\n  status: pending',
));
try {
    const invalidRetired = run('validate');
    assert.notEqual(invalidRetired.status, 0);
    assert.match(invalidRetired.stderr, /retired finding requires horizontalAudit\.status not-required/);
} finally {
    writeFileSync(retiredFinding, retiredFindingBody);
}
writeFileSync(retiredFinding, retiredFindingBody.replace(
    'followUpDisposition: not-required',
    'followUpDisposition: registered',
));
try {
    const missingFollowUp = run('validate');
    assert.notEqual(missingFollowUp.status, 0);
    assert.match(missingFollowUp.stderr, /registered follow-up disposition requires candidates/);
} finally {
    writeFileSync(retiredFinding, retiredFindingBody);
}

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

const rerecorded = run('record', 'evaluation/evaluations/EV-00803.md');
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
evaluationRecordVersion: 2
candidateId: ${transitionCandidate}
campaign: FZ-BUILTIN
status: in-progress
priority: low
focus: state event transition test
area: 評価基盤
areaSource: confirmed
findings: []
tests:
  - tests/tooling/evaluation-state.test.ts
horizontalAudit:
  confirmed: []
  ruledOut: []
unresolved: []
---

# state event transition test

## 評価内容
状態イベントの遷移を検証する。
## 実施方法
最小のv2評価記録を遷移させる。
## 期待値
プログラム上の期待出力: \`ok\`
遷移履歴が追記される。
## 結果
プログラム上の実測出力: \`ok\`
遷移履歴を確認した。
## 判定
判定状態: \`verified-no-bug\`
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
    const pending = run('complete', transitionCandidate, 'EV-TEST-EVENTS', 'in-progress', transitionToken);
    assert.equal(pending.status, 0, pending.stderr);
    const pendingSnapshot = yaml.load(readFileSync(transitionResult, 'utf8')) as Record<string, string>;
    assert.equal(pendingSnapshot.stateVersion, 1);
    assert.equal(pendingSnapshot.status, 'in-progress');
    assert.equal(pendingSnapshot.completedAt, undefined,
        'intermediate evaluation states must not receive completedAt');
    writeFileSync(transitionEvaluation,
        transitionRecord.replace('status: in-progress', 'status: verified-no-bug'));
    const secondClaim = run('claim', transitionCandidate);
    assert.equal(secondClaim.status, 0, secondClaim.stderr);
    transitionToken = JSON.parse(secondClaim.stdout).token;
    const resolved = run('transition', transitionCandidate, 'EV-TEST-EVENTS', 'verified-no-bug', transitionToken);
    assert.equal(resolved.status, 0, resolved.stderr);
    const resolvedSnapshot = yaml.load(readFileSync(transitionResult, 'utf8')) as Record<string, string>;
    assert.equal(resolvedSnapshot.stateVersion, 1);
    assert.equal(resolvedSnapshot.status, 'verified-no-bug');
    assert.equal(resolvedSnapshot.completedAt, undefined,
        'state events, rather than snapshots, own transition timestamps');
    const events = yaml.load(readFileSync(transitionEvents, 'utf8')) as Array<Record<string, string>>;
    assert.equal(events.length, 2);
    assert.equal(events[1].fromStatus, 'in-progress');
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
const causeProbe = `${root}/evaluation/EV-TEST-CAUSE.tmp.md`;
const causeProbeBody = `---
id: EV-TEST-CAUSE
evaluationRecordVersion: 2
candidateId: FZ-GRAMMAR-001
campaign: FZ-GRAMMAR
status: bug-found
priority: low
focus: cause schema test
area: 評価基盤
areaSource: confirmed
expectation:
  kind: hypothesis
  statement: cause schema is valid
  references: []
  verification: completed
directCauseKey: immediate-dispatch-gap
findings: []
tests:
  - tests/tooling/evaluation-state.test.ts
horizontalAudit:
  confirmed: []
  ruledOut: []
  unresolved: []
---

## 評価内容
原因キーを検証する。
## 実施方法
最小記録を検証する。
## 期待値
プログラム上の期待出力: \`ok\`
期待値は有効である。
## 結果
プログラム上の実測出力: \`ok\`
検証を実施した。
## 判定
判定状態: \`bug-found\`
`;
writeFileSync(causeProbe, causeProbeBody);
try {
    const invalidCause = run('record', causeProbe);
    assert.notEqual(invalidCause.status, 0);
    assert.match(invalidCause.stderr, /directCauseKey requires root causeKey/);
} finally {
    if (existsSync(causeProbe)) unlinkSync(causeProbe);
    const recordedCause = `${root}/evaluation/evaluations/EV-TEST-CAUSE.md`;
    if (existsSync(recordedCause)) unlinkSync(recordedCause);
}

console.log('[PASS] evaluation cause-layer validation');

// Root-cause hypotheses require an independent review before a remediation
// task can reference them. Remediation tasks are distinct from campaign items.
const rootCauseProbe = `${root}/evaluation/root-causes/RC-99999.md`;
const remediationProbe = `${root}/evaluation/remediations/ROOT-99999.md`;
const rootCauseBody = `---
id: RC-99999
status: confirmed
causeKey: vba-object-reference-value-classification-missing
statement: test root cause
proposedBy: test-proposer
procedureVersion: 1
originEvaluations:
  - EV-00390
evidence:
  - test evidence
alternatives: []
unresolved: []
reviews:
  - reviewer: independent-test-reviewer
    verdict: confirm
    summary: test confirmation
acceptanceCriteria:
  - test criterion
decision: create a remediation task
---
`;
const remediationBody = `---
id: ROOT-99999
procedureVersion: 1
status: queued
priority: low
title: test remediation
rootCauseId: RC-99999
causeKey: vba-object-reference-value-classification-missing
originEvaluations:
  - EV-00390
originFindings:
  - BUG-00386
scope:
  include:
    - test include
  exclude:
    - test exclude
acceptanceCriteria:
  - test criterion
tests: []
---
`;
mkdirSync(`${root}/evaluation/root-causes`, { recursive: true });
mkdirSync(`${root}/evaluation/remediations`, { recursive: true });
writeFileSync(rootCauseProbe, rootCauseBody);
writeFileSync(remediationProbe, remediationBody);
try {
    const validRootTask = run('validate');
    assert.equal(validRootTask.status, 0, validRootTask.stderr);
    const rootContext = run('remediation-context', 'ROOT-99999');
    assert.equal(rootContext.status, 0, rootContext.stderr);
    assert.equal(JSON.parse(rootContext.stdout).task.rootCauseId, 'RC-99999');
    assert.equal(JSON.parse(rootContext.stdout).task.procedureVersion, 1);

    writeFileSync(rootCauseProbe, rootCauseBody.replace('status: confirmed', 'status: hypothesis'));
    const prematureTask = run('validate');
    assert.notEqual(prematureTask.status, 0);
    assert.match(prematureTask.stderr, /root cause RC-99999 is not confirmed/);
} finally {
    if (existsSync(rootCauseProbe)) unlinkSync(rootCauseProbe);
    if (existsSync(remediationProbe)) unlinkSync(remediationProbe);
}

console.log('[PASS] root-cause review and remediation task validation');

// Expected behavior must have a verified source before a terminal judgment.
const expectationProbe = `${root}/evaluation/EV-TEST-EXPECTATION.tmp.md`;
const expectationProbeBody = `---
id: EV-TEST-EXPECTATION
evaluationRecordVersion: 2
candidateId: FZ-GRAMMAR-001
campaign: FZ-GRAMMAR
status: bug-found
priority: low
focus: expectation provenance test
area: 評価基盤
areaSource: confirmed
expectation:
  kind: hypothesis
  statement: an unverified expected value
  references: []
  verification: pending
findings: []
tests:
  - tests/tooling/evaluation-state.test.ts
horizontalAudit:
  confirmed: []
  ruledOut: []
  unresolved: []
---

## 評価内容
期待値の根拠を検証する。
## 実施方法
最小記録を検証する。
## 期待値
プログラム上の期待出力: \`pending\`
根拠が未検証である。
## 結果
プログラム上の実測出力: \`pending\`
まだ確定していない。
## 判定
判定状態: \`bug-found\`
`;
writeFileSync(expectationProbe, expectationProbeBody);
try {
    const invalidExpectation = run('record', expectationProbe);
    assert.notEqual(invalidExpectation.status, 0);
    assert.match(invalidExpectation.stderr, /hypothesis expectation must be verified/);
} finally {
    if (existsSync(expectationProbe)) unlinkSync(expectationProbe);
    const recordedExpectation = `${root}/evaluation/evaluations/EV-TEST-EXPECTATION.md`;
    if (existsSync(recordedExpectation)) unlinkSync(recordedExpectation);
}

console.log('[PASS] evaluation expectation provenance validation');

// New records cannot orphan themselves from the campaign manifest.
const orphanProbe = `${root}/evaluation/EV-TEST-ORPHAN.tmp.md`;
writeFileSync(orphanProbe, expectationProbeBody
    .replaceAll('EV-TEST-EXPECTATION', 'EV-TEST-ORPHAN')
    .replace('candidateId: FZ-GRAMMAR-001', 'candidateId: NOT-REGISTERED')
    .replace('verification: pending', 'verification: completed'));
try {
    const orphan = run('record', orphanProbe);
    assert.notEqual(orphan.status, 0);
    assert.match(orphan.stderr, /unknown candidate NOT-REGISTERED/);
} finally {
    if (existsSync(orphanProbe)) unlinkSync(orphanProbe);
    const recordedOrphan = `${root}/evaluation/evaluations/EV-TEST-ORPHAN.md`;
    if (existsSync(recordedOrphan)) unlinkSync(recordedOrphan);
}

console.log('[PASS] campaign candidate linkage validation');

// Excel synchronization reports missing source probes without mutating state.
const probeState = `${root}/evaluation/evaluations/EV-TEST-EXCEL-PROBE.md`;
const probeStateBody = `---
id: EV-TEST-EXCEL-PROBE
candidateId: FZ-GRAMMAR-001
campaign: FZ-GRAMMAR
status: needs-excel-probe
priority: low
focus: Excel probe state test
area: 評価基盤
areaSource: confirmed
excelProbeIds:
  - XL-999
findings: []
tests:
  - tests/excel/queue/ExcelQueueVerification.bas (XL-999)
horizontalAudit:
  confirmed: []
  ruledOut: []
  unresolved:
    - pending Excel result
---
`;
writeFileSync(probeState, probeStateBody);
try {
    const synchronized = run('excel-sync', 'EV-TEST-EXCEL-PROBE');
    assert.equal(synchronized.status, 0, synchronized.stderr);
    const state = JSON.parse(synchronized.stdout);
    assert.equal(state.requiredState, 'needs-excel-probe');
    assert.deepEqual(state.missingSourceIds, ['XL-999']);
    assert.match(readFileSync(probeState, 'utf8'), /^status: needs-excel-probe$/m);
} finally {
    if (existsSync(probeState)) unlinkSync(probeState);
}

console.log('[PASS] Excel probe synchronization');

// A result is usable only when all required IDs, the completion marker, and
// the normalized current source hash agree.
const queueDirectory = `${root}/tests/excel/queue`;
const queueResult = `${root}/tests/excel/queue/ExcelQueueVerification.result`;
const queueResultBody = existsSync(queueResult) ? readFileSync(queueResult, 'utf8') : null;
const readyState = `${root}/evaluation/evaluations/EV-TEST-EXCEL-READY.md`;
const readyStateBody = probeStateBody
    .replaceAll('EV-TEST-EXCEL-PROBE', 'EV-TEST-EXCEL-READY')
    .replace('status: needs-excel-probe', 'status: needs-excel')
    .replaceAll('XL-999', 'XL-001');
const normalizedSource = readdirSync(queueDirectory)
    .filter((name) => /\.(?:bas|cls|frm)$/i.test(name))
    .sort()
    .map((name) => {
        const source = readFileSync(`${queueDirectory}/${name}`, 'utf8')
            .replace(/\r\n/g, '\n')
            .replace(/^\s*Private Const QUEUE_SOURCE_SHA256\s+As String\s*=\s*"[^"]*"\s*$/mi,
                'Private Const QUEUE_SOURCE_SHA256 As String = "__QUEUE_SOURCE_SHA256__"');
        return `${name}\n${source}\n`;
    })
    .join('');
const sourceHash = createHash('sha256').update(normalizedSource, 'utf8').digest('hex');
writeFileSync(readyState, readyStateBody);
try {
    writeFileSync(queueResult,
        `XL-001 LOF=0 BYTES=\nQUEUE_COMPLETE=True\nQUEUE_SOURCE_SHA256=${sourceHash}\n`);
    const ready = run('excel-sync', 'EV-TEST-EXCEL-READY');
    assert.equal(ready.status, 0, ready.stderr);
    assert.equal(JSON.parse(ready.stdout).requiredState, 'result-ready');

    const unreconciled = run('validate');
    assert.notEqual(unreconciled.status, 0);
    assert.match(unreconciled.stderr, /synchronized Excel result is ready/);

    writeFileSync(queueResult,
        'XL-001 LOF=0 BYTES=\nQUEUE_COMPLETE=True\nQUEUE_SOURCE_SHA256=' + '0'.repeat(64) + '\n');
    const stale = run('excel-sync', 'EV-TEST-EXCEL-READY');
    assert.equal(stale.status, 0, stale.stderr);
    assert.equal(JSON.parse(stale.stdout).requiredState, 'needs-excel');
} finally {
    if (queueResultBody === null) unlinkSync(queueResult);
    else writeFileSync(queueResult, queueResultBody);
    if (existsSync(readyState)) unlinkSync(readyState);
}

console.log('[PASS] Excel result synchronization gate');

// The combined matrix command writes its Radix completion marker to a
// separate result file; that file must satisfy the same synchronization gate.
const radixResult = `${root}/tests/excel/queue/RadixMatrix.result`;
const radixResultBody = existsSync(radixResult) ? readFileSync(radixResult, 'utf8') : null;
const radixState = `${root}/evaluation/evaluations/EV-TEST-EXCEL-RADIX.md`;
const radixStateBody = `---
id: EV-TEST-EXCEL-RADIX
candidateId: FZ-GRAMMAR-001
campaign: FZ-GRAMMAR
status: needs-excel
priority: low
focus: Radix matrix result synchronization test
area: 評価基盤
areaSource: confirmed
excelProbeIds:
  - XL-176
findings: []
tests:
  - tests/excel/queue/RadixMatrixVerification.bas (XL-176)
horizontalAudit:
  confirmed: []
  ruledOut: []
  unresolved:
    - pending Radix result
---
`;
writeFileSync(radixState, radixStateBody);
try {
    writeFileSync(radixResult,
        `CASE=XL-176 KIND=Byte INPUT=&H100 ERR=6\nRADIX_MATRIX_COMPLETE=True\nQUEUE_SOURCE_SHA256=${sourceHash}\n`);
    const radixReady = run('excel-sync', 'EV-TEST-EXCEL-RADIX');
    assert.equal(radixReady.status, 0, radixReady.stderr);
    assert.equal(JSON.parse(radixReady.stdout).requiredState, 'result-ready');
} finally {
    if (radixResultBody === null) unlinkSync(radixResult);
    else writeFileSync(radixResult, radixResultBody);
    if (existsSync(radixState)) unlinkSync(radixState);
}

console.log('[PASS] Radix matrix result synchronization gate');
