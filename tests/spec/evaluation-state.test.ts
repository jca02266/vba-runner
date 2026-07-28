import { strict as assert } from 'node:assert';
import { spawnSync } from 'node:child_process';
import { existsSync, readFileSync, unlinkSync, writeFileSync } from 'node:fs';

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
assert.match(validated.stdout, /validated 104 evaluation records/);

const targetCandidate = 'FZ-GRAMMAR-003';
const persistedResult = `${root}/evaluation/states/${targetCandidate}.result.yml`;
const persistedResultBody = existsSync(persistedResult) ? readFileSync(persistedResult, 'utf8') : null;
if (persistedResultBody !== null) unlinkSync(persistedResult);

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

try {
    const whileClaimed = run('next');
    assert.equal(whileClaimed.status, 0, whileClaimed.stderr);
    assert.notEqual(JSON.parse(whileClaimed.stdout).id, targetCandidate);

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
    assert.equal(afterComplete.status, 0, afterComplete.stderr);
    assert.notEqual(JSON.parse(afterComplete.stdout).id, targetCandidate);
} finally {
    const result = `${root}/evaluation/states/${targetCandidate}.result.yml`;
    if (existsSync(result)) unlinkSync(result);
    if (persistedResultBody !== null) writeFileSync(persistedResult, persistedResultBody);
    const released = run('release', targetCandidate, claimState.token);
    if (released.status !== 0 && !/not claimed/.test(released.stderr)) {
        assert.equal(released.status, 0, released.stderr);
    }
}

console.log('[PASS] evaluation state CLI claim/next/release');
