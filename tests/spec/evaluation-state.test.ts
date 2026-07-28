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
assert.match(validated.stdout, /validated 92 evaluation records/);

const persistedResult = `${root}/evaluation/states/FZ-BUILTIN-001.result.yml`;
const persistedResultBody = existsSync(persistedResult) ? readFileSync(persistedResult, 'utf8') : null;
if (persistedResultBody !== null) unlinkSync(persistedResult);

const first = run('next');
assert.equal(first.status, 0, first.stderr);
const candidate = JSON.parse(first.stdout);
assert.equal(candidate.status, 'queued');

const context = run('context', 'FZ-GRAMMAR-001');
assert.equal(context.status, 0, context.stderr);
assert.equal(JSON.parse(context.stdout).candidate.effectiveStatus, 'fixed');

const claimed = run('claim', candidate.id);
assert.equal(claimed.status, 0, claimed.stderr);
const claimState = JSON.parse(claimed.stdout);
assert.equal(claimState.id, candidate.id);
assert.match(claimState.token, /^[0-9a-f]+$/);

try {
    const whileClaimed = run('next');
    assert.equal(whileClaimed.status, 0, whileClaimed.stderr);
    assert.notEqual(JSON.parse(whileClaimed.stdout).id, candidate.id);

    const duplicate = run('claim', candidate.id);
    assert.notEqual(duplicate.status, 0);
    assert.match(duplicate.stderr, /already claimed/);

    const unauthorized = run('complete', candidate.id, 'EV-00186', 'verified-no-bug', 'wrong-token');
    assert.notEqual(unauthorized.status, 0);
    assert.match(unauthorized.stderr, /token is invalid/);

    const mismatched = run('complete', candidate.id, 'EV-00185', 'fixed', claimState.token);
    assert.notEqual(mismatched.status, 0);
    assert.match(mismatched.stderr, /belongs to .* not /);

    const completed = run('complete', candidate.id, 'EV-00186', 'verified-no-bug', claimState.token);
    assert.equal(completed.status, 0, completed.stderr);

    const resultFile = `${root}/evaluation/states/${candidate.id}.result.yml`;
    const validResult = readFileSync(resultFile, 'utf8');
    writeFileSync(resultFile, validResult.replace('evaluationId: EV-00186', 'evaluationId: EV-NOT-FOUND'));
    const invalidResult = run('validate');
    assert.notEqual(invalidResult.status, 0);
    assert.match(invalidResult.stderr, /unknown evaluation EV-NOT-FOUND/);
    writeFileSync(resultFile, validResult);

    const afterComplete = run('next');
    assert.equal(afterComplete.status, 0, afterComplete.stderr);
    assert.notEqual(JSON.parse(afterComplete.stdout).id, candidate.id);
} finally {
    const result = `${root}/evaluation/states/${candidate.id}.result.yml`;
    if (existsSync(result)) unlinkSync(result);
    if (persistedResultBody !== null) writeFileSync(persistedResult, persistedResultBody);
    const released = run('release', candidate.id, claimState.token);
    if (released.status !== 0 && !/not claimed/.test(released.stderr)) {
        assert.equal(released.status, 0, released.stderr);
    }
}

console.log('[PASS] evaluation state CLI claim/next/release');
