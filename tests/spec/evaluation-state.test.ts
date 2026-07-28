import { strict as assert } from 'node:assert';
import { spawnSync } from 'node:child_process';

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
assert.match(validated.stdout, /validated 9 evaluation records/);

const first = run('next');
assert.equal(first.status, 0, first.stderr);
const candidate = JSON.parse(first.stdout);
assert.equal(candidate.status, 'queued');

const claimed = run('claim', candidate.id);
assert.equal(claimed.status, 0, claimed.stderr);

try {
    const whileClaimed = run('next');
    assert.equal(whileClaimed.status, 0, whileClaimed.stderr);
    assert.notEqual(JSON.parse(whileClaimed.stdout).id, candidate.id);

    const duplicate = run('claim', candidate.id);
    assert.notEqual(duplicate.status, 0);
    assert.match(duplicate.stderr, /already claimed/);
} finally {
    const released = run('release', candidate.id);
    assert.equal(released.status, 0, released.stderr);
}

console.log('[PASS] evaluation state CLI claim/next/release');
