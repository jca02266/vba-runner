import { strict as assert } from 'node:assert';
import { spawnSync } from 'node:child_process';
import { existsSync, readFileSync, unlinkSync, writeFileSync } from 'node:fs';
import * as yaml from 'js-yaml';

const root = process.cwd();
const cli = 'scripts/eval.mjs';
const run = (...args: string[]) => spawnSync(process.execPath, [cli, ...args], {
    cwd: root,
    encoding: 'utf8',
});

const candidate = { id: 'FZ-GRAMMAR-004', campaign: 'FZ-GRAMMAR' };
const evaluationId = 'EV-TEST-V2-WORKFLOW';
const evaluationFile = `${root}/evaluation/evaluations/${evaluationId}.md`;
const resultFile = `${root}/evaluation/states/${candidate.id}.result.yml`;
const eventsFile = `${root}/evaluation/states/${candidate.id}.events.yml`;
const claimFile = `${root}/evaluation/states/${candidate.id}.claim.yml`;
const oldResult = existsSync(resultFile) ? readFileSync(resultFile, 'utf8') : null;
const oldEvents = existsSync(eventsFile) ? readFileSync(eventsFile, 'utf8') : null;
const oldClaim = existsSync(claimFile) ? readFileSync(claimFile, 'utf8') : null;
const body = `---
id: ${evaluationId}
evaluationRecordVersion: 2
candidateId: ${candidate.id}
campaign: ${candidate.campaign}
status: in-progress
priority: low
focus: v2 workflow CLI gate
area: 評価基盤
areaSource: confirmed
expectation:
  kind: hypothesis
  statement: workflow gate accepts the recorded shape
  references: []
  verification: completed
findings: []
tests: []
---

# v2 workflow gate

## 評価内容
状態遷移ゲートを確認する。

## 実施方法
最小の評価記録を検証する。

## 期待値
プログラム上の期待出力: \`ok\`
期待値は記録形式の検証を通過する。

## 結果
プログラム上の実測出力: \`ok\`
記録形式の検証を通過した。

## 判定
判定状態: \`verified-no-bug\`
`;

const legacyProbe = `${root}/evaluation/EV-TEST-V2-MISSING.tmp.md`;
writeFileSync(legacyProbe, body
    .replace(`id: ${evaluationId}`, 'id: EV-TEST-V2-MISSING')
    .replace(/^evaluationRecordVersion: 2\n/m, ''));
try {
    const rejected = run('record', legacyProbe);
    assert.notEqual(rejected.status, 0);
    assert.match(rejected.stderr, /evaluationRecordVersion: 2/);
} finally {
    if (existsSync(legacyProbe)) unlinkSync(legacyProbe);
}

if (oldResult !== null) unlinkSync(resultFile);
if (oldEvents !== null) unlinkSync(eventsFile);
if (oldClaim !== null) unlinkSync(claimFile);
writeFileSync(evaluationFile, body);
let token: string | undefined;
try {
    const recorded = run('record', evaluationFile);
    assert.equal(recorded.status, 0, recorded.stderr);
    const claimed = run('claim', candidate.id);
    assert.equal(claimed.status, 0, claimed.stderr);
    token = (JSON.parse(claimed.stdout) as { token: string }).token;
    const pending = run('complete', candidate.id, evaluationId, 'in-progress', token);
    assert.equal(pending.status, 0, pending.stderr);
    const snapshot = yaml.load(readFileSync(resultFile, 'utf8')) as { status: string };
    assert.equal(snapshot.status, 'in-progress');
    writeFileSync(evaluationFile, body.replace('status: in-progress', 'status: verified-no-bug'));
    const terminalClaim = run('claim', candidate.id);
    assert.equal(terminalClaim.status, 0, terminalClaim.stderr);
    token = (JSON.parse(terminalClaim.stdout) as { token: string }).token;
    const terminal = run('transition', candidate.id, evaluationId, 'verified-no-bug', token);
    assert.equal(terminal.status, 0, terminal.stderr);
    const rollbackClaim = run('claim', candidate.id, '--rollback');
    assert.equal(rollbackClaim.status, 0, rollbackClaim.stderr);
    token = (JSON.parse(rollbackClaim.stdout) as { token: string }).token;
    writeFileSync(evaluationFile, body);
    const rolledBack = run('rollback', candidate.id, evaluationId, 'in-progress', token, '期待値の再確認が必要');
    assert.equal(rolledBack.status, 0, rolledBack.stderr);
    const events = yaml.load(readFileSync(eventsFile, 'utf8')) as Array<Record<string, string>>;
    assert.equal(events.at(-1)?.rollbackFrom, 'verified-no-bug');
    assert.equal(events.at(-1)?.rollbackReason, '期待値の再確認が必要');
} finally {
    if (existsSync(evaluationFile)) unlinkSync(evaluationFile);
    if (existsSync(resultFile)) unlinkSync(resultFile);
    if (existsSync(eventsFile)) unlinkSync(eventsFile);
    if (existsSync(claimFile)) unlinkSync(claimFile);
    if (oldResult !== null) writeFileSync(resultFile, oldResult);
    if (oldEvents !== null) writeFileSync(eventsFile, oldEvents);
    if (oldClaim !== null) writeFileSync(claimFile, oldClaim);
}

console.log('[PASS] evaluation v2 record and transition gates');
