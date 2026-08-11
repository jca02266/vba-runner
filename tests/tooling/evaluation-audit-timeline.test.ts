import { strict as assert } from 'node:assert';
import { spawnSync } from 'node:child_process';
import { existsSync, unlinkSync, writeFileSync } from 'node:fs';

const root = process.cwd();
const cli = 'scripts/eval.mjs';
const temporaryEvaluation = `${root}/evaluation/evaluations/EV-TEST-AUDIT-TIMELINE.md`;
const temporaryRecord = `---
id: EV-TEST-AUDIT-TIMELINE
candidateId: LEGACY-AUDIT-TIMELINE
campaign: LEGACY
status: verified-no-bug
priority: low
focus: timeline coverage audit test
area: 評価基盤
areaSource: confirmed
findings: []
tests: []
horizontalAudit:
  confirmed: []
  ruledOut: []
  unresolved: []
---

# timeline coverage audit test
`;

function runAudit() {
    return spawnSync(process.execPath, [cli, 'audit'], {
        cwd: root,
        encoding: 'utf8',
    });
}

if (existsSync(temporaryEvaluation)) unlinkSync(temporaryEvaluation);
writeFileSync(temporaryEvaluation, temporaryRecord);
try {
    const audit = runAudit();
    assert.notEqual(audit.status, 0, 'audit must reject an EV absent from events and snapshots');
    assert.match(audit.stderr, /EV-TEST-AUDIT-TIMELINE\(verified-no-bug\)/);
} finally {
    if (existsSync(temporaryEvaluation)) unlinkSync(temporaryEvaluation);
}

const restored = runAudit();
assert.equal(restored.status, 0, restored.stderr);
console.log('[PASS] evaluation audit detects timeline gaps');
