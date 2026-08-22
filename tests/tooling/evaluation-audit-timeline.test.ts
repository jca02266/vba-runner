import { strict as assert } from 'node:assert';
import { spawnSync } from 'node:child_process';
import { existsSync, unlinkSync, writeFileSync } from 'node:fs';

const root = process.cwd();
const cli = 'scripts/eval.mjs';
const temporaryEvaluation = `${root}/evaluation/evaluations/EV-TEST-AUDIT-TIMELINE.md`;
const excelEvaluation = `${root}/evaluation/evaluations/EV-TEST-AUDIT-EXCEL.md`;
const excelResult = `${root}/evaluation/states/LEGACY-AUDIT-EXCEL.result.yml`;
const excelEvents = `${root}/evaluation/states/LEGACY-AUDIT-EXCEL.events.yml`;
const excelCampaign = `${root}/evaluation/campaigns/TEST-AUDIT-EXCEL.yml`;
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

const excelRecord = `---
id: EV-TEST-AUDIT-EXCEL
evaluationRecordVersion: 2
candidateId: LEGACY-AUDIT-EXCEL
campaign: TEST-AUDIT-EXCEL
status: in-progress
priority: low
focus: Excel state coverage audit test
area: 評価基盤
areaSource: confirmed
expectation:
  kind: hypothesis
  statement: Excel probe is still pending
  references: []
  verification: pending
excelProbeIds:
  - XL-999
findings: []
tests: []
---

# Excel state coverage audit test

## 評価内容
Excel待ち状態の監査を確認する。

## 実施方法
未準備のXLプローブを持つ評価をin-progressで監査する。

## 期待値
期待出力: needs-excel-probeの不整合を報告する。

## 結果
実測出力: 監査が不整合を報告する。

## 判定
判定状態: in-progress
`;
writeFileSync(excelEvaluation, excelRecord);
writeFileSync(excelCampaign, 'id: TEST-AUDIT-EXCEL\nitems:\n  - id: LEGACY-AUDIT-EXCEL\n    title: Excel state audit test\n    status: queued\n    priority: 99\n    focus: Excel state coverage audit test\n');
writeFileSync(excelResult, 'stateVersion: 1\ncandidateId: LEGACY-AUDIT-EXCEL\nevaluationId: EV-TEST-AUDIT-EXCEL\nstatus: in-progress\n');
writeFileSync(excelEvents, '- candidateId: LEGACY-AUDIT-EXCEL\n  evaluationId: EV-TEST-AUDIT-EXCEL\n  status: in-progress\n  occurredAt: \'2026-08-22T00:00:00.000Z\'\n');
try {
    const excelAudit = runAudit();
    assert.notEqual(excelAudit.status, 0, 'audit must reject an in-progress evaluation whose Excel queue requires needs-excel-probe');
    assert.match(excelAudit.stderr, /EV-TEST-AUDIT-EXCEL\(in-progress, expected needs-excel-probe\)/);
} finally {
    for (const file of [excelEvaluation, excelResult, excelEvents, excelCampaign]) {
        if (existsSync(file)) unlinkSync(file);
    }
}

const restoredAfterExcel = runAudit();
assert.equal(restoredAfterExcel.status, 0, restoredAfterExcel.stderr);
console.log('[PASS] evaluation audit detects timeline gaps');
