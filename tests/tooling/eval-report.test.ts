import { strict as assert } from 'node:assert';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

const outputDir = mkdtempSync(join(tmpdir(), 'vba-runner-eval-report-'));
const output = join(outputDir, 'report.html');
try {
    execFileSync(process.execPath, ['scripts/eval-report.mjs', '--output', output], {
        cwd: process.cwd(), stdio: 'pipe', encoding: 'utf8',
    });
    const html = readFileSync(output, 'utf8');
    assert.match(html, /id="finding-area-select"/);
    assert.match(html, /value="all">すべて/);
    assert.match(html, /findingAreaSeries/);
    assert.match(html, /findingChart\.data\.labels = selected\.labels/);
    assert.match(html, /selected\.labels/);
    assert.match(html, /findingChart\.data\.datasets\[0\]\.data/);
    assert.match(html, /相対収束度/);
    assert.match(html, /領域別の直近発見率 \/ 領域別の累積発見率/);
    assert.match(html, /評価件数の10%（切り上げ、ただし最低10件）/);
    assert.match(html, /#N\/A/);

    const markdown = join(outputDir, 'report.md');
    execFileSync(process.execPath, ['scripts/eval-report.mjs', '--output', markdown], {
        cwd: process.cwd(), stdio: 'pipe', encoding: 'utf8',
    });
    assert.match(readFileSync(markdown, 'utf8'), /^# 評価レポート/m);

    const csv = join(outputDir, 'report.csv');
    execFileSync(process.execPath, ['scripts/eval-report.mjs', '--output', csv], {
        cwd: process.cwd(), stdio: 'pipe', encoding: 'utf8',
    });
    assert.match(readFileSync(csv, 'utf8'), /^transitionedAt,evaluationId,status/m);

    const unsupported = join(outputDir, 'report.json');
    assert.throws(() => execFileSync(process.execPath, ['scripts/eval-report.mjs', '--output', unsupported], {
        cwd: process.cwd(), stdio: 'pipe', encoding: 'utf8',
    }), /supported extension/);
} finally {
    rmSync(outputDir, { recursive: true, force: true });
}

console.log('[PASS] evaluation report Finding area selector');
