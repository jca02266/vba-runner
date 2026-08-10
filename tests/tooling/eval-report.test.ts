import { strict as assert } from 'node:assert';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

const outputDir = mkdtempSync(join(tmpdir(), 'vba-runner-eval-report-'));
const output = join(outputDir, 'report.html');
try {
    execFileSync(process.execPath, ['scripts/eval-report.mjs', '--html', output], {
        cwd: process.cwd(), stdio: 'pipe', encoding: 'utf8',
    });
    const html = readFileSync(output, 'utf8');
    assert.match(html, /id="finding-area-select"/);
    assert.match(html, /value="all">すべて/);
    assert.match(html, /findingAreaSeries/);
    assert.match(html, /findingChart\.data\.labels = selected\.labels/);
    assert.match(html, /selected\.labels/);
    assert.match(html, /findingChart\.data\.datasets\[0\]\.data/);
} finally {
    rmSync(outputDir, { recursive: true, force: true });
}

console.log('[PASS] evaluation report Finding area selector');
