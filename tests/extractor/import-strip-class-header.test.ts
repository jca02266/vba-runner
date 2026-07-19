/**
 * vba-extractor `import` の VBE 標準 .cls ヘッダー自動除去の回帰テスト。
 *
 * `export` が書き出す .cls（VBA プロジェクトの内部モジュールストリーム形式）は
 * `Attribute VB_Name = "..."` から始まりヘッダーを持たない。一方、Excel VBE の
 * 「ファイルのエクスポート」メニューが出力する標準 .cls は
 * `VERSION 1.0 CLASS` / `BEGIN ... END`（MultiUse 等のコンポーネント属性）を
 * 先頭に持つ。この標準形式の .cls を `import` にそのまま渡すと、ヘッダー文字列が
 * 不正な VBA コードとしてモジュールストリームへ書き込まれてしまう。
 * `import` はこのヘッダーを検出したら自動的に除去してから書き込む。
 */
import { execFileSync } from 'child_process';
import { mkdtempSync, readFileSync, writeFileSync, rmSync } from 'fs';
import { tmpdir } from 'os';
import { join } from 'path';
import { assert } from '../../test-libs/test-runner';

const XLSM = 'sample/excel/test.xlsm';
const CLI = 'tools/extractor/vba-extractor.ts';

const tmp = mkdtempSync(join(tmpdir(), 'vba-extractor-header-test-'));
const srcDir = join(tmp, 'src');
const outXlsm = join(tmp, 'out.xlsm');
const reexportDir = join(tmp, 'reexported');

try {
    // 1. export で内部形式（ヘッダーなし）の Class1.cls を取得
    execFileSync('npx', ['tsx', CLI, 'export', XLSM, srcDir], { stdio: 'pipe' });
    const original = readFileSync(join(srcDir, 'Class1.cls'), 'utf8');
    assert.ok(!/^VERSION\s/i.test(original), 'export 直後の .cls はヘッダーなし（前提確認）');

    // 2. VBE 標準エクスポート形式を模して、先頭にヘッダーブロックを人工的に付加
    const withHeader = 'VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  \'True\r\nEND\r\n' + original;
    writeFileSync(join(srcDir, 'Class1.cls'), withHeader);

    // 3. import（確認プロンプトに 'y' を渡す）
    const importOutput = execFileSync(
        'npx', ['tsx', CLI, 'import', XLSM, srcDir, outXlsm],
        { input: 'y\n', stdio: 'pipe' },
    ).toString();
    assert.ok(/stripped VBE-style class header from 1 \.cls file/.test(importOutput),
        'import がヘッダー除去を実行したことをログで報告する');

    // 4. 生成された xlsm を再 export し、モジュールソースにヘッダー文字列が
    //    混入しておらず、元の内容と完全に一致することを確認する
    execFileSync('npx', ['tsx', CLI, 'export', outXlsm, reexportDir], { stdio: 'pipe' });
    const reexported = readFileSync(join(reexportDir, 'Class1.cls'), 'utf8');
    assert.strictEqual(reexported, original, '再 export した .cls が元の内部形式ソースと完全一致（ヘッダー混入なし）');
    assert.ok(!/VERSION\s+1\.0\s+CLASS/i.test(reexported), 'モジュールストリームに VERSION ヘッダー文字列が残っていない');

    console.log('[PASS] VBE 標準 .cls ヘッダーが import 時に自動除去される');
} finally {
    rmSync(tmp, { recursive: true, force: true });
}

console.log('\n✅ import-strip-class-header: 全テスト通過');
