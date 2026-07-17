/**
 * Bug 29-H: .cls ファイルのヘッダー処理
 * VERSION 行はあるが BEGIN/END ブロックがない場合、本体が消えないことを検証する。
 */
import { stripVBAFileHeader } from '../../src/engine/preprocessor';
import { assert } from '../../test-libs/test-runner';

// 通常の .cls ファイル (BEGIN/END あり)
{
    const src = [
        'VERSION 1.0 CLASS',
        'BEGIN',
        '  MultiUse = -1',
        'END',
        'Attribute VB_Name = "Counter"',
        'Class Counter',
        'Private mVal As Long',
        'End Class',
    ].join('\n');
    const result = stripVBAFileHeader(src);
    assert.ok(!result.includes('BEGIN'), 'BEGIN は除去される');
    assert.ok(!result.includes('MultiUse'), 'MultiUse は除去される');
    assert.ok(result.includes('Class Counter'), '本体は保持される');
    console.log('[PASS] Bug 29-H: BEGIN/END あり → ヘッダーのみ除去');
}

// BEGIN/END ブロックなし (.cls の簡易形式)
{
    const src = [
        'VERSION 1.0 CLASS',
        'Attribute VB_Name = "Counter"',
        '',
        'Class Counter',
        'Private mVal As Long',
        'End Class',
    ].join('\n');
    const result = stripVBAFileHeader(src);
    assert.ok(result.includes('Class Counter'), 'Bug 29-H: BEGIN/END なしでも本体が保持される');
    assert.ok(!result.includes('VERSION'), 'VERSION 行は除去される');
    console.log('[PASS] Bug 29-H: BEGIN/END なし → VERSION 行のみ除去、本体保持');
}

// VERSION 行なし → 変更なし
{
    const src = 'Class Foo\nEnd Class\n';
    const result = stripVBAFileHeader(src);
    assert.strictEqual(result, src, 'Bug 29-H: VERSION なし → そのまま返す');
    console.log('[PASS] Bug 29-H: VERSION 行なし → 変更なし');
}

console.log('\n✅ Preprocessor CLS Header: 全テスト通過');
