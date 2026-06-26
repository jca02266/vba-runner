/**
 * eval() で組み込み関数の戻り値に算術演算を適用するテスト
 *
 * バグ修正: isCallableLeftmostLeaf() が CallExpression に対して true を返していたため、
 * `UBound(arr) + 1` のような式が曖昧な文と誤判定されて fast-path をスキップし、
 * フォールバックの文パーサーが `UBound(arr + 1)` として誤解析していた。
 * 配列オブジェクトへの加算で Error 424 が発生していた。
 */
import { VBARunner, assert } from '../../test-libs/test-runner';

const r = new VBARunner(null, { quiet: true });

// UBound の戻り値への算術
{
    r.eval('Dim arr(2) As Long');
    const result = r.eval('UBound(arr) + 1');
    assert.strictEqual(result, 3, 'UBound(arr) + 1 should return 3');
    console.log('[PASS] UBound(arr) + 1');
}

// Len の戻り値への算術
{
    const result = r.eval('Len("hello") + 1');
    assert.strictEqual(result, 6, 'Len("hello") + 1 should return 6');
    console.log('[PASS] Len("hello") + 1');
}

// 括弧ワークアラウンドも継続して動作すること
{
    const result = r.eval('(UBound(arr)) + 1');
    assert.strictEqual(result, 3, '(UBound(arr)) + 1 should return 3');
    console.log('[PASS] (UBound(arr)) + 1 (workaround form)');
}

// ネストした算術
{
    const result = r.eval('Len("ab") + Len("cd")');
    assert.strictEqual(result, 4, 'Len("ab") + Len("cd") should return 4');
    console.log('[PASS] Len("ab") + Len("cd")');
}

// 既存の代入・比較が壊れていないこと
{
    r.eval('Dim y As Long');
    r.eval('y = 10');
    assert.strictEqual(r.eval('y'), 10, 'y should be 10 after assignment');
    console.log('[PASS] assignment y = 10 still works');
}

console.log('\n✅ eval-builtin-arithmetic: 全テスト通過');
