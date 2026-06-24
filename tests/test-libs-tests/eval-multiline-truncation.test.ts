/**
 * VBARunner.eval() に複数行（改行区切り）コードを渡した際、1行目が単独の式として
 * パース可能な場合に2行目以降が黒く無視されるバグのテスト
 *
 * 仕様バグ修正: evalExpression() の「単一式として全体を消費したか」判定が、parseExpressionPublic()
 * の直後が Newline であることだけを見ており、その Newline の後に何があるか確認していなかった。
 * `=` は VBA では代入とも等価比較とも解釈できるため、"x = 10\nDebug.Print 1" のような
 * 入力では1行目 "x = 10" が単独で「等価比較式」としてパースしきれてしまい、直後が
 * Newline であることから「全体を消費した」と誤判定 → 2行目以降が一切実行されず、
 * かつ `x = 10` も代入ではなく比較として評価されてしまっていた（x は未代入のまま）。
 * 修正: Newline の場合は、それ以降が（複数の Newline を挟んで）EOF であることまで
 * 確認するようにした。
 */
import { VBARunner, assert } from '../../test-libs/test-runner';

// --- 1. 1行目が単独の式としてパース可能でも、2行目以降が実行される ---
{
    const printed: string[] = [];
    const r = new VBARunner(null, { onPrint: (s: string) => printed.push(s) });
    r.eval('x = 10\nDebug.Print "after"');
    assert.deepStrictEqual(printed, ['after'], '1行目の後ろの Debug.Print が実行される');
    assert.strictEqual(r.eval('x'), 10, '1行目の x = 10 が代入として実行される（比較式として捨てられない）');
}
console.log('[PASS] 複数行コードの1行目が式としてパース可能でも後続行が黒く無視されない');

// --- 2. 既存の「単一式 + 末尾改行のみ」の挙動は維持される ---
{
    const r = new VBARunner();
    assert.strictEqual(r.eval('1 + 2 + 3\n'), 6, '末尾改行のみの単一式は引き続き値を返す');
    assert.strictEqual(r.eval('1 + 2 + 3'), 6, '改行なしの単一式は引き続き値を返す');
}
console.log('[PASS] 単一式（末尾改行あり/なし）の既存挙動は維持される');

console.log('\n✅ eval() 複数行入力の黒い無視バグ: 全テスト通過');
