/**
 * VBARunner.eval() に裸の代入文（`x = 10` 単体、`arr(1) = "a"` 単体）を渡すと
 * 代入ではなく等価比較式として誤評価され、代入が一切実行されないバグのテスト
 *
 * 仕様バグ修正: evalExpression() は「単一式としてパースでき、入力全体を消費した」場合
 * 即座にその式を評価して値を返す高速パスを持つ。VBA の `=` は代入文の先頭にも等価比較
 * 演算子としても現れるため、`x = 10` や `arr(1) = "a"` のような入力は
 * parseExpressionPublic() に「x と 10 を比較する BinaryExpression」として完全に
 * パースされてしまい、高速パスの条件（入力全体を消費）を満たしてしまう。結果、
 * 代入が一切実行されないまま比較結果の真偽値だけが返っていた。
 * 修正: 入力全体を消費した式がトップレベルの `=` BinaryExpression である場合は
 * 高速パスを使わず、文として解析・実行するフォールバックに委ねるようにした
 * （実 VBA でも裸の `lhs = rhs` 文は常に代入として解釈される）。
 */
import { VBARunner, assert } from '../../test-libs/test-runner';

// --- 1. 最も基本的な単一代入文（Dim なし、暗黙宣言） ---
{
    const r = new VBARunner();
    r.eval('x = 10');
    assert.strictEqual(r.eval('x'), 10, 'eval("x = 10") 単体で代入が実行される');
}
console.log('[PASS] eval("x = 10") 単体で正しく代入される');

// --- 2. 配列要素への代入を別々の eval() 呼び出しで行う ---
{
    const r = new VBARunner();
    r.eval('Dim parts(1 To 3) As String');
    r.eval('parts(1) = "a"');
    r.eval('parts(2) = "b"');
    r.eval('parts(3) = "c"');
    assert.strictEqual(r.eval('Join(parts, "; ")'), 'a; b; c', '配列要素への代入が別々の eval() 呼び出しでも永続化される');
}
console.log('[PASS] 配列要素への代入が別々の eval() 呼び出しでも永続化される');

// --- 3. Debug.Print 経由での比較式評価は引き続き正しく動作する（回帰防止） ---
{
    const printed: string[] = [];
    const r = new VBARunner(null, { onPrint: (s: string) => printed.push(s) });
    r.eval('Dim z : z = 5');
    r.eval('Debug.Print z = 5');
    assert.deepStrictEqual(printed, ['True'], 'Debug.Print 経由の等価比較は引き続き機能する');
}
console.log('[PASS] Debug.Print 経由の等価比較は引き続き正しく動作する（既存挙動の維持）');

console.log('\n✅ eval() 裸の代入文: 全テスト通過');
