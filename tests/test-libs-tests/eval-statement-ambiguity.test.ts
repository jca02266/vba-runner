/**
 * VBARunner.eval() に渡した裸の文（`x + 1`、`Helper + 1` 等）が、単独で渡した場合と
 * 複数文の最後に置いた場合で異なる解釈をされてしまうバグのテスト
 *
 * 仕様バグ修正: VBA の statement 文法では、識別子（や `obj.Member`・`arr(i)` などの
 * 呼び出し可能な形）で始まり `=`/`+`/`-` が続く裸の文は、単独行でも複数文中でも
 * 常に同一の解釈（代入文 or 暗黙の Call 文）になる（`<`/`&`/`And` 等の比較・連結・
 * 論理演算子はこの曖昧性を持たず、statement としては Parse error になるため対象外）。
 * しかし evalExpression() の「単一式として全体を消費したか」高速パスが `=` のケースしか
 * 見ておらず、`+`/`-` で識別子・呼び出し可能な形が左辺（再帰的に最も左の葉まで）に
 * 来るケースを見逃していたため、単独 eval() では誤って算術式として評価されていた
 * （変数なら誤った加算結果、必須引数を持つ Function/Sub なら暗黙的に Empty(0) 扱いされ
 * 呼び出しが発生しないまま `0 + 1` 等の誤った値を返していた）。
 */
import { VBARunner, assert } from '../../test-libs/test-runner';

// --- 1. 変数: 単独 eval() でも複数文中と同じく呼び出し不可能エラーになる ---
{
    const single = new VBARunner();
    single.eval('Dim x : x = 5');
    let singleThrew = false;
    try { single.eval('x + 1'); } catch { singleThrew = true; }

    const chained = new VBARunner();
    let chainedThrew = false;
    try { chained.eval('Dim x : x = 5 : x + 1'); } catch { chainedThrew = true; }

    assert.ok(singleThrew && chainedThrew, '変数の x + 1 は単独でも複数文中でも Error 424 になる（実VBA仕様: 裸の文は常にCall文として解釈される）');
}
console.log('[PASS] 変数: 単独 eval() でも複数文中と同じ結果（呼び出し不可能エラー）になる');

// --- 2. Function（必須引数1個）: 単独でも複数文中でも実際に呼び出される ---
{
    const code = 'Function Helper(n) As Long\n    Helper = n * 100\nEnd Function';

    const single = new VBARunner();
    single.eval(code);
    const singleResult = single.eval('Helper + 1');

    const chained = new VBARunner();
    chained.eval(code);
    const chainedResult = chained.eval('Dim d : d = 0 : Helper + 1');

    assert.strictEqual(singleResult, undefined, '単独 eval() でも Helper + 1 は Call 文として実行され、戻り値（文として）は捨てられる');
    assert.strictEqual(chainedResult, undefined, '複数文中の Helper + 1 と同じ結果');
}
console.log('[PASS] Function: 単独 eval() でも複数文中と同じく Helper(1) が呼び出される');

// --- 3. Sub（必須引数1個）: 単独でも複数文中でも実際に呼び出される ---
{
    const code = 'Sub Helper(n)\n    Debug.Print "called with " & n\nEnd Sub';

    const printed1: string[] = [];
    const single = new VBARunner(null, { onPrint: (s: string) => printed1.push(s) });
    single.eval(code);
    single.eval('Helper + 1');

    const printed2: string[] = [];
    const chained = new VBARunner(null, { onPrint: (s: string) => printed2.push(s) });
    chained.eval(code);
    chained.eval('Dim d : d = 0 : Helper + 1');

    assert.deepStrictEqual(printed1, ['called with 1'], '単独 eval() でも Helper(1) が実際に呼び出される');
    assert.deepStrictEqual(printed2, ['called with 1'], '複数文中の Helper + 1 と同じ結果');
}
console.log('[PASS] Sub: 単独 eval() でも複数文中と同じく Helper(1) が呼び出される');

// --- 4. 既存の正しい挙動が壊れていないことの確認（回帰防止） ---
{
    const v = new VBARunner();
    assert.strictEqual(v.eval('1 + 2 + 3'), 6, 'リテラルのみの式は引き続き算術評価される');
}
{
    const v = new VBARunner();
    v.eval('Dim x : x = 5');
    assert.strictEqual(v.eval('(x) + 1'), 6, '括弧で明示的に囲んだ場合は曖昧性が解消され算術式として評価される');
}
{
    const v = new VBARunner();
    v.eval('Dim parts(1 To 3) As String');
    v.eval('parts(1) = "a"');
    assert.strictEqual(v.eval('parts(1)'), 'a', '配列要素への代入は別々の eval() 呼び出しでも引き続き永続化される');
}
console.log('[PASS] 既存の正しい挙動（リテラル式・括弧での曖昧性解消・配列要素代入）は維持される');

console.log('\n✅ eval() statement/expression 曖昧性の一貫性: 全テスト通過');
