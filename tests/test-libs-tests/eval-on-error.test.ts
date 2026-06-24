/**
 * VBARunner.eval() が On Error Resume Next / On Error GoTo を正しく処理することのテスト
 *
 * evalExpression() の複数文フォールバックは、On Error の分岐ロジック（Resume Next /
 * GoTo / Resume）を持つ executeStatements() を経由していなかったため、eval() 内で
 * On Error を使ってもエラーがそのまま伝播してしまっていた。
 * Sub/Function 内（run() 経由）では正しく動いていたのと同じ挙動を eval() でも
 * 再現できることを検証する。
 */
import { VBARunner, assert } from '../../test-libs/test-runner';

// --- 1. eval() で On Error Resume Next が効く ---
{
    const printed: string[] = [];
    const r = new VBARunner(null, { onPrint: (s: string) => printed.push(s) });
    r.eval('On Error Resume Next: Err.Raise 1001, "S", "msg": Dim x As String: x = "after": Debug.Print x');
    assert.deepStrictEqual(printed, ['after'], 'eval() 内の Err.Raise が Resume Next で継続する');
}
console.log('[PASS] eval() で On Error Resume Next が効く');

// --- 2. eval() で On Error GoTo Label が効く ---
{
    const printed: string[] = [];
    const r = new VBARunner(null, { onPrint: (s: string) => printed.push(s) });
    r.eval('On Error GoTo Handler: Err.Raise 1001, "S", "boom": Dim r: r = "unreached": GoTo Done: Handler: r = "handled": Done: Debug.Print r');
    assert.deepStrictEqual(printed, ['handled'], 'eval() 内の Err.Raise が GoTo Handler で捕捉される');
}
console.log('[PASS] eval() で On Error GoTo Label が効く');

// --- 3. eval() で設定した On Error 状態が次の eval() 呼び出しに漏れない ---
{
    const r = new VBARunner();
    r.eval('On Error Resume Next');
    let threw = false;
    try {
        r.eval('Err.Raise 5000, "S", "should not be swallowed"');
    } catch (e: any) {
        threw = true;
        assert.ok(e.message.includes('should not be swallowed'), 'エラーメッセージが伝播する');
    }
    assert.ok(threw, '直前の eval() の On Error Resume Next が次の eval() に漏れ残らない');
}
console.log('[PASS] eval() の On Error 状態は次の eval() 呼び出しに漏れない');

console.log('\n✅ VBARunner.eval() + On Error: 全テスト通過');
