/**
 * VBARunner.eval() で定義したプロシージャも Option Explicit の静的検査対象になることのテスト
 *
 * 仕様バグ修正: resolveIdentifiers()（Pass 2）は VBARunner 構築時にロードしたモジュール
 * 群しか Option Explicit の静的解析対象にしないため、`new VBARunner()` の後に eval() で
 * 定義した Option Explicit 付きプロシージャは一度も解析されず、未宣言変数の使用が
 * 黒く通っていた。evalExpression() の複数文フォールバック内で都度 checkOptionExplicit
 * を実行し optionExplicitViolations に追加登録するよう修正した。
 */
import { VBARunner, assert } from '../../test-libs/test-runner';

// --- 1. eval() で定義した Option Explicit 付き関数の未宣言変数使用がコンパイルエラーになる ---
{
    const r = new VBARunner();
    r.eval(`
Option Explicit
Public Function ParseTicket(t As String) As String
    Dim z As String
    z = totallyUndeclaredVar
    ParseTicket = "never"
End Function
`);
    let threw = false;
    try {
        r.eval('ParseTicket("x")');
    } catch (e: any) {
        threw = true;
        assert.ok(/Option Explicit/.test(e.message), `Option Explicit 違反のエラーメッセージになる: ${e.message}`);
    }
    assert.ok(threw, 'eval() で定義した関数でも Option Explicit 違反がコンパイルエラーになる');
}
console.log('[PASS] eval() で定義した関数も Option Explicit の静的検査対象になる');

// --- 2. Option Explicit がない通常の eval() コードは影響を受けない ---
{
    const r = new VBARunner();
    r.eval('Public Function NoOptionExplicit() As Long\n    NoOptionExplicit = anyUndeclaredVar + 1\nEnd Function');
    const v = r.eval('NoOptionExplicit()');
    assert.strictEqual(v, 1, 'Option Explicit なしの場合は未宣言変数が暗黙の0として動作する（既存挙動を維持）');
}
console.log('[PASS] Option Explicit がない場合は既存の暗黙宣言の挙動を維持する');

console.log('\n✅ VBARunner.eval() + Option Explicit: 全テスト通過');
