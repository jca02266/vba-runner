/**
 * `VBA.Func()` 修飾呼び出しが標準ライブラリを強制するテスト
 *
 * VBA 仕様: `VBA.InStr(...)` のような型ライブラリ修飾呼び出しは、
 * ユーザーが同名の関数を定義していても必ず標準ライブラリを呼ぶ。
 * （非修飾呼び出しはユーザー定義優先のまま）
 */
import { evalVBAModules, assert } from '../../test-libs/test-runner';

function run(modules: Array<{ name: string; code: string }>, entry: string): any {
    const ev = evalVBAModules(modules);
    return ev.callProcedure(entry, []);
}

// ---------------------------------------------------------------------------
// VBA.InStr: ユーザー定義 InStr が存在しても標準ライブラリを呼ぶ
// ---------------------------------------------------------------------------
{
    const result = run([
        { name: 'UserDef', code: `
Public Function InStr(s As String, tok As String) As Long
    InStr = 999
End Function` },
        { name: 'Main', code: `
Function TestUnqualified() As Long
    TestUnqualified = InStr("hello", "l")
End Function
Function TestVbaQualified() As Long
    TestVbaQualified = VBA.InStr("hello", "l")
End Function` }
    ], 'TestUnqualified');
    assert.strictEqual(result, 999, '非修飾 InStr はユーザー定義優先 → 999');
    console.log('[PASS] 非修飾 InStr → ユーザー定義 (999)');
}
{
    const result = run([
        { name: 'UserDef', code: `
Public Function InStr(s As String, tok As String) As Long
    InStr = 999
End Function` },
        { name: 'Main', code: `
Function TestVbaQualified() As Long
    TestVbaQualified = VBA.InStr("hello", "l")
End Function` }
    ], 'TestVbaQualified');
    assert.strictEqual(result, 3, 'VBA.InStr は標準ライブラリ → 3');
    console.log('[PASS] VBA.InStr → 標準ライブラリ (3)');
}

// ---------------------------------------------------------------------------
// VBA.Len: 別の組み込み関数でも同様
// ---------------------------------------------------------------------------
{
    const result = run([
        { name: 'UserDef', code: `
Public Function Len(s As String) As Long
    Len = 999
End Function` },
        { name: 'Main', code: `
Function TestVbaLen() As Long
    TestVbaLen = VBA.Len("hello")
End Function` }
    ], 'TestVbaLen');
    assert.strictEqual(result, 5, 'VBA.Len は標準ライブラリ → 5');
    console.log('[PASS] VBA.Len → 標準ライブラリ (5)');
}

// ---------------------------------------------------------------------------
// ユーザー定義がない場合も正常動作
// ---------------------------------------------------------------------------
{
    const result = run([
        { name: 'Main', code: `
Function TestNoConflict() As Long
    TestNoConflict = VBA.InStr("hello world", "world")
End Function` }
    ], 'TestNoConflict');
    assert.strictEqual(result, 7, 'VBA.InStr（ユーザー定義なし）→ 7');
    console.log('[PASS] VBA.InStr（競合なし）→ 標準ライブラリ (7)');
}

console.log('\n✅ vba-stdlib-qualifier: 全テスト通過');
