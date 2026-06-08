/**
 * 関数戻り値への Let-coercion テスト
 *
 * Function F() As T に対して F = <異なる型の値> を代入したとき、
 * VBA 仕様通りに型変換・丸め・Type Mismatch が発生することを確認する。
 */
import { evalVBASingle, assert } from '../../test-libs/test-runner';

function run(code: string): string {
    let out = '';
    evalVBASingle(code, { onPrint: (s) => { out += s + '\n'; } });
    return out.trim();
}

function runExpectError(code: string): string {
    try { return run(code); }
    catch (e: any) { return `ERROR:${e.number ?? e.message ?? e}`; }
}

// ─── Long 戻り型 ──────────────────────────────────────────────────────────────

// 数値文字列 → Long に変換される
{
    const r = run(`
Function F() As Long
    F = "123"
End Function
Debug.Print F()
`);
    assert.strictEqual(r, '123', 'Long ← "123": 変換後 123');
    assert.strictEqual(typeof +r === 'number', true, 'Long ← "123": 数値として返る');
    console.log('[PASS] Long ← "123"');
}

// 非数値文字列 → Type Mismatch (Error 13)
{
    const r = runExpectError(`
Function F() As Long
    F = "hello"
End Function
Debug.Print F()
`);
    assert.strictEqual(r.startsWith('ERROR:'), true, 'Long ← "hello": Type Mismatch エラー');
    assert.strictEqual(r.includes('13'), true, 'Long ← "hello": Error 13');
    console.log('[PASS] Long ← "hello" → Type Mismatch');
}

// Double → Long 丸め（銀行丸め: 3.5 → 4, 2.5 → 2）
{
    const r = run(`
Function F() As Long
    F = 3.7
End Function
Debug.Print F()
`);
    assert.strictEqual(r, '4', 'Long ← 3.7: 丸めて 4');
    console.log('[PASS] Long ← 3.7 → 4（丸め）');
}

{
    const r = run(`
Function F() As Long
    F = 2.5
End Function
Debug.Print F()
`);
    assert.strictEqual(r, '2', 'Long ← 2.5: 銀行丸めで 2');
    console.log('[PASS] Long ← 2.5 → 2（銀行丸め）');
}

// ─── String 戻り型 ────────────────────────────────────────────────────────────

// Long → String に変換される
{
    const r = run(`
Function F() As String
    F = 42
End Function
Debug.Print F()
`);
    assert.strictEqual(r, '42', 'String ← 42: "42" に変換');
    console.log('[PASS] String ← 42 → "42"');
}

// Boolean → String
{
    const r = run(`
Function F() As String
    F = True
End Function
Debug.Print F()
`);
    assert.strictEqual(r, 'True', 'String ← True: "True" に変換');
    console.log('[PASS] String ← True → "True"');
}

// ─── Boolean 戻り型 ───────────────────────────────────────────────────────────

// 非ゼロ数値 → True (-1)
{
    const r = run(`
Function F() As Boolean
    F = 5
End Function
Debug.Print F()
`);
    assert.strictEqual(r, 'True', 'Boolean ← 5: True');
    console.log('[PASS] Boolean ← 5 → True');
}

// 0 → False
{
    const r = run(`
Function F() As Boolean
    F = 0
End Function
Debug.Print F()
`);
    assert.strictEqual(r, 'False', 'Boolean ← 0: False');
    console.log('[PASS] Boolean ← 0 → False');
}

// ─── Integer 戻り型 ───────────────────────────────────────────────────────────

// Double → Integer 丸め
{
    const r = run(`
Function F() As Integer
    F = 1.9
End Function
Debug.Print F()
`);
    assert.strictEqual(r, '2', 'Integer ← 1.9: 2');
    console.log('[PASS] Integer ← 1.9 → 2');
}

// Integer オーバーフロー (>32767) → Error 6
{
    const r = runExpectError(`
Function F() As Integer
    F = 40000
End Function
Debug.Print F()
`);
    assert.strictEqual(r.startsWith('ERROR:'), true, 'Integer ← 40000: Overflow');
    assert.strictEqual(r.includes('6'), true, 'Integer ← 40000: Error 6');
    console.log('[PASS] Integer ← 40000 → Overflow');
}

// ─── 呼び出し側の型と値 ──────────────────────────────────────────────────────

// Long 関数の戻り値は丸め済みの整数値として呼び出し側に届く
{
    const r = run(`
Function F() As Long
    F = 3.7
End Function
Dim v As Variant
v = F()
Debug.Print v
`);
    assert.strictEqual(r, '4', '呼び出し側: 丸め済みの値 4 が届く');
    console.log('[PASS] 呼び出し側に丸め済み値 4 が届く');
}

// TypeName(F()) は宣言型を返す
{
    const r = run(`
Function F() As Long
    F = 3.7
End Function
Debug.Print TypeName(F())
`);
    assert.strictEqual(r, 'Long', 'TypeName(F()) = Long');
    console.log('[PASS] TypeName(F()) = Long');
}

// TypeName(CLng(1)) = Long
{
    const r = run('Debug.Print TypeName(CLng(1))');
    assert.strictEqual(r, 'Long', 'TypeName(CLng(1)) = Long');
    console.log('[PASS] TypeName(CLng(1)) = Long');
}

console.log('\n✅ 関数戻り値 Let-coercion: 全テスト通過');
