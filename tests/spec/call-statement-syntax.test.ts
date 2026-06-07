/**
 * Call Statement 構文バリエーションテスト
 *
 * MS-VBAL §5.4.2.1 Call Statement:
 *   call-statement = "Call" (simple-name-expression / member-access-expression
 *                            / index-expression / with-expression)
 *   call-statement =/ (simple-name-expression / member-access-expression
 *                      / with-expression) argument-list
 *
 * argument-list = [positional-or-named-argument-list]  (省略可 = 空引数リスト)
 *
 * テスト観点:
 *   1. Call キーワードあり / なし
 *   2. 引数なし / あり（値・関数呼び出し）
 *   3. メンバーアクセス呼び出し
 *   4. With ブロック内呼び出し（.Method）
 *   5. 名前付き引数
 */

import { evalVBASingle, assert } from '../../test-libs/test-runner';

function run(code: string, name: string, args: unknown[] = []): unknown {
    return evalVBASingle(code).callProcedure(name, args);
}

function assertNoThrow(label: string, fn: () => void) {
    assert.doesNotThrow(fn, label);
}

// ─── 1. Call キーワードあり / なし ──────────────────────────────────────────

// CS01: Call キーワードありで引数なし（括弧あり）
{
    const code = `
Sub MySub()
End Sub
Sub Test()
    Call MySub()
End Sub
`;
    assertNoThrow('CS01', () => run(code, 'Test'));
    console.log('[PASS] CS01: Call MySub() — Call キーワードあり、引数なし（括弧あり）');
}

// CS02: Call キーワードありで括弧なし
{
    const code = `
Sub MySub()
End Sub
Sub Test()
    Call MySub
End Sub
`;
    assertNoThrow('CS02', () => run(code, 'Test'));
    console.log('[PASS] CS02: Call MySub — Call キーワードあり、括弧なし');
}

// CS03: Call キーワードなし、引数なし（空 argument-list）
{
    const code = `
Sub MySub()
End Sub
Sub Test()
    MySub
End Sub
`;
    assertNoThrow('CS03', () => run(code, 'Test'));
    console.log('[PASS] CS03: MySub — Call なし、引数なし（空 argument-list）');
}

// CS04: Call キーワードなし、引数あり（値リテラル）
{
    const code = `
Dim g As Integer
Sub MySub(n As Integer)
    g = n
End Sub
Sub Test()
    MySub 42
End Sub
`;
    assertNoThrow('CS04', () => run(code, 'Test'));
    console.log('[PASS] CS04: MySub 42 — Call なし、引数あり（リテラル）');
}

// CS05: Call キーワードなし、引数あり（複数）
{
    const code = `
Dim g As Integer
Sub AddNums(a As Integer, b As Integer)
    g = a + b
End Sub
Sub Test()
    AddNums 10, 20
End Sub
`;
    assertNoThrow('CS05', () => run(code, 'Test'));
    console.log('[PASS] CS05: AddNums 10, 20 — 複数引数');
}

// CS06: Call キーワードなし、引数が関数呼び出し式
// "手続き 関数()" パターン — 手続きの引数として関数の戻り値を渡す
{
    const code = `
Dim g As Integer
Function GetVal() As Integer
    GetVal = 99
End Function
Sub MySub(n As Integer)
    g = n
End Sub
Sub Test()
    MySub GetVal()
End Sub
`;
    assertNoThrow('CS06', () => run(code, 'Test'));
    const ev = evalVBASingle(code);
    ev.callProcedure('Test', []);
    assert.strictEqual(ev.env.get('g'), 99, 'CS06: GetVal() の戻り値が MySub に渡ること');
    console.log('[PASS] CS06: MySub GetVal() — 引数が関数呼び出し式（手続き 関数()パターン）');
}

// CS07: Call キーワードあり、index-expression（括弧付き呼び出し）
{
    const code = `
Dim g As Integer
Function AddNums(a As Integer, b As Integer) As Integer
    AddNums = a + b
End Function
Sub Test()
    Call AddNums(10, 20)
End Sub
`;
    assertNoThrow('CS07', () => run(code, 'Test'));
    console.log('[PASS] CS07: Call AddNums(10, 20) — Call + index-expression');
}

// ─── 2. メンバーアクセス呼び出し ────────────────────────────────────────────

// CS08: Call なし、メンバーアクセス、引数なし
{
    const code = `
Sub Test()
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.RemoveAll
End Sub
`;
    assertNoThrow('CS08', () => run(code, 'Test'));
    console.log('[PASS] CS08: d.RemoveAll — メンバーアクセス、引数なし');
}

// CS09: Call なし、メンバーアクセス、引数あり（複数）
{
    const code = `
Sub Test()
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "key", "value"
End Sub
`;
    assertNoThrow('CS09', () => run(code, 'Test'));
    console.log('[PASS] CS09: d.Add "key", "value" — メンバーアクセス、複数引数');
}

// CS10: Call あり、メンバーアクセス（括弧付き）
{
    const code = `
Sub Test()
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Call d.Add("key", "value")
End Sub
`;
    assertNoThrow('CS10', () => run(code, 'Test'));
    console.log('[PASS] CS10: Call d.Add("key", "value") — Call + メンバーアクセス');
}

// ─── 3. With ブロック内呼び出し（with-expression: .Method）──────────────────

// CS11: With ブロック内の .Method 呼び出し
{
    const code = `
Sub Test()
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    With d
        .Add "k1", 1
        .Add "k2", 2
    End With
End Sub
`;
    assertNoThrow('CS11', () => run(code, 'Test'));
    console.log('[PASS] CS11: With d ... .Add "k1", 1 — With ブロック内メンバー呼び出し');
}

// ─── 4. 名前付き引数 ────────────────────────────────────────────────────────

// CS12: 名前付き引数（順序を入れ替えて渡す）
{
    const code = `
Dim g As Integer
Sub MySub(a As Integer, b As Integer)
    g = a - b
End Sub
Sub Test()
    MySub b:=3, a:=10
End Sub
`;
    const ev = evalVBASingle(code);
    ev.callProcedure('Test', []);
    assert.strictEqual(ev.env.get('g'), 7, 'CS12: a=10, b=3 → g = 10-3 = 7');
    console.log('[PASS] CS12: MySub b:=3, a:=10 — 名前付き引数（順序逆転）');
}

// ─── 5. 返値を捨てる関数呼び出し ────────────────────────────────────────────

// CS13: 関数の戻り値を捨てる — GetVal（括弧なし・Call なし）
{
    const code = `
Function GetVal() As Integer
    GetVal = 42
End Function
Sub Test()
    GetVal
End Sub
`;
    assertNoThrow('CS13', () => run(code, 'Test'));
    console.log('[PASS] CS13: GetVal — 返値を捨てる（括弧なし、Call なし）');
}

// CS14: 関数の戻り値を捨てる — Call GetVal（括弧なし・Call あり）
{
    const code = `
Function GetVal() As Integer
    GetVal = 42
End Function
Sub Test()
    Call GetVal
End Sub
`;
    assertNoThrow('CS14', () => run(code, 'Test'));
    console.log('[PASS] CS14: Call GetVal — 返値を捨てる（括弧なし、Call あり）');
}

// CS15: 関数の戻り値を捨てる — Call GetVal()（括弧あり・Call あり）
{
    const code = `
Function GetVal() As Integer
    GetVal = 42
End Function
Sub Test()
    Call GetVal()
End Sub
`;
    assertNoThrow('CS15', () => run(code, 'Test'));
    console.log('[PASS] CS15: Call GetVal() — 返値を捨てる（括弧あり、Call あり）');
}

console.log('\n✅ Call Statement 構文バリエーション: 全テスト通過');
