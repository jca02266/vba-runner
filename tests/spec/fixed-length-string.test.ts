import { evalVBASingle, evalVBAModules, assert } from '../../test-libs/test-runner';

function runFunc(code: string, name: string, args: any[] = []): any {
    return evalVBASingle(code).callProcedure(name, args);
}

// ── 基本動作 ──────────────────────────────────────────────────────────────────

{
    // 初期値は NUL 文字 × N
    const ev = evalVBASingle(`
Function GetInit() As Long
  Dim s As String * 5
  GetInit = Len(s)
End Function
`);
    assert.strictEqual(ev.callProcedure('GetInit', []), 5, 'initial Len is N');
    console.log('[PASS] 初期値 Len=N');
}

{
    // 短い文字列はスペースでパディング
    const ev = evalVBASingle(`
Function GetPad() As String
  Dim s As String * 8
  s = "Hello"
  GetPad = s
End Function
`);
    assert.strictEqual(ev.callProcedure('GetPad', []), 'Hello   ', 'short string padded');
    console.log('[PASS] 短い文字列はスペースパディング');
}

{
    // 長い文字列は右から切り捨て
    const ev = evalVBASingle(`
Function GetTrunc() As String
  Dim s As String * 5
  s = "Hello World"
  GetTrunc = s
End Function
`);
    assert.strictEqual(ev.callProcedure('GetTrunc', []), 'Hello', 'long string truncated');
    console.log('[PASS] 長い文字列は切り捨て');
}

{
    // 空文字列代入はすべてスペース
    const ev = evalVBASingle(`
Function GetEmpty() As String
  Dim s As String * 4
  s = ""
  GetEmpty = s
End Function
`);
    assert.strictEqual(ev.callProcedure('GetEmpty', []), '    ', 'empty string becomes all spaces');
    console.log('[PASS] 空文字列代入はスペース × N');
}

{
    // ちょうど N 文字の代入はそのまま
    const ev = evalVBASingle(`
Function GetExact() As String
  Dim s As String * 5
  s = "ABCDE"
  GetExact = s
End Function
`);
    assert.strictEqual(ev.callProcedure('GetExact', []), 'ABCDE', 'exact-length string unchanged');
    console.log('[PASS] ちょうど N 文字はそのまま');
}

// ── Len() ─────────────────────────────────────────────────────────────────────

{
    // Len() は常に N を返す
    const ev = evalVBASingle(`
Function TestLen() As String
  Dim s As String * 10
  s = "Hi"
  TestLen = CStr(Len(s))
End Function
`);
    assert.strictEqual(ev.callProcedure('TestLen', []), '10', 'Len always returns N');
    console.log('[PASS] Len() は常に N を返す');
}

// ── 複数宣言・同一行 ──────────────────────────────────────────────────────────

{
    // 複数の固定長変数を同じ行で宣言
    const ev = evalVBASingle(`
Function GetMulti() As String
  Dim a As String * 3, b As String * 6
  a = "ABCDE"
  b = "Hi"
  GetMulti = a & "|" & b
End Function
`);
    assert.strictEqual(ev.callProcedure('GetMulti', []), 'ABC|Hi    ', 'multiple fixed-length in same Dim');
    console.log('[PASS] 同一 Dim 行での複数固定長変数');
}

// ── UDT メンバー ──────────────────────────────────────────────────────────────

{
    // UDT の String * N メンバーが正しく動作する
    const ev = evalVBASingle(`
Type Person
  Name As String * 10
  Code As String * 4
End Type

Function TestUDT() As String
  Dim p As Person
  p.Name = "Alice"
  p.Code = "ABCDEF"
  TestUDT = p.Name & "|" & p.Code
End Function
`);
    assert.strictEqual(ev.callProcedure('TestUDT', []), 'Alice     |ABCD', 'UDT fixed-length member coercion');
    console.log('[PASS] UDT の String * N メンバー');
}

{
    // UDT の初期値: 固定長メンバーは NUL × N
    const ev = evalVBASingle(`
Type Rec
  Field As String * 5
End Type

Function TestUDTInit() As Long
  Dim r As Rec
  TestUDTInit = Len(r.Field)
End Function
`);
    assert.strictEqual(ev.callProcedure('TestUDTInit', []), 5, 'UDT fixed-length member initial Len');
    console.log('[PASS] UDT 固定長メンバーの初期 Len=N');
}

// ── モジュールレベル変数 ──────────────────────────────────────────────────────

{
    const ev = evalVBASingle(`
Dim gStr As String * 6

Sub SetGlobal()
  gStr = "Hi"
End Sub

Function GetGlobal() As String
  GetGlobal = gStr
End Function
`);
    ev.callProcedure('SetGlobal', []);
    assert.strictEqual(ev.callProcedure('GetGlobal', []), 'Hi    ', 'module-level fixed-length string');
    console.log('[PASS] モジュールレベル固定長文字列');
}

// ── Mid() との組み合わせ ──────────────────────────────────────────────────────

{
    // Mid() で部分取り出し
    const ev = evalVBASingle(`
Function TestMid() As String
  Dim s As String * 10
  s = "ABCDE"
  TestMid = Mid(s, 1, 3)
End Function
`);
    assert.strictEqual(ev.callProcedure('TestMid', []), 'ABC', 'Mid() on fixed-length string');
    console.log('[PASS] Mid() による部分取り出し');
}

// ── Trim() ───────────────────────────────────────────────────────────────────

{
    // Trim() でパディングを除去できる
    const ev = evalVBASingle(`
Function TestTrim() As String
  Dim s As String * 10
  s = "Hello"
  TestTrim = Trim(s)
End Function
`);
    assert.strictEqual(ev.callProcedure('TestTrim', []), 'Hello', 'Trim() removes padding spaces');
    console.log('[PASS] Trim() でパディングを除去');
}

console.log('\n✅ 固定長文字列（String * N）: 全テスト通過');
