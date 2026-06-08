import { evalVBASingle, assert } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    return evalVBASingle(code);
}

function runFunc(code: string, name: string, args: any[] = []): any {
    return evalVBA(code).callProcedure(name, args);
}

// Test 1: Private Type が AST に含まれること（メンバアクセスが成功すること）
{
    const result = runFunc(`
Private Type Point
    X As Long
    Y As Long
End Type

Function F() As String
    Dim p As Point
    p.X = 10
    p.Y = 20
    F = p.X & "," & p.Y
End Function
`, 'F');
    assert.strictEqual(result, '10,20', 'Private Type: member access works');
    console.log('[PASS] Private Type: member access works');
}

// Test 2: Public Type が AST に含まれること
{
    const result = runFunc(`
Public Type Color
    R As Integer
    G As Integer
    B As Integer
End Type

Function F() As Long
    Dim c As Color
    c.R = 255
    c.G = 128
    c.B = 0
    F = c.R + c.G + c.B
End Function
`, 'F');
    assert.strictEqual(result, 383, 'Public Type: member access works');
    console.log('[PASS] Public Type: member access works');
}

// Test 3: Type 内の配列メンバ（バグ再現: utc_StandardName(0 To 31) As Integer 形式）
{
    const result = runFunc(`
Private Type DataBlock
    Values(0 To 3) As Integer
    Name As String
End Type

Function F() As String
    Dim d As DataBlock
    d.Name = "test"
    F = d.Name
End Function
`, 'F');
    assert.strictEqual(result, 'test', 'Type with array member: non-array member accessible');
    console.log('[PASS] Type with array member: non-array member accessible');
}

// Test 4: Nested UDT（型の中に型）
{
    const result = runFunc(`
Private Type Inner
    Val As Long
End Type

Private Type Outer
    Inner1 As Inner
    Label As String
End Type

Function F() As String
    Dim o As Outer
    o.Inner1.Val = 42
    o.Label = "ok"
    F = o.Label & ":" & o.Inner1.Val
End Function
`, 'F');
    assert.strictEqual(result, 'ok:42', 'Nested UDT: inner member access');
    console.log('[PASS] Nested UDT: inner member access');
}

// Test 5: Boolean メンバは初期値 0 (VBA False)
{
    const result = runFunc(`
Private Type Options
    UseDouble As Boolean
    AllowUnquoted As Boolean
End Type

Function F() As Long
    Dim o As Options
    If o.UseDouble Then
        F = 1
    Else
        F = 0
    End If
End Function
`, 'F');
    assert.strictEqual(result, 0, 'Type Boolean member initializes to False (0)');
    console.log('[PASS] Type Boolean member initializes to False (0)');
}

console.log('\n✅ private-type-declaration: 全テスト通過');
