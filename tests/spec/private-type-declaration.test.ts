import { evalVBASingle, evalVBAModules, assert } from '../../test-libs/test-runner';

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

// Test 6: Private Type は宣言元モジュールの外から非修飾参照できない
{
    let threw = false;
    try {
        const ev = evalVBAModules([
            {
                name: 'Producer',
                code: String.raw`Option Explicit
Private Type T
    X As Long
End Type
Public Sub Mutate(ByRef value As T)
    value.X = 9
End Sub`,
            },
            {
                name: 'Caller',
                code: String.raw`Option Explicit
Public Function RunProbe() As Long
    Dim value As T
    value.X = 3
    Producer.Mutate value
    RunProbe = value.X
End Function`,
            },
        ]);
        ev.callProcedure('RunProbe', []);
    } catch {
        threw = true;
    }
    assert.ok(threw, 'Private Type: cross-module bare reference is rejected');
    console.log('[PASS] Private Type: cross-module bare reference is rejected');
}

// Test 7: Public Type は別モジュールから Module.Type で解決できる
{
    const ev = evalVBAModules([
        {
            name: 'ModuleA',
            code: String.raw`Option Explicit
Public Type T
    X As Long
End Type`,
        },
        {
            name: 'ModuleB',
            code: String.raw`Option Explicit
Public Function UseQualifiedType() As Long
    Dim value As ModuleA.T
    value.X = 9
    UseQualifiedType = value.X
End Function`,
        },
    ]);
    assert.strictEqual(ev.callProcedure('UseQualifiedType', []), 9,
        'Public Type: qualified cross-module reference works');
console.log('[PASS] Public Type: qualified cross-module reference works');
}

// Test 8: 標準モジュールのPublic FunctionはPrivate UDTを戻り値にできる
{
    const ev = evalVBAModules([
        {
            name: 'ModuleA',
            code: String.raw`Option Explicit
Private Type T
    X As Long
End Type
Public Function Make() As T
    Dim value As T
    value.X = 7
    Make = value
End Function`,
        },
    ]);
    const value = ev.callProcedure('ModuleA.Make', []) as { x: number };
    assert.strictEqual(value.x, 7,
        'Standard module Public Function may return a Private UDT');
    console.log('[PASS] Standard module Public Function may return a Private UDT');
}
