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

// Test 15: クラスPublic FunctionのPrivate UDT引数は静的に拒否する
{
    const ev = evalVBA(String.raw`Class ParameterHolder
Private Type HiddenRecord
    Value As Long
End Type
Public Function Expose(ByVal item As HiddenRecord) As Long
    Expose = item.Value
End Function
End Class
Public Function RunInvalidClassParameter() As Long
    Dim holder As New ParameterHolder
    RunInvalidClassParameter = holder.Expose(Nothing)
End Function`);
    let threw = false;
    try {
        ev.callProcedure('RunInvalidClassParameter', []);
    } catch (error: any) {
        threw = true;
        assert.ok(/Public class procedure.*Private UDT|Compile error/.test(String(error?.message)),
            'class Public Function Private UDT parameter is rejected');
    }
    assert.ok(threw, 'class Public Function Private UDT parameter is rejected');
    console.log('[PASS] Class Public Function Private UDT parameter is rejected');
}

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

// Test 9: クラス内Private Typeをクラス手続きから解決する
{
    const ev = evalVBA(String.raw`Class Holder
Option Explicit
Private Type Hidden
    Value As Long
End Type
Public Function ReadValue() As Long
    Dim value As Hidden
    value.Value = 42
    ReadValue = value.Value
End Function
End Class
Public Function RunClassUdt() As Long
    Dim holder As New Holder
    RunClassUdt = holder.ReadValue()
End Function`);
    assert.strictEqual(ev.callProcedure('RunClassUdt', []), 42,
        'Class-local Private Type resolves in class procedure');
    console.log('[PASS] Class-local Private Type resolves in class procedure');
}

// Test 10: クラスPublic FunctionのPrivate UDT戻り値は静的に拒否する
{
    const ev = evalVBA(String.raw`Class Holder
Private Type Hidden
    Value As Long
End Type
Public Function MakeHidden() As Hidden
    Dim value As Hidden
    value.Value = 42
    MakeHidden = value
End Function
End Class
Public Function RunInvalidClassReturn() As Long
    Dim holder As New Holder
    RunInvalidClassReturn = holder.MakeHidden().Value
End Function`);
    let threw = false;
    try {
        ev.callProcedure('RunInvalidClassReturn', []);
    } catch (error: any) {
        threw = true;
        assert.ok(/Public class function.*Private UDT|Compile error/.test(String(error?.message)),
            'class Public Function Private UDT return is rejected');
    }
    assert.ok(threw, 'class Public Function Private UDT return is rejected');
    console.log('[PASS] Class Public Function Private UDT return is rejected');
}

// Test 11: クラス内Private UDTのProperty Get/Let経路
{
    const ev = evalVBA(String.raw`Class ClassCase
Private Type TItem
    Value As Long
End Type
Private mItem As TItem
Private Property Get Item() As TItem
    Item = mItem
End Property
Private Property Let Item(ByVal value As TItem)
    mItem = value
End Property
Public Function Probe() As Long
    Dim localItem As TItem
    localItem.Value = 41
    Item = localItem
    Probe = Item.Value + 1
End Function
End Class
Public Function RunClassPropertyUdt() As Long
    Dim instance As New ClassCase
    RunClassPropertyUdt = instance.Probe
End Function`);
    assert.strictEqual(ev.callProcedure('RunClassPropertyUdt', []), 42,
        'Class-local Private UDT Property Get/Let works');
    console.log('[PASS] Class-local Private UDT Property Get/Let works');
}

// Test 12: クラス内Private UDT固定配列の要素初期化
{
    const ev = evalVBA(String.raw`Class ArrayHolder
Private Type TItem
    Value As Long
End Type
Private items(0 To 1) As TItem
Public Function Probe() As Long
    items(0).Value = 7
    Probe = items(0).Value
End Function
End Class
Public Function RunClassArrayUdt() As Long
    Dim instance As New ArrayHolder
    RunClassArrayUdt = instance.Probe
End Function`);
    assert.strictEqual(ev.callProcedure('RunClassArrayUdt', []), 7,
        'Class-local Private UDT fixed array works');
    console.log('[PASS] Class-local Private UDT fixed array works');
}

// Test 13: クラス内Private UDT動的配列のReDim
{
    const ev = evalVBA(String.raw`Class DynamicHolder
Private Type TItem
    Value As Long
End Type
Private items() As TItem
Public Function Probe() As Long
    ReDim items(0 To 1)
    items(0).Value = 21
    Probe = items(0).Value * 2
End Function
End Class
Public Function RunClassDynamicUdt() As Long
    Dim instance As New DynamicHolder
    RunClassDynamicUdt = instance.Probe
End Function`);
    assert.strictEqual(ev.callProcedure('RunClassDynamicUdt', []), 42,
        'Class-local Private UDT dynamic array ReDim works');
    console.log('[PASS] Class-local Private UDT dynamic array ReDim works');
}

// Test 14: クラス内Private UDT配列のPreserve/Eraseライフサイクル
{
    const ev = evalVBA(String.raw`Class LifecycleHolder
Private Type TItem
    Value As Long
End Type
Private items() As TItem
Public Function Probe() As String
    Dim s As String
    On Error Resume Next
    ReDim items(1 To 2)
    s = "redim=" & Err.Number
    Err.Clear: items(1).Value = 11: s = s & ";set1=" & Err.Number
    Err.Clear: ReDim Preserve items(1 To 3): s = s & ";preserve=" & Err.Number
    Err.Clear: Erase items: s = s & ";erase=" & Err.Number
    Err.Clear: ReDim items(1 To 1): s = s & ";rebuild-redim=" & Err.Number
    Err.Clear: items(1).Value = 99: s = s & ";rebuild-set=" & Err.Number
    Probe = s
End Function
End Class
Public Function RunClassLifecycleUdt() As String
    Dim instance As New LifecycleHolder
    RunClassLifecycleUdt = instance.Probe
End Function`);
    assert.strictEqual(ev.callProcedure('RunClassLifecycleUdt', []),
        'redim=0;set1=0;preserve=0;erase=0;rebuild-redim=0;rebuild-set=0',
        'Class-local Private UDT array lifecycle works');
console.log('[PASS] Class-local Private UDT array lifecycle works');
}

// Test 16: Public UDTの修飾名と裸名はByRef配列で同一視するが、所有者の異なる同名型は拒否する
{
    const ev = evalVBAModules([
        {
            name: 'Producer',
            code: String.raw`Option Explicit
Public Type RecordT
    Id As Long
End Type
Public Sub Sum(ByRef values() As Producer.RecordT)
    values(0).Id = values(0).Id + 1
End Sub`,
        },
        {
            name: 'Other',
            code: String.raw`Option Explicit
Public Type RecordT
    Id As Long
End Type`,
        },
        {
            name: 'Consumer',
            code: String.raw`Option Explicit
Public Function QualifiedArrayByRef() As Long
    Dim values() As RecordT
    ReDim values(0 To 0)
    values(0).Id = 41
    Producer.Sum values
    QualifiedArrayByRef = values(0).Id
End Function
Public Function WrongOwnerArrayByRef() As Long
    Dim values() As Other.RecordT
    ReDim values(0 To 0)
    Producer.Sum values
    WrongOwnerArrayByRef = values(0).Id
End Function`,
        },
    ]);
    assert.strictEqual(ev.callProcedure('QualifiedArrayByRef', []), 42,
        'qualified and bare Public UDT array names share ByRef identity');
    assert.throws(() => ev.callProcedure('WrongOwnerArrayByRef', []), /Type mismatch/,
        'same-named UDTs from different modules remain distinct');
    console.log('[PASS] Qualified/bare Public UDT array identity and owner separation');
}
