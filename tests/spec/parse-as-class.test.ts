import { evalVBASingle, evalVBAModules, assert } from '../../test-libs/test-runner';

function evalWithClass(clsSource: string, className: string, moduleSource: string = '') {
    const modules: Array<{ name: string; code: string; parseAsClass?: string }> = [
        { name: className, code: clsSource, parseAsClass: className },
    ];
    if (moduleSource) modules.push({ name: 'Module', code: moduleSource });
    return evalVBAModules(modules);
}

// Test 1: parseAsClass produces the same result as explicit Class...End Class wrapping
{
    const clsSource = `
Public Value As Integer
Function Double() As Integer
    Double = Value * 2
End Function
`;
    const modSource = `
Function Test1() As Integer
    Dim obj As New MyClass
    obj.Value = 21
    Test1 = obj.Double()
End Function
`;
    // Via parseAsClass option
    const ev1 = evalWithClass(clsSource, 'MyClass', modSource);
    const r1 = ev1.callProcedure('Test1', []);

    // Via explicit Class...End Class syntax
    const ev2 = evalVBAModules([
        { name: 'MyClass', code: `Class MyClass\n${clsSource}\nEnd Class` },
        { name: 'Module', code: modSource },
    ]);
    const r2 = ev2.callProcedure('Test1', []);

    assert.strictEqual(r1, 42, 'parseAsClass: method call returns correct value');
    assert.strictEqual(r1, r2, 'parseAsClass produces same result as explicit Class...End Class');
    console.log('[PASS] parseAsClass equivalent to explicit wrapping');
}

// Test 2: Public field access works
{
    const ev = evalWithClass(`
Public Name As String
Public Age As Integer
`, 'Person', `
Function Test2() As String
    Dim p As New Person
    p.Name = "Alice"
    p.Age = 30
    Test2 = p.Name & "," & p.Age
End Function
`);
    assert.strictEqual(ev.callProcedure('Test2', []), 'Alice,30', 'Public fields accessible');
    console.log('[PASS] Public fields via parseAsClass');
}

// Test 3: Private fields with Property Get/Let
{
    const ev = evalWithClass(`
Private mValue As Integer

Property Get Value() As Integer
    Value = mValue
End Property

Property Let Value(v As Integer)
    If v >= 0 Then mValue = v
End Property
`, 'Counter', `
Function Test3() As Integer
    Dim c As New Counter
    c.Value = 5
    Test3 = c.Value
End Function
`);
    assert.strictEqual(ev.callProcedure('Test3', []), 5, 'Property Get/Let via parseAsClass');
    console.log('[PASS] Property Get/Let via parseAsClass');
}

// Test 4: Multiple classes loaded via parseAsClass
{
    const ev = evalVBAModules([
        { name: 'Dog', code: `
Public Name As String
Function Speak() As String
    Speak = Name & " says Woof"
End Function
`, parseAsClass: 'Dog' },
        { name: 'Cat', code: `
Public Name As String
Function Speak() As String
    Speak = Name & " says Meow"
End Function
`, parseAsClass: 'Cat' },
        { name: 'Module', code: `
Function Test4() As String
    Dim d As New Dog
    Dim c As New Cat
    d.Name = "Rex"
    c.Name = "Whiskers"
    Test4 = d.Speak() & "|" & c.Speak()
End Function
` },
    ]);
    assert.strictEqual(ev.callProcedure('Test4', []), 'Rex says Woof|Whiskers says Meow', 'Multiple classes via parseAsClass');
    console.log('[PASS] Multiple classes via parseAsClass');
}

// Test 5: Implements via parseAsClass
{
    const ev = evalVBAModules([
        { name: 'IAnimal', code: `
Class IAnimal
Function Speak() As String
End Function
End Class
` },
        { name: 'Cow', code: `
Implements IAnimal
Function IAnimal_Speak() As String
    IAnimal_Speak = "Moo"
End Function
`, parseAsClass: 'Cow' },
        { name: 'Module', code: `
Function Test5() As String
    Dim a As IAnimal
    Set a = New Cow
    Test5 = a.Speak()
End Function
` },
    ]);
    assert.strictEqual(ev.callProcedure('Test5', []), 'Moo', 'Implements works via parseAsClass');
    console.log('[PASS] Implements via parseAsClass');
}

// Test 6: Class_Initialize runs when using parseAsClass
{
    const ev = evalWithClass(`
Public Count As Integer
Private Sub Class_Initialize()
    Count = 100
End Sub
`, 'InitTest', `
Function Test6() As Integer
    Dim obj As New InitTest
    Test6 = obj.Count
End Function
`);
    assert.strictEqual(ev.callProcedure('Test6', []), 100, 'Class_Initialize runs via parseAsClass');
    console.log('[PASS] Class_Initialize via parseAsClass');
}

// Test 7: parseAsClass with Option Explicit (realistic .cls file style)
{
    const ev = evalWithClass(`
Option Explicit
Private mTotal As Double

Sub Add(amount As Double)
    mTotal = mTotal + amount
End Sub

Function Total() As Double
    Total = mTotal
End Function
`, 'Accumulator', `
Function Test7() As Double
    Dim acc As New Accumulator
    acc.Add 10.5
    acc.Add 4.5
    Test7 = acc.Total()
End Function
`);
    assert.strictEqual(ev.callProcedure('Test7', []), 15, 'parseAsClass with Option Explicit');
    console.log('[PASS] parseAsClass with Option Explicit');
}

// Test 8: Explicit Class...End Class syntax still works (no regression)
{
    const ev = evalVBASingle(`
Class Rectangle
    Public Width As Double
    Public Height As Double
    Function Area() As Double
        Area = Width * Height
    End Function
End Class

Function Test8() As Double
    Dim r As New Rectangle
    r.Width = 4
    r.Height = 3
    Test8 = r.Area()
End Function
`);
    assert.strictEqual(ev.callProcedure('Test8', []), 12, 'Explicit Class...End Class unaffected');
    console.log('[PASS] Explicit Class...End Class (no regression)');
}

// Test 9: B-1 — クラス内 Private Const がメソッドから参照できる
{
    const clsSource = `
Option Explicit
Private Const LOW As Long = 1
Private Const HIGH As Long = 100
Public Function InRange(ByVal v As Long) As Boolean
    InRange = (v >= LOW And v <= HIGH)
End Function
`;
    const modSource = `
Function Test9() As String
    Dim obj As New Validator
    If obj.InRange(50) Then
        Test9 = "ok"
    Else
        Test9 = "fail"
    End If
End Function
`;
    const ev = evalWithClass(clsSource, 'Validator', modSource);
    assert.strictEqual(ev.callProcedure('Test9', []), 'ok', 'Private Const accessible inside class method');
    console.log('[PASS] parseAsClass: B-1 — Private Const クラス内参照');
}

// Test 10: B-2 — クラス内プライベートメソッドを同クラスの公開メソッドから呼べる
{
    const clsSource = `
Option Explicit
Public Function Compute(ByVal n As Long) As Long
    Compute = Double(n) + Triple(n)
End Function
Private Function Double(ByVal n As Long) As Long
    Double = n * 2
End Function
Private Function Triple(ByVal n As Long) As Long
    Triple = n * 3
End Function
`;
    const modSource = `
Function Test10() As Long
    Dim obj As New Calculator
    Test10 = obj.Compute(4)
End Function
`;
    const ev = evalWithClass(clsSource, 'Calculator', modSource);
    assert.strictEqual(ev.callProcedure('Test10', []), 20, 'Private helpers callable from public method (4*2 + 4*3 = 20)');
    console.log('[PASS] parseAsClass: B-2 — クラス内プライベートメソッド呼び出し');
}

// Test 11: B-2 — クラス自身のスコープがグローバルより優先される
{
    const clsSource = `
Option Explicit
Public Function GetName() As String
    GetName = MyName()
End Function
Private Function MyName() As String
    MyName = "class"
End Function
`;
    const modSource = `
Function MyName() As String
    MyName = "global"
End Function
Function Test11() As String
    Dim obj As New Named
    Test11 = obj.GetName()
End Function
`;
    const ev = evalWithClass(clsSource, 'Named', modSource);
    assert.strictEqual(ev.callProcedure('Test11', []), 'class', 'Class scope takes priority over global scope');
    console.log('[PASS] parseAsClass: B-2 — クラス自身のスコープがグローバルより優先');
}

// Test 12: B-5 — Me.Property 代入がクラス内で動作する
{
    const clsSource = `
Option Explicit
Private m_val As Long

Public Sub SetValue(ByVal n As Long)
    Me.m_val = n
End Sub

Public Function GetValue() As Long
    GetValue = m_val
End Function
`;
    const modSource = `
Function Test12() As Long
    Dim obj As New Counter
    obj.SetValue 42
    Test12 = obj.GetValue()
End Function
`;
    const ev = evalWithClass(clsSource, 'Counter', modSource);
    assert.strictEqual(ev.callProcedure('Test12', []), 42, 'Me.Property assignment works inside class method');
    console.log('[PASS] parseAsClass: B-5 — Me.Property 代入がクラス内で動作する');
}

// Test 13: B-6 — Private WithEvents フィールドが Class_Initialize で初期化される
{
    const helperCls = `
Option Explicit
Public Value As Long
Public Event Changed(ByVal newVal As Long)

Public Sub SetValue(ByVal n As Long)
    Value = n
    RaiseEvent Changed(n)
End Sub
`;
    const clsSource = `
Option Explicit
Private WithEvents m_helper As Helper
Private m_log As String

Private Sub Class_Initialize()
    Set m_helper = New Helper
    m_log = ""
End Sub

Private Sub m_helper_Changed(ByVal newVal As Long)
    m_log = m_log & "changed:" & newVal & ";"
End Sub

Public Sub Run()
    m_helper.SetValue 10
    m_helper.SetValue 20
End Sub

Public Function Log() As String
    Log = m_log
End Function
`;
    const modSource = `
Function Test13() As String
    Dim obj As New Watcher
    obj.Run
    Test13 = obj.Log()
End Function
`;
    const modules = [
        { name: 'Helper', code: helperCls, parseAsClass: 'Helper' },
        { name: 'Watcher', code: clsSource, parseAsClass: 'Watcher' },
        { name: 'Module', code: modSource },
    ];
    const ev = evalVBAModules(modules);
    assert.strictEqual(ev.callProcedure('Test13', []), 'changed:10;changed:20;', 'Private WithEvents field initialized in Class_Initialize and events received');
    console.log('[PASS] parseAsClass: B-6 — Private WithEvents フィールドが Class_Initialize で初期化される');
}

console.log('\n✅ parseAsClass: 全テスト通過');
