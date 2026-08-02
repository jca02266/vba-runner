import { evalVBASingle, evalVBAModules, assert } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    return evalVBASingle(code);
}

function runFunc(code: string, name: string, args: any[] = []): any {
    return evalVBA(code).callProcedure(name, args);
}

// --- 1. 基本的なクラス: フィールドとメソッド ---
{
    const code = `
    Class Counter
        Public Count As Integer

        Sub Increment()
            Count = Count + 1
        End Sub

        Function GetCount() As Integer
            GetCount = Count
        End Function
    End Class

    Function TestCounter()
        Dim c As New Counter
        c.Increment()
        c.Increment()
        c.Increment()
        TestCounter = c.GetCount()
    End Function
    `;
    assert.strictEqual(runFunc(code, 'TestCounter'), 3, 'Counter.Increment 3回');
    console.log('[PASS] 基本クラス: フィールドとメソッド');
}

// --- 2. Property Get / Let ---
{
    const code = `
    Class Temperature
        Private celsius As Double

        Property Get Value() As Double
            Value = celsius
        End Property

        Property Let Value(v As Double)
            celsius = v
        End Property

        Function Fahrenheit() As Double
            Fahrenheit = celsius * 9 / 5 + 32
        End Function
    End Class

    Function TestProperty()
        Dim t As New Temperature
        t.Value = 100
        TestProperty = t.Fahrenheit()
    End Function
    `;
    assert.strictEqual(runFunc(code, 'TestProperty'), 212, 'Temperature.Fahrenheit = 212');
    console.log('[PASS] Property Get / Let');
}

// --- 3. Class_Initialize ---
{
    const code = `
    Class Greeter
        Public Name As String

        Sub Class_Initialize()
            Name = "World"
        End Sub

        Function Hello() As String
            Hello = "Hello, " & Name & "!"
        End Function
    End Class

    Function TestInitialize()
        Dim g As New Greeter
        TestInitialize = g.Hello()
    End Function

    Function TestInitializeSet()
        Dim g As New Greeter
        g.Name = "VBA"
        TestInitializeSet = g.Hello()
    End Function
    `;
    assert.strictEqual(runFunc(code, 'TestInitialize'), 'Hello, World!', 'Class_Initialize でName="World"');
    assert.strictEqual(runFunc(code, 'TestInitializeSet'), 'Hello, VBA!', 'フィールド上書き後');
    console.log('[PASS] Class_Initialize');
}

// --- 4. 複数インスタンスは独立 ---
{
    const code = `
    Class Box
        Public Value As Integer
    End Class

    Function TestIndependent()
        Dim a As New Box
        Dim b As New Box
        a.Value = 10
        b.Value = 20
        TestIndependent = a.Value + b.Value
    End Function
    `;
    assert.strictEqual(runFunc(code, 'TestIndependent'), 30, '2インスタンスは独立');
    console.log('[PASS] 複数インスタンスは独立');
}

// --- 5. Set文によるインスタンス生成と代入 ---
{
    const code = `
    Class Point
        Public X As Integer
        Public Y As Integer

        Function Distance() As Double
            Distance = Sqr(X * X + Y * Y)
        End Function
    End Class

    Function TestSetNew()
        Dim p As Point
        Set p = New Point
        p.X = 3
        p.Y = 4
        TestSetNew = p.Distance()
    End Function
    `;
    assert.strictEqual(runFunc(code, 'TestSetNew'), 5, 'Set p = New Point; Distance=5');
    console.log('[PASS] Set文によるインスタンス生成');
}

// --- 6. Me キーワード ---
{
    const code = `
    Class Builder
        Private parts As String

        Sub Class_Initialize()
            parts = ""
        End Sub

        Function Add(s As String) As Builder
            parts = parts & s
            Set Add = Me
        End Function

        Function Build() As String
            Build = parts
        End Function
    End Class

    Function TestMe()
        Dim b As New Builder
        Dim result As String
        b.Add("Hello")
        b.Add(" ")
        b.Add("World")
        TestMe = b.Build()
    End Function
    `;
    assert.strictEqual(runFunc(code, 'TestMe'), 'Hello World', 'Me キーワード (メソッドチェーン)');
    console.log('[PASS] Me キーワード');
}

// --- 7. Private フィールドはクラス外からアクセス不可 ---
{
    const code = `
    Class Secret
        Private value As Integer

        Sub SetValue(v As Integer)
            value = v
        End Sub

        Function GetValue() As Integer
            GetValue = value
        End Function
    End Class

    Function TestPrivate()
        Dim s As New Secret
        s.SetValue(42)
        TestPrivate = s.GetValue()
    End Function
    `;
    assert.strictEqual(runFunc(code, 'TestPrivate'), 42, 'Private フィールドはメソッド経由でアクセス');
    console.log('[PASS] Private フィールドのカプセル化');
}

// --- 8. TypeName ---
{
    const code = `
    Class Foo
        Public X As Integer
    End Class

    Function TestTypeName()
        Dim f As New Foo
        TestTypeName = TypeName(f)
    End Function
    `;
    assert.strictEqual(runFunc(code, 'TestTypeName'), 'Foo', 'TypeName(instance)="Foo"');
    console.log('[PASS] TypeName');
}

// --- 9. 再帰メソッド呼び出し ---
{
    const code = `
    Class Fibonacci
        Function Calc(n As Integer) As Long
            If n <= 1 Then
                Calc = n
            Else
                Calc = Me.Calc(n - 1) + Me.Calc(n - 2)
            End If
        End Function
    End Class

    Function TestFib()
        Dim fib As New Fibonacci
        TestFib = fib.Calc(10)
    End Function
    `;
    assert.strictEqual(runFunc(code, 'TestFib'), 55, 'Fibonacci(10)=55');
    console.log('[PASS] 再帰メソッド呼び出し');
}

// --- 仕様バグ修正: Boolean/Currency/Byte/LongLong/LongPtr フィールドの既定値が
// Empty のままで、True/False 判定や算術に使うと Error 438 等になっていた ---
{
    const code = `
    Class Flags
        Public IsActive As Boolean
        Public Balance As Currency
        Public Level As Byte
    End Class

    Function TestDefaults() As String
        Dim f As New Flags
        Dim result As String
        If f.IsActive Then
            result = "true"
        Else
            result = "false"
        End If
        TestDefaults = result & "," & f.Balance & "," & f.Level
    End Function
    `;
    assert.strictEqual(runFunc(code, 'TestDefaults'), 'false,0,0', 'Boolean は False、Currency/Byte は 0 で既定初期化される');
console.log('[PASS] Boolean/Currency/Byte フィールドの既定値');
}

// A scalar As Object parameter must enforce the Object-reference contract at
// the class-method binding boundary.  Values rejected before the body runs
// must not mutate the receiver, while Nothing and a bound object remain valid.
{
    const code = String.raw`
    Class Receiver
        Public Hits As Long

        Public Sub Accept(ByVal value As Object)
            Hits = Hits + 1
        End Sub
    End Class

    Class Payload
        Public Value As Long
    End Class

    Function ProbeObjectArgument() As String
        Dim receiver As New Receiver, payload As New Payload, none As Object
        Dim errPrimitive As Long, errNull As Long, errNothing As Long, errObject As Long
        Set none = Nothing
        On Error Resume Next
        receiver.Accept 42
        errPrimitive = Err.Number
        Err.Clear
        receiver.Accept Null
        errNull = Err.Number
        Err.Clear
        receiver.Accept none
        errNothing = Err.Number
        Err.Clear
        receiver.Accept payload
        errObject = Err.Number
        ProbeObjectArgument = CStr(errPrimitive) & ":" & CStr(errNull) & ":" & _
            CStr(errNothing) & ":" & CStr(errObject) & ":" & CStr(receiver.Hits)
    End Function
    `;
    const result = runFunc(code, 'ProbeObjectArgument') as string;
    const [primitive, nullValue, nothing, boundObject, hits] = result.split(':').map(Number);
    assert.ok(primitive > 0, 'primitive As Object argument is rejected');
    assert.ok(nullValue > 0, 'Null As Object argument is rejected');
    assert.strictEqual(nothing, 0, 'Nothing As Object argument is accepted');
    assert.strictEqual(boundObject, 0, 'bound Object argument is accepted');
    assert.strictEqual(hits, 2, 'only accepted Object arguments execute the body');
    console.log('[PASS] クラスメソッドのスカラーAs Object引数境界');
}

// An array-designated Object parameter must reject a scalar container while
// preserving a valid Object() array and its Nothing elements.
{
    const code = String.raw`
    Class ArrayReceiver
        Public Seen As Long

        Public Sub Accept(ByRef values() As Object)
            If values(0) Is Nothing Then
                Seen = Seen + 10
            Else
                Seen = Seen + 100
            End If
        End Sub
    End Class

    Function ProbeObjectArrayArgument() As String
        Dim receiver As New ArrayReceiver, values(0 To 0) As Object, errScalar As Long
        On Error Resume Next
        receiver.Accept 42
        errScalar = Err.Number
        Err.Clear
        receiver.Accept values
        ProbeObjectArrayArgument = CStr(errScalar) & ":" & CStr(Err.Number) & ":" & CStr(receiver.Seen)
    End Function
    `;
    const result = runFunc(code, 'ProbeObjectArrayArgument') as string;
    const [scalarError, arrayError, seen] = result.split(':').map(Number);
    assert.ok(scalarError > 0, 'scalar value is rejected for an array parameter');
    assert.strictEqual(arrayError, 0, 'Object() array is accepted');
    assert.strictEqual(seen, 10, 'Nothing element remains valid');
    console.log('[PASS] As Object()配列引数のコンテナ境界');
}

// --- 仕様バグ修正: クラスフィールドの固定長配列が Empty のままで Subscript out of range になっていた ---
{
    const code = `
    Class FixedArrayHolder
        Public Items(4) As Long
        Public Names(2) As String

        Sub Init()
            Items(0) = 10
            Items(4) = 99
            Names(0) = "hello"
            Names(2) = "world"
        End Sub

        Function SumItems() As Long
            Dim i As Long, s As Long
            For i = 0 To 4
                s = s + Items(i)
            Next i
            SumItems = s
        End Function
    End Class

    Function TestFixedArray() As String
        Dim obj As New FixedArrayHolder
        obj.Init
        TestFixedArray = obj.Items(0) & "," & obj.Items(4) & "," & obj.Names(0) & "," & obj.Names(2) & "," & obj.SumItems()
    End Function
    `;
    assert.strictEqual(runFunc(code, 'TestFixedArray'), '10,99,hello,world,109', 'クラスフィールドの固定長配列が正しく初期化される');
    console.log('[PASS] クラスフィールドの固定長配列');
}

// 動的配列フィールド（ReDim）は以前から動作していたが、回帰確認
{
    const code = `
    Class DynArrayHolder
        Public Items() As Long

        Sub Init(ByVal n As Long)
            ReDim Items(n - 1)
        End Sub
    End Class

    Function TestDynArray() As Long
        Dim obj As New DynArrayHolder
        obj.Init 3
        obj.Items(0) = 5
        obj.Items(2) = 7
        TestDynArray = obj.Items(0) + obj.Items(2)
    End Function
    `;
    assert.strictEqual(runFunc(code, 'TestDynArray'), 12, 'クラスフィールドの動的配列（ReDim）が正しく動作する');
    console.log('[PASS] クラスフィールドの動的配列（ReDim）');
}

// [回帰] Bug W1: With ブロック内でユーザー定義クラスの Property Get にアクセスできない
// `evaluateImplicitWithObjectExpression` に `__vbaClass__` ブランチが欠けていた
{
    const code = `
    Class Scorer
        Private mScore As Double

        Public Property Get Score() As Double
            Score = mScore
        End Property

        Public Property Let Score(ByVal v As Double)
            mScore = v
        End Property

        Public Function Summary() As String
            Summary = "Score:" & mScore
        End Function
    End Class

    Function TestWithPropertyGet() As Double
        Dim s As Scorer
        Set s = New Scorer
        s.Score = 42
        Dim result As Double
        With s
            result = .Score
        End With
        TestWithPropertyGet = result
    End Function

    Function TestWithFunction() As String
        Dim s As Scorer
        Set s = New Scorer
        s.Score = 42
        Dim txt As String
        With s
            txt = .Summary()
        End With
        TestWithFunction = txt
    End Function
    `;
    assert.strictEqual(runFunc(code, 'TestWithPropertyGet'), 42, 'With ブロック内で Property Get が動作する');
    assert.strictEqual(runFunc(code, 'TestWithFunction'), 'Score:42', 'With ブロック内で Function 呼び出しが動作する');
    console.log('[PASS] With ブロック内 Property Get + Function（Bug W1）');
}

// Bug CA: `.Data(0,0) = val` inside With block — CallExpression with ImplicitWithObjectExpression callee was unhandled
{
    const code = `
Class Matrix
    Public Data(2, 2) As Long
    Public Tags(3) As String
End Class
Function TestCA_2D() As Long
    Dim m As New Matrix
    With m
        .Data(0, 0) = 1
        .Data(1, 1) = 5
        .Data(2, 2) = 9
    End With
    TestCA_2D = m.Data(0, 0) + m.Data(1, 1) + m.Data(2, 2)
End Function
Function TestCA_1D() As String
    Dim m As New Matrix
    With m
        .Tags(0) = "a"
        .Tags(1) = "b"
        .Tags(2) = "c"
    End With
    TestCA_1D = m.Tags(0) & m.Tags(1) & m.Tags(2)
End Function
`;
    assert.strictEqual(runFunc(code, 'TestCA_2D'), 15, 'Bug CA: With ブロック内 .Data(i,j) = val (2D 配列フィールド)');
    assert.strictEqual(runFunc(code, 'TestCA_1D'), 'abc', 'Bug CA: With ブロック内 .Tags(i) = val (1D 配列フィールド)');
    console.log('[PASS] Bug CA: With ブロック内 .ArrayField(i[,j]) = val');
}

// Bug BY: `Public P As New ClassName` class field — not initialized (returned Nothing instead of instance)
{
    const code = `
Class Point
    Public X As Long
End Class
Class Container
    Public P As New Point
End Class
Function TestBY_Direct() As String
    Dim c As New Container
    c.P.X = 42
    TestBY_Direct = CStr(c.P.X)
End Function
Function TestBY_With() As String
    Dim c As New Container
    With c
        With .P
            .X = 99
        End With
    End With
    TestBY_With = CStr(c.P.X)
End Function
`;
    assert.strictEqual(runFunc(code, 'TestBY_Direct'), '42', 'Bug BY: Public P As New Point — direct chain access c.P.X');
    assert.strictEqual(runFunc(code, 'TestBY_With'), '99', 'Bug BY: Public P As New Point — nested With .P block');
    console.log('[PASS] Bug BY: Public P As New ClassName — クラスフィールドの As New 初期化');
}

// Bug BZ: c.PropGet("key") — Property Get (no params) returns Object; args should subscript the returned value
{
    const code = `
Class Bag
    Private m_dict As Object
    Property Get Dict() As Object
        Set Dict = m_dict
    End Property
    Property Set Dict(d As Object)
        Set m_dict = d
    End Property
    Property Get Items() As Collection
        Dim col As New Collection
        col.Add "a"
        col.Add "b"
        Set Items = col
    End Property
End Class
Function TestBZ_Dict() As String
    Dim b As New Bag
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "x", 99
    Set b.Dict = d
    TestBZ_Dict = CStr(b.Dict("x"))
End Function
Function TestBZ_Col() As String
    Dim b As New Bag
    TestBZ_Col = b.Items(1) & "," & b.Items(2)
End Function
`;
    assert.strictEqual(runFunc(code, 'TestBZ_Dict'), '99', 'Bug BZ: c.Dict("x") — Property Get returns Dictionary; subscript with key');
    assert.strictEqual(runFunc(code, 'TestBZ_Col'), 'a,b', 'Bug BZ: c.Items(n) — Property Get returns Collection; subscript with index');
    console.log('[PASS] Bug BZ: Property Get 返却オブジェクトへの添字アクセス (c.Prop("key") / c.Prop(n))');
}

// Bug CC: `Public Code As String * 5` class field — fixed-length not enforced on assignment
{
    const code = `
Class Record
    Public Code As String * 5
    Public Name As String * 10
End Class
Function TestCC_Assign() As String
    Dim r As New Record
    r.Code = "AB"
    r.Name = "Hello World Extra"
    TestCC_Assign = "|" & r.Code & "|" & Len(r.Code) & "|" & r.Name & "|" & Len(r.Name)
End Function
Function TestCC_Default() As String
    Dim r As New Record
    TestCC_Default = "|" & Len(r.Code) & "|" & Len(r.Name)
End Function
Function TestCC_With() As String
    Dim r As New Record
    With r
        .Code = "XY"
        .Name = "Short"
    End With
    TestCC_With = "|" & r.Code & "|" & Len(r.Code) & "|" & r.Name & "|" & Len(r.Name)
End Function
`;
    assert.strictEqual(runFunc(code, 'TestCC_Assign'), '|AB   |5|Hello Worl|10', 'Bug CC: fixed-length string truncated/padded on assign');
    assert.strictEqual(runFunc(code, 'TestCC_Default'), '|5|10', 'Bug CC: fixed-length string field defaults to spaces');
    assert.strictEqual(runFunc(code, 'TestCC_With'), '|XY   |5|Short     |10', 'Bug CC: fixed-length string enforced in With block');
    console.log('[PASS] Bug CC: クラスフィールド固定長文字列の切り詰め・空白パディング');
}

// Bug CD: `Public Next As Object` — VBA keyword "Next" not accepted as class field name
{
    const code = `
Class Node
    Public Value As Long
    Public Next As Object
End Class
Function TestCD_List() As String
    Dim n1 As New Node, n2 As New Node, n3 As New Node
    n1.Value = 1 : n2.Value = 2 : n3.Value = 3
    Set n1.Next = n2
    Set n2.Next = n3
    Dim cur As Object
    Set cur = n1
    Dim s As String
    Do While Not cur Is Nothing
        s = s & cur.Value & ","
        Set cur = cur.Next
    Loop
    TestCD_List = s
End Function
Function TestCD_IsNothing() As String
    Dim n As New Node
    If n.Next Is Nothing Then TestCD_IsNothing = "nothing" Else TestCD_IsNothing = "not-nothing"
End Function
`;
    assert.strictEqual(runFunc(code, 'TestCD_List'), '1,2,3,', 'Bug CD: linked list traversal with Public Next As Object');
    assert.strictEqual(runFunc(code, 'TestCD_IsNothing'), 'nothing', 'Bug CD: unset Public Next As Object Is Nothing');
    console.log('[PASS] Bug CD: VBA キーワード "Next" をクラスフィールド名として使用可能');
}

// Bug CE: When Property Set is declared before Property Get with the same name,
// a call/read expression like `w.Inner("key")` incorrectly invoked the setter
{
    const code = `
Class Wrapper
    Private m_obj As Object
    Property Set Inner(obj As Object)
        Set m_obj = obj
    End Property
    Property Get Inner() As Object
        Set Inner = m_obj
    End Property
    Property Get HasInner() As Boolean
        HasInner = Not (m_obj Is Nothing)
    End Property
End Class
Function TestCE_GetOverSet() As String
    Dim w As New Wrapper
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "key", 42
    Set w.Inner = d
    Dim result As String
    If w.HasInner Then result = "has:" & w.Inner("key") Else result = "empty"
    Set w.Inner = Nothing
    If w.HasInner Then result = result & ",still-has" Else result = result & ",cleared"
    TestCE_GetOverSet = result
End Function
`;
    assert.strictEqual(runFunc(code, 'TestCE_GetOverSet'), 'has:42,cleared', 'Bug CE: Property Get takes priority over Property Set in read context');
    console.log('[PASS] Bug CE: 読み取りコンテキストでは Property Get が Property Set より優先される');
}

// --- Bug CI: 複数モジュールで Dim obj As New ClassName が Long 0 に化けていた ---
// evalVBAModules でクラス名と同名のモジュールキーが VbaNamespaceRef として env に入ると、
// evaluateVariableDeclaration の Enum チェックがそれを enum と誤判定して
// setVariableType('Long') を呼んでしまい、obj が 0 になっていた。
{
    const ev = evalVBAModules([
        { name: 'MyClass', code: 'Class MyClass\nPublic Value As Integer\nEnd Class' },
        { name: 'Module', code: `
Function TestCI_TypeName() As String
    Dim obj As New MyClass
    TestCI_TypeName = TypeName(obj)
End Function
Function TestCI_Member() As Integer
    Dim obj As New MyClass
    obj.Value = 21
    TestCI_Member = obj.Value
End Function
` },
    ]);
    assert.strictEqual(ev.callProcedure('TestCI_TypeName', []), 'MyClass',
        'Bug CI: Dim obj As New MyClass in multi-module — TypeName is MyClass, not Long');
    assert.strictEqual(ev.callProcedure('TestCI_Member', []), 21,
        'Bug CI: Dim obj As New MyClass in multi-module — member assignment works');
    console.log('[PASS] Bug CI: 複数モジュールで Dim obj As New ClassName が正しくインスタンス化される');
}

// --- Bug 29-F/G: Sub メソッドを eval() 経由で呼べる ---
{
    const counterCode = 'Class Counter\nPrivate mVal As Long\nPublic Sub Increment()\n    mVal = mVal + 1\nEnd Sub\nPublic Property Get Value() As Long\n    Value = mVal\nEnd Property\nEnd Class';

    // 29-F: eval("a.Increment") → fast path (expression) → calls Sub
    {
        const ev = evalVBAModules([
            { name: 'Counter', code: counterCode },
            { name: 'Module1', code: 'Sub Setup()\nDim a As New Counter\nSet gA = a\nEnd Sub\nDim gA As Counter' },
        ]);
        // Use Dim/Set to create instance without touching private API
        ev.evalExpression('Dim a As New Counter');
        ev.evalExpression('a.Increment');
        ev.evalExpression('a.Increment');
        assert.strictEqual(ev.evalExpression('a.Value'), 2, 'Bug 29-F: eval("a.Increment") は Sub を呼び出す（戻り値0で黙殺しない）');
        console.log('[PASS] Bug 29-F: eval("a.Increment") で Sub が実行される');
    }

    // 29-G: multi-statement eval with Dim → slow path → no Error 450
    {
        const ev = evalVBAModules([
            { name: 'Counter', code: counterCode },
            { name: 'Module1', code: '' },
        ]);
        let out = '';
        (ev as any).onPrint = (s: string) => { out += s; };
        ev.evalExpression('Dim g As New Counter : g.Increment : Debug.Print g.Value');
        assert.strictEqual(out.trim(), '1', 'Bug 29-G: Dim g As New Counter : g.Increment が Error 450 を出さず Sub を呼ぶ');
        console.log('[PASS] Bug 29-G: multi-statement eval で Sub コールが正常動作する');
    }
}

// Bug 33-A（評価 #33）: クラス内の非修飾・括弧なし自メンバー参照（Property Get / Function）
// が未定義変数扱いで Empty になっていた（Option Explicit でも検出されない silent bug）
{
    const cls = `
Option Explicit
Public Property Get PropX() As String
    PropX = "P"
End Property
Public Function FnY() As String
    FnY = "F"
End Function
Public Function CallsBoth() As String
    CallsBoth = PropX & FnY & "!"
End Function
`;
    const main = `
Function T33A() As String
    Dim w As New SelfRef33A
    T33A = w.CallsBoth
End Function
`;
    const ev = evalVBAModules([
        { name: 'SelfRef33A', code: cls, parseAsClass: 'SelfRef33A' },
        { name: 'Main', code: main },
    ]);
    assert.strictEqual(ev.callProcedure('T33A', []), 'PF!',
        'Bug 33-A: クラス内の括弧なし自 Property Get / Function 参照が暗黙の Me.<name> に解決される');
    console.log('[PASS] Bug 33-A: クラス内の括弧なし自メンバー参照');
}

// Bug 73-A: クラスフィールドの動的型付き配列がReDim後に要素型を失う
{
    const code = `
Class TypedBucket
    Private values() As Integer
    Public Sub SetFirst(ByVal value As Double)
        ReDim values(0 To 0)
        values(0) = value
    End Sub
    Public Function FirstValue() As Double
        FirstValue = values(0)
    End Function
End Class

Function TestTypedClassArray() As Double
    Dim bucket As New TypedBucket
    bucket.SetFirst 1.5
    TestTypedClassArray = bucket.FirstValue
End Function
`;
    assert.strictEqual(runFunc(code, 'TestTypedClassArray'), 2,
        'クラスフィールドのInteger配列もReDim後に要素代入を強制変換する');
    console.log('[PASS] Bug 73-A: クラスフィールドのReDim後も配列要素型を保持する');
}

// Bug 130-A: クラスフィールド配列要素をByRef引数へ渡した書き戻しが失われる
{
    const code = `
Class ByRefArrayStore
    Private values() As Long
    Public Function Repro() As Long
        Dim amount As Long
        ReDim values(0 To 0)
        values(0) = 10
        amount = 5
        AddTo values(0), amount
        Repro = values(0)
    End Function
    Private Sub AddTo(ByRef target As Long, ByRef amount As Long)
        target = target + amount
    End Sub
End Class

Function TestClassArrayByRef() As Long
    Dim store As New ByRefArrayStore
    TestClassArrayByRef = store.Repro
End Function
`;
    assert.strictEqual(runFunc(code, 'TestClassArrayByRef'), 15,
        'クラスフィールド配列要素へのByRef書き戻しを保持する');
    console.log('[PASS] Bug 130-A: クラスフィールド配列要素のByRef書き戻し');
}

// Bug 132-A: Property Set の ByRef オブジェクト引数の書き戻しが失われる
{
    const ev = evalVBAModules([
        { name: 'SetChild', code: `
Class SetChild
    Public Value As Long
End Class
` },
        { name: 'SetHolder', code: `
Class SetHolder
    Private mChild As SetChild
    Public Property Set Ref(ByRef value As SetChild)
        Set mChild = value
        Set value = Nothing
    End Property
    Public Property Get RefValue() As Long
        RefValue = mChild.Value
    End Property
End Class
` },
        { name: 'SetDriver', code: `
Function TestPropertySetByRef() As String
    Dim holder As New SetHolder
    Dim child As New SetChild
    child.Value = 77
    Set holder.Ref = child
    If child Is Nothing Then
        TestPropertySetByRef = CStr(holder.RefValue) & ":nothing"
    Else
        TestPropertySetByRef = CStr(holder.RefValue) & ":alive"
    End If
End Function
` },
    ]);
    assert.strictEqual(ev.callProcedure('TestPropertySetByRef', []), '77:nothing',
        'Property SetのByRefオブジェクト引数を書き戻す');
    console.log('[PASS] Bug 132-A: Property SetのByRef書き戻し');
}

// Bug 139-A: Property Let の ByRef value 引数を呼出元へ書き戻す
{
    const ev = evalVBAModules([
        { name: 'LetLedger', code: `
Class LetLedger
    Private mEntries As Collection
    Private Sub Class_Initialize()
        Set mEntries = New Collection
    End Sub
    Public Property Let Label(ByRef value As String)
        value = UCase(Trim(value))
        If Len(value) = 0 Then value = "INVALID"
        mEntries.Add value
    End Property
    Public Property Get Summary() As String
        Dim item As Variant
        For Each item In mEntries
            If Summary <> "" Then Summary = Summary & ","
            Summary = Summary & item
        Next item
    End Property
End Class
` },
        { name: 'LetDriver', code: `
Function TestPropertyLetByRef() As String
    Dim ledger As New LetLedger
    Dim label As String
    Dim labels(0 To 1) As String
    label = " green "
    labels(0) = " red "
    labels(1) = "  "
    ledger.Label = label
    ledger.Label = labels(0)
    ledger.Label = labels(1)
    TestPropertyLetByRef = label & "|" & labels(0) & "|" & labels(1) & "|" & ledger.Summary
End Function
` },
    ]);
    assert.strictEqual(ev.callProcedure('TestPropertyLetByRef', []), 'GREEN|RED|INVALID|GREEN,RED,INVALID',
        'Property LetのByRef value引数を書き戻す');
    console.log('[PASS] Bug 139-A: Property LetのByRef書き戻し');
}

// Bug 141-A: With内の添字付きProperty Letも右辺へ書き戻す
{
    const ev = evalVBASingle(`
        Class IndexedLedger
            Public Property Let Label(ByVal index As Long, ByRef value As String)
                value = UCase(Trim(value))
            End Property
        End Class

        Function TestIndexedPropertyLetByRef() As String
            Dim ledger As New IndexedLedger
            Dim label As String
            label = " green "
            With ledger
                .Label(0) = label
            End With
            TestIndexedPropertyLetByRef = label
        End Function
    `);
    assert.strictEqual(ev.callProcedure('TestIndexedPropertyLetByRef', []), 'GREEN',
        'With内の添字付きProperty Letを書き戻す');
    console.log('[PASS] Bug 141-A: With内の添字付きProperty Let書き戻し');
}

// Bug 141-B: 添字付きProperty Setでも右辺オブジェクトを書き戻す
{
    const ev = evalVBAModules([
        { name: 'IndexedChild', code: `
Class IndexedChild
    Public Value As Long
End Class
` },
        { name: 'IndexedHolder', code: `
Class IndexedHolder
    Private mChild As IndexedChild
    Public Property Set Ref(ByVal index As Long, ByRef value As IndexedChild)
        Set mChild = value
        Set value = Nothing
    End Property
    Public Property Get RefValue() As Long
        RefValue = mChild.Value
    End Property
End Class
` },
        { name: 'IndexedDriver', code: `
Function TestIndexedPropertySetByRef() As String
    Dim holder As New IndexedHolder
    Dim child As New IndexedChild
    child.Value = 9
    Set holder.Ref(1) = child
    If child Is Nothing Then
        TestIndexedPropertySetByRef = CStr(holder.RefValue) & ":nothing"
    Else
        TestIndexedPropertySetByRef = CStr(holder.RefValue) & ":alive"
    End If
End Function
` },
    ]);
    assert.strictEqual(ev.callProcedure('TestIndexedPropertySetByRef', []), '9:nothing',
        '添字付きProperty SetのByRef valueを書き戻す');
    console.log('[PASS] Bug 141-B: 添字付きProperty Set書き戻し');
}

// Bug 151-A: 名前付きProperty Let引数でも暗黙ByRefを書き戻す
{
    const ev = evalVBASingle(`
        Class NamedLet
            Public Property Let Value(ByVal index As Long, ByRef text As String)
                text = UCase(text)
            End Property
        End Class

        Function TestNamedPropertyLetByRef() As String
            Dim obj As New NamedLet
            Dim text As String
            text = "named"
            Call obj.Value(index:=1, text:=text)
            TestNamedPropertyLetByRef = text
        End Function
    `);
    assert.strictEqual(ev.callProcedure('TestNamedPropertyLetByRef', []), 'NAMED',
        '名前付きProperty Let引数のByRef書き戻しを保持する');
    console.log('[PASS] Bug 151-A: 名前付きProperty LetのByRef書き戻し');
}

// Call-style Property Let with a named value must select the setter even when
// a same-named Property Get is also present.
{
    const ev = evalVBASingle(`
        Class NamedValue
            Private mValue As Variant
            Public Property Get Value() As Variant
                Value = mValue
            End Property
            Public Property Let Value(ByRef value As Variant)
                mValue = value
            End Property
        End Class

        Function TestNamedPropertyLetValue() As String
            Dim obj As New NamedValue
            Dim value As Variant
            value = 42
            Call obj.Value(Value:=value)
            TestNamedPropertyLetValue = CStr(obj.Value)
        End Function
    `);
    assert.strictEqual(ev.callProcedure('TestNamedPropertyLetValue', []), '42',
        'Call-style named Property Let selects the setter');
    console.log('[PASS] Call-style named Property Let setter selection');
}

// Bug 153-A: ローカル2次元クラス配列の要素代入後もReDim Preserveできる
{
    const ev = evalVBASingle(`
        Class PreserveBox
            Public Value As Long
        End Class

        Function TestLocalObjectArrayPreserve() As String
            Dim boxes() As PreserveBox
            ReDim boxes(1 To 1, 2 To 2)
            Set boxes(1, 2) = New PreserveBox
            boxes(1, 2).Value = 12
            ReDim Preserve boxes(1 To 1, 2 To 3)
            Set boxes(1, 3) = New PreserveBox
            boxes(1, 3).Value = 13
            TestLocalObjectArrayPreserve = CStr(boxes(1, 2).Value) & ":" & CStr(boxes(1, 3).Value)
        End Function
    `);
    assert.strictEqual(ev.callProcedure('TestLocalObjectArrayPreserve', []), '12:13',
        'ローカル2次元クラス配列の要素を保持してPreserve拡張する');
    console.log('[PASS] Bug 153-A: ローカル2次元クラス配列のPreserve');
}

// Bug 182-A: a class ByVal argument must not evaluate a member call twice.
{
    const ev = evalVBASingle(`
        Class LineCollector
            Public lines As Collection
            Private Sub Class_Initialize()
                Set lines = New Collection
            End Sub
            Public Sub AddLine(ByVal value As String)
                lines.Add value
            End Sub
            Public Function Joined() As String
                Dim item As Variant
                For Each item In lines
                    Joined = Joined & item & "|"
                Next item
            End Function
        End Class

        Function TestByValMemberCall() As String
            Dim fso As New FileSystemObject
            Dim ts As Object
            Dim c As New LineCollector
            Set ts = fso.CreateTextFile("C:\\byval-readline.txt", True, False)
            ts.WriteLine "one"
            ts.WriteLine "two"
            ts.Close
            Set ts = fso.OpenTextFile("C:\\byval-readline.txt", 1, False, False)
            c.AddLine ts.ReadLine
            c.AddLine ts.ReadLine
            ts.Close
            TestByValMemberCall = c.Joined
        End Function
    `);
    assert.strictEqual(ev.callProcedure('TestByValMemberCall', []), 'one|two|',
        'class ByVal member-call arguments are evaluated once');
    console.log('✅ Bug 182-A: class ByVal member-call evaluation');
}

console.log('\n✅ Class Module: 全テスト通過');
