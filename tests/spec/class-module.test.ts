import { evalVBASingle, assert } from '../../test-libs/test-runner';

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

console.log('\n✅ Class Module: 全テスト通過');
