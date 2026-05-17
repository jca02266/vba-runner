/**
 * Implements インターフェース呼び出しのテスト (§5.2.4.2)
 *
 * VBA の Implements ではインターフェース変数 (As IAnimal) 経由でメソッドを呼ぶと
 * 具象クラスの InterfaceName_MethodName へディスパッチされる。
 */
import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { Evaluator } from '../../src/engine/evaluator';
import { assert } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    const ev = new Evaluator(console.log);
    ev.evaluate(ast);
    return ev;
}

function runFunc(code: string, name: string, args: any[] = []): any {
    return evalVBA(code).callProcedure(name, args);
}

// --- 1. 基本: インターフェースメソッド呼び出しディスパッチ ---
{
    const code = `
Class IAnimal
    Public Sub Speak()
    End Sub
End Class

Class Dog
    Implements IAnimal

    Public Sub IAnimal_Speak()
        Debug.Print "Woof"
    End Sub

    Public Function IAnimal_Sound() As String
        IAnimal_Sound = "Woof"
    End Function
End Class

Function TestBasicDispatch() As String
    Dim obj As Object
    Set obj = New Dog
    obj.Speak
    TestBasicDispatch = "ok"
End Function
`;
    const result = runFunc(code, 'TestBasicDispatch');
    assert.strictEqual(result, 'ok', 'Implements: obj.Speak dispatches to IAnimal_Speak');
    console.log('[PASS] 基本インターフェースメソッドディスパッチ');
}

// --- 2. インターフェース変数経由の呼び出し (直接インスタンス) ---
{
    const code = `
Class IGreeter
    Public Function Greet(name As String) As String
    End Function
End Class

Class EnglishGreeter
    Implements IGreeter

    Public Function IGreeter_Greet(name As String) As String
        IGreeter_Greet = "Hello, " & name
    End Function
End Class

Function TestInterfaceGreet(name As String) As String
    Dim g As Object
    Set g = New EnglishGreeter
    TestInterfaceGreet = g.Greet(name)
End Function
`;
    const result = runFunc(code, 'TestInterfaceGreet', ['Alice']);
    assert.strictEqual(result, 'Hello, Alice', 'Implements: 引数ありメソッドのディスパッチ');
    console.log('[PASS] 引数ありメソッドのインターフェースディスパッチ');
}

// --- 3. インターフェース Property Get のディスパッチ ---
{
    const code = `
Class IShape
    Public Property Get Area() As Double
    End Property
End Class

Class Rectangle
    Implements IShape

    Private w As Double
    Private h As Double

    Public Sub Init(width As Double, height As Double)
        w = width
        h = height
    End Sub

    Public Property Get IShape_Area() As Double
        IShape_Area = w * h
    End Property
End Class

Function TestAreaProperty() As Double
    Dim s As Object
    Set s = New Rectangle
    s.Init 3, 4
    TestAreaProperty = s.Area
End Function
`;
    const result = runFunc(code, 'TestAreaProperty');
    assert.strictEqual(result, 12, 'Implements: Property Get のディスパッチ (3*4=12)');
    console.log('[PASS] インターフェース Property Get ディスパッチ');
}

// --- 4. 複数インターフェース実装 ---
{
    const code = `
Class IFlyable
    Public Sub Fly()
    End Sub
End Class

Class ISwimmable
    Public Sub Swim()
    End Sub
End Class

Class Duck
    Implements IFlyable
    Implements ISwimmable

    Public Function IFlyable_Fly() As String
        IFlyable_Fly = "flap flap"
    End Function

    Public Function ISwimmable_Swim() As String
        ISwimmable_Swim = "splash splash"
    End Function
End Class

Function TestMultiInterface() As String
    Dim d As Object
    Set d = New Duck
    Dim r1 As String
    Dim r2 As String
    r1 = d.Fly
    r2 = d.Swim
    TestMultiInterface = r1 & "," & r2
End Function
`;
    const result = runFunc(code, 'TestMultiInterface');
    assert.strictEqual(result, 'flap flap,splash splash', 'Implements: 複数インターフェースの両メソッドをディスパッチ');
    console.log('[PASS] 複数インターフェースのディスパッチ');
}

// --- 5. ダイレクト呼び出し (IAnimal_Speak) とインターフェース呼び出し (Speak) の共存 ---
{
    const code = `
Class ISpeaker
    Public Sub Speak()
    End Sub
End Class

Class Cat
    Implements ISpeaker

    Public Function ISpeaker_Speak() As String
        ISpeaker_Speak = "Meow"
    End Function
End Class

Function TestDirectAndInterface() As String
    Dim c As Object
    Set c = New Cat
    Dim r1 As String
    Dim r2 As String
    r1 = c.ISpeaker_Speak()    ' 直接呼び出し
    r2 = c.Speak()             ' インターフェース経由ディスパッチ
    TestDirectAndInterface = r1 & "," & r2
End Function
`;
    const result = runFunc(code, 'TestDirectAndInterface');
    assert.strictEqual(result, 'Meow,Meow', 'Implements: 直接呼び出しとインターフェース経由呼び出しが同じ結果');
    console.log('[PASS] 直接呼び出しとインターフェースディスパッチの共存');
}

// --- 6. VBA ベテランが書くパターン: コレクションにインターフェース変数を格納 ---
{
    const code = `
Class IProcessor
    Public Function Process(val As Long) As Long
    End Function
End Class

Class Doubler
    Implements IProcessor

    Public Function IProcessor_Process(val As Long) As Long
        IProcessor_Process = val * 2
    End Function
End Class

Class Tripler
    Implements IProcessor

    Public Function IProcessor_Process(val As Long) As Long
        IProcessor_Process = val * 3
    End Function
End Class

Function TestPolymorphism() As Long
    Dim processors As New Collection
    Dim d As Object
    Dim t As Object
    Set d = New Doubler
    Set t = New Tripler
    processors.Add d
    processors.Add t

    Dim total As Long
    total = 0
    Dim p As Object
    For Each p In processors
        total = total + p.Process(5)
    Next p
    TestPolymorphism = total
End Function
`;
    const result = runFunc(code, 'TestPolymorphism');
    assert.strictEqual(result, 25, 'Implements: ポリモーフィズム - Doubler(5*2=10) + Tripler(5*3=15) = 25');
    console.log('[PASS] ポリモーフィズム (コレクション + インターフェース)');
}

console.log('\n✅ Implements インターフェースディスパッチ: 全テスト通過');
