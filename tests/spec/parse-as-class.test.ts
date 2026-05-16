import { Lexer } from '../../src/compiler/lexer';
import { Parser } from '../../src/compiler/parser';
import { Evaluator } from '../../src/compiler/evaluator';
import { assert } from '../../test-libs/test-runner';

function evalWithClass(clsSource: string, className: string, moduleSource: string = ''): Evaluator {
    const ev = new Evaluator(console.log);

    // Parse .cls content directly via parseAsClass (no string wrapping needed)
    const clsAst = new Parser(new Lexer(clsSource).tokenize(), { parseAsClass: className }).parse();
    ev.evaluate(clsAst);

    if (moduleSource) {
        const modAst = new Parser(new Lexer(moduleSource).tokenize()).parse();
        ev.evaluate(modAst);
    }
    return ev;
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
    const ev2 = new Evaluator(console.log);
    ev2.evaluate(new Parser(new Lexer(`Class MyClass\n${clsSource}\nEnd Class`).tokenize()).parse());
    ev2.evaluate(new Parser(new Lexer(modSource).tokenize()).parse());
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
    const dogSrc = `
Public Name As String
Function Speak() As String
    Speak = Name & " says Woof"
End Function
`;
    const catSrc = `
Public Name As String
Function Speak() As String
    Speak = Name & " says Meow"
End Function
`;
    const ev = new Evaluator(console.log);
    ev.evaluate(new Parser(new Lexer(dogSrc).tokenize(), { parseAsClass: 'Dog' }).parse());
    ev.evaluate(new Parser(new Lexer(catSrc).tokenize(), { parseAsClass: 'Cat' }).parse());

    const modAst = new Parser(new Lexer(`
Function Test4() As String
    Dim d As New Dog
    Dim c As New Cat
    d.Name = "Rex"
    c.Name = "Whiskers"
    Test4 = d.Speak() & "|" & c.Speak()
End Function
`).tokenize()).parse();
    ev.evaluate(modAst);

    assert.strictEqual(ev.callProcedure('Test4', []), 'Rex says Woof|Whiskers says Meow', 'Multiple classes via parseAsClass');
    console.log('[PASS] Multiple classes via parseAsClass');
}

// Test 5: Implements via parseAsClass
{
    const iface = `
Class IAnimal
Function Speak() As String
End Function
End Class
`;
    const impl = `
Implements IAnimal
Function IAnimal_Speak() As String
    IAnimal_Speak = "Moo"
End Function
`;
    const ev = new Evaluator(console.log);
    ev.evaluate(new Parser(new Lexer(iface).tokenize()).parse());
    ev.evaluate(new Parser(new Lexer(impl).tokenize(), { parseAsClass: 'Cow' }).parse());

    const mod = `
Function Test5() As String
    Dim a As IAnimal
    Set a = New Cow
    Test5 = a.Speak()
End Function
`;
    ev.evaluate(new Parser(new Lexer(mod).tokenize()).parse());
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
    const code = `
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
`;
    const ev = new Evaluator(console.log);
    ev.evaluate(new Parser(new Lexer(code).tokenize()).parse());
    assert.strictEqual(ev.callProcedure('Test8', []), 12, 'Explicit Class...End Class unaffected');
    console.log('[PASS] Explicit Class...End Class (no regression)');
}

console.log('\n✅ parseAsClass: 全テスト通過');
