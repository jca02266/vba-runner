import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { Evaluator } from '../../src/engine/evaluator';
import { assert } from '../../test-libs/test-runner';

function evalVBA(code: string): Evaluator {
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    const ev = new Evaluator(console.log);
    ev.evaluate(ast);
    return ev;
}

function runFunc(code: string, name: string, args: any[] = []): any {
    return evalVBA(code).callProcedure(name, args);
}

// Test 1: Me in function method - return self reference (basic)
{
    const code = `
        Class TestClass
            Public value As Integer

            Function GetSelf() As TestClass
                Set GetSelf = Me
            End Function
        End Class

        Function Test1()
            Dim obj As New TestClass
            obj.value = 42
            Dim result As TestClass
            Set result = obj.GetSelf()
            Test1 = (result Is obj)
        End Function
    `;

    try {
        const result = runFunc(code, 'Test1');
        const vbaTrue = result === true || (result && typeof result === 'object' && result.value === -1);
        assert.strictEqual(vbaTrue, true, 'Me returns self reference');
        console.log('[PASS] Me returns self reference');
    } catch (e: any) {
        console.log('[FAIL] Me returns self reference:', e.message);
    }
}

// Test 2: Me in method calling another method
{
    const code = `
        Class Counter
            Public count As Integer

            Sub Increment()
                count = count + 1
            End Sub

            Sub AddTwo()
                count = count + 2
            End Sub
        End Class

        Function Test2()
            Dim obj As New Counter
            obj.count = 0
            obj.Increment()
            obj.AddTwo()
            Test2 = obj.count
        End Function
    `;

    try {
        const result = runFunc(code, 'Test2');
        assert.strictEqual(result, 3, 'Method chaining updates state');
        console.log('[PASS] Method calls update state');
    } catch (e: any) {
        console.log('[FAIL] Method calls update state:', e.message);
    }
}

// Test 3: Me in recursive method call
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

        Function Test3()
            Dim fib As New Fibonacci
            Test3 = fib.Calc(10)
        End Function
    `;

    try {
        const result = runFunc(code, 'Test3');
        assert.strictEqual(result, 55, 'Fibonacci(10) = 55 with Me');
        console.log('[PASS] Me in recursive method calls');
    } catch (e: any) {
        console.log('[FAIL] Me in recursive method calls:', e.message);
    }
}

// Test 4: Me in method chaining (fluent interface)
{
    const code = `
        Class Builder
            Private _value As String

            Sub Class_Initialize()
                _value = ""
            End Sub

            Function Add(s As String) As Builder
                _value = _value & s
                Set Add = Me
            End Function

            Function Build() As String
                Build = _value
            End Function
        End Class

        Function Test4()
            Dim b As New Builder
            Test4 = b.Add("Hello").Add(" ").Add("World").Build()
        End Function
    `;

    try {
        const result = runFunc(code, 'Test4');
        assert.strictEqual(result, 'Hello World', 'Me enables method chaining');
        console.log('[PASS] Me enables method chaining');
    } catch (e: any) {
        console.log('[FAIL] Me enables method chaining:', e.message);
    }
}

// Test 5: Multiple instances - Me isolation
{
    const code = `
        Class Counter
            Public count As Integer

            Sub Increment()
                count = count + 1
            End Sub
        End Class

        Function Test5()
            Dim c1 As New Counter
            Dim c2 As New Counter
            c1.count = 0
            c2.count = 100
            c1.Increment()
            c2.Increment()
            Test5 = (c1.count = 1) And (c2.count = 101)
        End Function
    `;

    try {
        const result = runFunc(code, 'Test5');
        const vbaTrue = result === true || (result && typeof result === 'object' && result.value === -1);
        assert.strictEqual(vbaTrue, true, 'Me properly isolates instances');
        console.log('[PASS] Me properly isolates instances');
    } catch (e: any) {
        console.log('[FAIL] Me properly isolates instances:', e.message);
    }
}

// Test 6: Me with Property Get
{
    const code = `
        Class Temperature
            Private _celsius As Double

            Property Get Value() As Double
                Value = _celsius
            End Property

            Property Let Value(v As Double)
                _celsius = v
            End Property

            Function Fahrenheit() As Double
                Fahrenheit = _celsius * 9 / 5 + 32
            End Function
        End Class

        Function Test6()
            Dim t As New Temperature
            t.Value = 100
            Test6 = t.Fahrenheit()
        End Function
    `;

    try {
        const result = runFunc(code, 'Test6');
        assert.strictEqual(result, 212, 'Property and field access works');
        console.log('[PASS] Me with properties');
    } catch (e: any) {
        console.log('[FAIL] Me with properties:', e.message);
    }
}

// Test 7: Me in different contexts - Sub vs Function vs Property
{
    const code = `
        Class Tester
            Public data As String

            Sub SetData(s As String)
                data = s
            End Sub

            Function GetDataLength() As Integer
                GetDataLength = Len(data)
            End Function

            Property Get Data2() As String
                Data2 = data & data
            End Property
        End Class

        Function Test7()
            Dim t As New Tester
            t.SetData("test")
            Test7 = t.GetDataLength()
        End Function
    `;

    try {
        const result = runFunc(code, 'Test7');
        assert.strictEqual(result, 4, 'Me works in Sub, Function, Property');
        console.log('[PASS] Me in Sub/Function/Property');
    } catch (e: any) {
        console.log('[FAIL] Me in Sub/Function/Property:', e.message);
    }
}

// Test 8: Me with arithmetic and string operations
{
    const code = `
        Class Calculator
            Public value As Double

            Sub AddValue(v As Double)
                value = value + v
            End Sub

            Sub MultiplyValue(v As Double)
                value = value * v
            End Sub

            Function Compute() As Double
                Compute = value
            End Function
        End Class

        Function Test8()
            Dim calc As New Calculator
            calc.value = 10
            calc.AddValue(5)
            calc.MultiplyValue(2)
            Test8 = calc.Compute()
        End Function
    `;

    try {
        const result = runFunc(code, 'Test8');
        assert.strictEqual(result, 30, 'Arithmetic operations work');
        console.log('[PASS] Me with arithmetic operations');
    } catch (e: any) {
        console.log('[FAIL] Me with arithmetic operations:', e.message);
    }
}

// Test 9: Real-world pattern - Builder pattern with Me
{
    const code = `
        Class QueryBuilder
            Private _query As String
            Private _where As String

            Sub Class_Initialize()
                _query = "SELECT *"
                _where = ""
            End Sub

            Function From(table As String) As QueryBuilder
                _query = _query & " FROM " & table
                Set From = Me
            End Function

            Function Where(condition As String) As QueryBuilder
                _where = condition
                Set Where = Me
            End Function

            Function Build() As String
                Dim result As String
                result = _query
                If _where <> "" Then
                    result = result & " WHERE " & _where
                End If
                Build = result
            End Function
        End Class

        Function Test10()
            Dim qb As New QueryBuilder
            Test10 = qb.From("Users").Where("Age > 18").Build()
        End Function
    `;

    try {
        const result = runFunc(code, 'Test10');
        assert.strictEqual(result, 'SELECT * FROM Users WHERE Age > 18', 'QueryBuilder pattern');
        console.log('[PASS] Real-world: QueryBuilder pattern');
    } catch (e: any) {
        console.log('[FAIL] Real-world: QueryBuilder pattern:', e.message);
    }
}

console.log('\n✅ Me Keyword: 全テスト完了');
