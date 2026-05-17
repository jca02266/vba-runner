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

// Test 1: ParamArray with zero arguments (empty array, LBound=0, UBound=-1)
{
    const code = `
        Function Test1() As String
            Dim result As String
            result = TestParamArray()
            Test1 = result
        End Function

        Function TestParamArray(ParamArray args()) As String
            Dim result As String
            result = "LBound=" & LBound(args) & ",UBound=" & UBound(args) & ",Len=" & (UBound(args) - LBound(args) + 1)
            TestParamArray = result
        End Function
    `;

    try {
        const result = runFunc(code, 'Test1');
        assert.strictEqual(result, 'LBound=0,UBound=-1,Len=0', 'ParamArray with no args is empty array (UBound=-1)');
        console.log('[PASS] ParamArray zero arguments');
    } catch (e: any) {
        console.log('[FAIL] ParamArray zero arguments:', e.message);
    }
}

// Test 2: ParamArray with one extra argument
{
    const code = `
        Function Test2() As String
            Dim result As String
            result = TestParamArray(10)
            Test2 = result
        End Function

        Function TestParamArray(ParamArray args()) As String
            TestParamArray = "Count=" & (UBound(args) - LBound(args) + 1) & ",First=" & args(0)
        End Function
    `;

    try {
        const result = runFunc(code, 'Test2');
        assert.strictEqual(result, 'Count=1,First=10', 'ParamArray with one argument');
        console.log('[PASS] ParamArray with one argument');
    } catch (e: any) {
        console.log('[FAIL] ParamArray with one argument:', e.message);
    }
}

// Test 3: ParamArray with multiple arguments
{
    const code = `
        Function Test3() As String
            Dim result As String
            result = TestParamArray(1, 2, 3, 4, 5)
            Test3 = result
        End Function

        Function TestParamArray(ParamArray args()) As String
            Dim sum As Integer, i As Integer
            sum = 0
            For i = LBound(args) To UBound(args)
                sum = sum + args(i)
            Next
            TestParamArray = "Sum=" & sum
        End Function
    `;

    try {
        const result = runFunc(code, 'Test3');
        assert.strictEqual(result, 'Sum=15', 'ParamArray aggregates multiple arguments');
        console.log('[PASS] ParamArray with multiple arguments');
    } catch (e: any) {
        console.log('[FAIL] ParamArray with multiple arguments:', e.message);
    }
}

// Test 4: ParamArray with required positional parameter
{
    const code = `
        Function Test4() As String
            Test4 = TestFunc(100)
        End Function

        Function TestFunc(required As Integer, ParamArray args()) As String
            Dim count As Integer
            count = UBound(args) - LBound(args) + 1
            TestFunc = "Required=" & required & ",ParamCount=" & count
        End Function
    `;

    try {
        const result = runFunc(code, 'Test4');
        assert.strictEqual(result, 'Required=100,ParamCount=0', 'ParamArray after required param with no extra args');
        console.log('[PASS] ParamArray after required parameter');
    } catch (e: any) {
        console.log('[FAIL] ParamArray after required parameter:', e.message);
    }
}

// Test 5: ParamArray with required param + extra arguments
{
    const code = `
        Function Test5() As String
            Test5 = TestFunc(100, 1, 2, 3)
        End Function

        Function TestFunc(required As Integer, ParamArray args()) As String
            Dim count As Integer
            count = UBound(args) - LBound(args) + 1
            TestFunc = "Required=" & required & ",ParamCount=" & count & ",Sum=" & (args(0) + args(1) + args(2))
        End Function
    `;

    try {
        const result = runFunc(code, 'Test5');
        assert.strictEqual(result, 'Required=100,ParamCount=3,Sum=6', 'ParamArray receives extra arguments correctly');
        console.log('[PASS] ParamArray with required param and extra args');
    } catch (e: any) {
        console.log('[FAIL] ParamArray with required param and extra args:', e.message);
    }
}

// Test 6: ParamArray with For Each iteration
{
    const code = `
        Function Test6() As String
            Test6 = TestFunc(10, 20, 30)
        End Function

        Function TestFunc(ParamArray args()) As String
            Dim result As String
            result = ""
            Dim elem As Variant
            For Each elem In args
                result = result & elem & ","
            Next
            TestFunc = result
        End Function
    `;

    try {
        const result = runFunc(code, 'Test6');
        assert.strictEqual(result, '10,20,30,', 'ParamArray works with For Each');
        console.log('[PASS] ParamArray with For Each iteration');
    } catch (e: any) {
        console.log('[FAIL] ParamArray with For Each iteration:', e.message);
    }
}

// Test 7: ParamArray ByRef semantics (spec §5.3.1.5: elements behave as ByRef positional params)
{
    const code = `
        Function Test7() As String
            Dim a As Integer, b As Integer, c As Integer
            a = 1: b = 2: c = 3
            Call ModifyParams(a, b, c)
            Test7 = a & "," & b & "," & c
        End Function

        Sub ModifyParams(ParamArray params())
            Dim i As Integer
            For i = 0 To UBound(params)
                params(i) = params(i) * 10
            Next
        End Sub
    `;

    try {
        const result = runFunc(code, 'Test7');
        assert.strictEqual(result, '10,20,30', 'ParamArray ByRef: modifications propagate back to caller variables');
        console.log('[PASS] ParamArray ByRef semantics');
    } catch (e: any) {
        console.log('[FAIL] ParamArray ByRef semantics:', e.message);
    }
}

// Test 8: ParamArray with string values
{
    const code = `
        Function Test8() As String
            Test8 = Concatenate("Hello", " ", "World", "!")
        End Function

        Function Concatenate(ParamArray words()) As String
            Dim result As String
            result = ""
            Dim i As Integer
            For i = 0 To UBound(words)
                result = result & words(i)
            Next
            Concatenate = result
        End Function
    `;

    try {
        const result = runFunc(code, 'Test8');
        assert.strictEqual(result, 'Hello World!', 'ParamArray handles string arguments');
        console.log('[PASS] ParamArray with string values');
    } catch (e: any) {
        console.log('[FAIL] ParamArray with string values:', e.message);
    }
}

// Test 9: Real-world pattern - variadic logging function
{
    const code = `
        Function Test9() As String
            Dim log As String
            log = Log("INFO", "User", "logged", "in")
            Test9 = log
        End Function

        Function Log(level As String, ParamArray message()) As String
            Dim result As String
            result = level & ": "
            Dim i As Integer
            For i = 0 To UBound(message)
                result = result & message(i)
                If i < UBound(message) Then
                    result = result & " "
                End If
            Next
            Log = result
        End Function
    `;

    try {
        const result = runFunc(code, 'Test9');
        assert.strictEqual(result, 'INFO: User logged in', 'Real-world: variadic logging function');
        console.log('[PASS] Real-world: variadic logging function');
    } catch (e: any) {
        console.log('[FAIL] Real-world: variadic logging function:', e.message);
    }
}

// Test 10: ParamArray with mixed types (Variant array)
{
    const code = `
        Function Test10() As String
            Test10 = MixedTypes(10, "text", 3.14, True)
        End Function

        Function MixedTypes(ParamArray args()) As String
            Dim result As String
            result = "Count=" & (UBound(args) + 1)
            result = result & ",T1=" & TypeName(args(0))
            result = result & ",T2=" & TypeName(args(1))
            result = result & ",T3=" & TypeName(args(2))
            result = result & ",T4=" & TypeName(args(3))
            MixedTypes = result
        End Function
    `;

    try {
        const result = runFunc(code, 'Test10');
        const expected = 'Count=4,T1=Double,T2=String,T3=Double,T4=Boolean';
        assert.strictEqual(result, expected, 'ParamArray preserves element types as Variant');
        console.log('[PASS] ParamArray with mixed types');
    } catch (e: any) {
        console.log('[FAIL] ParamArray with mixed types:', e.message);
    }
}

// Test 11: ParamArray ByRef with literals - literals are r-values, silently not written back
{
    const code = `
        Function Test11() As String
            Call ModifyParams(10, 20)
            Test11 = "ok"
        End Function

        Sub ModifyParams(ParamArray params())
            params(0) = params(0) * 99
        End Sub
    `;

    try {
        const result = runFunc(code, 'Test11');
        assert.strictEqual(result, 'ok', 'ParamArray ByRef with literals: does not crash');
        console.log('[PASS] ParamArray ByRef with literals (r-values silently ignored)');
    } catch (e: any) {
        console.log('[FAIL] ParamArray ByRef with literals:', e.message);
    }
}

// Test 12: Real-world pattern - normalize array values via ParamArray ByRef
{
    const code = `
        Function Test12() As String
            Dim x As Double, y As Double, z As Double
            x = 3: y = 4: z = 5
            NormalizeInPlace x, y, z
            Test12 = x & "," & y & "," & z
        End Function

        Sub NormalizeInPlace(ParamArray vals())
            Dim total As Double, i As Integer
            total = 0
            For i = 0 To UBound(vals)
                total = total + vals(i)
            Next
            If total > 0 Then
                For i = 0 To UBound(vals)
                    vals(i) = vals(i) / total
                Next
            End If
        End Sub
    `;

    try {
        const result = runFunc(code, 'Test12');
        const parts = String(result).split(',').map(Number);
        const expected = [3/12, 4/12, 5/12];
        const ok = parts.every((v, i) => Math.abs(v - expected[i]) < 1e-10);
        assert.strictEqual(ok, true, 'ParamArray ByRef normalization pattern');
        console.log('[PASS] Real-world: normalize via ParamArray ByRef');
    } catch (e: any) {
        console.log('[FAIL] Real-world: normalize via ParamArray ByRef:', e.message);
    }
}

console.log('\n✅ ParamArray Edge Cases: 全テスト完了');
