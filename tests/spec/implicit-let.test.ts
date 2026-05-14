import { Lexer } from '../../src/compiler/lexer';
import { Parser } from '../../src/compiler/parser';
import { Evaluator } from '../../src/compiler/evaluator';
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

// Test 1: Implicit Let - function call without Call keyword
{
    const code = `
    Function GetValue()
        GetValue = 42
    End Function

    Function Test1()
        Dim x As Integer
        x = GetValue()
        Test1 = x
    End Function
    `;
    const result = runFunc(code, 'Test1');
    assert.strictEqual(result, 42, 'Implicit assignment of function return value');
    console.log('[PASS] Test 1: Implicit Let - function assignment');
}

// Test 2: Function call with implicit return value discard
{
    const code = `
    Function GetValue()
        GetValue = 42
    End Function

    Function Test2()
        GetValue()
        Test2 = 1
    End Function
    `;
    const result = runFunc(code, 'Test2');
    assert.strictEqual(result, 1, 'Return value discarded when not assigned');
    console.log('[PASS] Test 2: Return value discard');
}

// Test 3: Procedure call with Call keyword
{
    const code = `
    Sub DoSomething()
        ' Just a procedure
    End Sub

    Function Test3()
        Call DoSomething()
        Test3 = 1
    End Function
    `;
    const result = runFunc(code, 'Test3');
    assert.strictEqual(result, 1, 'Call keyword works for procedures');
    console.log('[PASS] Test 3: Call keyword');
}

// Test 4: Procedure call without Call keyword
{
    const code = `
    Sub DoSomething()
        ' Just a procedure
    End Sub

    Function Test4()
        DoSomething
        Test4 = 1
    End Function
    `;
    const result = runFunc(code, 'Test4');
    assert.strictEqual(result, 1, 'Procedure call without Call keyword');
    console.log('[PASS] Test 4: Procedure without Call');
}

// Test 5: Implicit Let with object property
{
    const code = `
    Function Test5()
        Dim result As Integer
        result = 100
        Test5 = result
    End Function
    `;
    const result = runFunc(code, 'Test5');
    assert.strictEqual(result, 100, 'Implicit Let with variable assignment');
    console.log('[PASS] Test 5: Variable assignment');
}

// Test 6: Multiple implicit assignments
{
    const code = `
    Function GetA()
        GetA = 10
    End Function

    Function GetB()
        GetB = 20
    End Function

    Function Test6()
        Dim a As Integer, b As Integer
        a = GetA()
        b = GetB()
        Test6 = a + b
    End Function
    `;
    const result = runFunc(code, 'Test6');
    assert.strictEqual(result, 30, 'Multiple implicit assignments');
    console.log('[PASS] Test 6: Multiple assignments');
}

// Test 7: Nested function calls
{
    const code = `
    Function Inner()
        Inner = 5
    End Function

    Function Outer()
        Outer = Inner() * 2
    End Function

    Function Test7()
        Test7 = Outer()
    End Function
    `;
    const result = runFunc(code, 'Test7');
    assert.strictEqual(result, 10, 'Nested function calls');
    console.log('[PASS] Test 7: Nested functions');
}

// Test 8: Function with arguments - implicit Let
{
    const code = `
    Function Add(a As Integer, b As Integer)
        Add = a + b
    End Function

    Function Test8()
        Dim result As Integer
        result = Add(3, 7)
        Test8 = result
    End Function
    `;
    const result = runFunc(code, 'Test8');
    assert.strictEqual(result, 10, 'Function with arguments via implicit Let');
    console.log('[PASS] Test 8: Function with arguments');
}

// Test 9: VBA veteran pattern - sequential assignments
{
    const code = `
    Function GetValue()
        GetValue = 42
    End Function

    Function Test9()
        Dim x As Integer, y As Integer
        x = GetValue()
        y = x
        Test9 = x + y
    End Function
    `;
    const result = runFunc(code, 'Test9');
    assert.strictEqual(result, 84, 'Sequential assignment pattern');
    console.log('[PASS] Test 9: Sequential assignments');
}

// Test 10: Implicit Let in expressions
{
    const code = `
    Function GetValue()
        GetValue = 10
    End Function

    Function Test10()
        Test10 = GetValue() + 5
    End Function
    `;
    const result = runFunc(code, 'Test10');
    assert.strictEqual(result, 15, 'Function call in expression');
    console.log('[PASS] Test 10: Expression context');
}

// Test 11: String implicit assignment
{
    const code = `
    Function GetName()
        GetName = "Hello"
    End Function

    Function Test11()
        Dim name As String
        name = GetName()
        Test11 = name
    End Function
    `;
    const result = runFunc(code, 'Test11');
    assert.strictEqual(result, 'Hello', 'String implicit assignment');
    console.log('[PASS] Test 11: String assignment');
}

// Test 12: Boolean implicit assignment
{
    const code = `
    Function IsTrue()
        IsTrue = True
    End Function

    Function Test12()
        Dim flag As Boolean
        flag = IsTrue()
        If flag Then
            Test12 = 1
        Else
            Test12 = 0
        End If
    End Function
    `;
    const result = runFunc(code, 'Test12');
    assert.strictEqual(result, 1, 'Boolean implicit assignment');
    console.log('[PASS] Test 12: Boolean assignment');
}

// Test 13: Array/variant return
{
    const code = `
    Function GetArray()
        Dim arr(2) As Integer
        arr(0) = 1
        arr(1) = 2
        arr(2) = 3
        GetArray = arr
    End Function

    Function Test13()
        Dim result As Variant
        result = GetArray()
        Test13 = 1
    End Function
    `;
    const result = runFunc(code, 'Test13');
    assert.strictEqual(result, 1, 'Variant return value');
    console.log('[PASS] Test 13: Variant return');
}

// Test 14: Return value in If condition
{
    const code = `
    Function IsPositive(n As Integer)
        IsPositive = (n > 0)
    End Function

    Function Test14()
        If IsPositive(5) Then
            Test14 = 1
        Else
            Test14 = 0
        End If
    End Function
    `;
    const result = runFunc(code, 'Test14');
    assert.strictEqual(result, 1, 'Return value in condition');
    console.log('[PASS] Test 14: Conditional context');
}

// Test 15: Common pattern - helper function return
{
    const code = `
    Function Double(n As Integer)
        Double = n * 2
    End Function

    Function Test15()
        Dim values(2) As Integer
        values(0) = Double(5)
        values(1) = Double(10)
        Test15 = values(0) + values(1)
    End Function
    `;
    const result = runFunc(code, 'Test15');
    assert.strictEqual(result, 30, 'Common array pattern with helper function');
    console.log('[PASS] Test 15: Helper function pattern');
}

console.log('\n✅ Implicit Let: 全テスト通過');
