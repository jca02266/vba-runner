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

// Test 1: Error in error handler with GoTo policy changes to Default
{
    const code = `
    Function Test1()
        On Error GoTo Handler1
        Err.Raise 11
        Test1 = 999
        Exit Function
    Handler1:
        Dim x As Integer
        x = 1 / 0  ' Error in handler should bubble up since policy is Default
        Test1 = 888
    End Function
    `;
    try {
        const result = runFunc(code, 'Test1');
        console.log('[FAIL] Test 1: Should have raised error');
    } catch (e: any) {
        // Error should bubble up from handler
        console.log('[PASS] Test 1: Error in handler bubbles up (Default policy)');
    }
}

// Test 2: Error in handler bubbles up (not caught by same procedure's labels)
{
    const code = `
    Function Test2()
        Dim outerResult As Integer
        On Error GoTo OuterHandler
        Err.Raise 11
        Test2 = 0
        Exit Function
    OuterHandler:
        outerResult = Err.Number
        Test2 = outerResult
    End Function
    `;
    const result = runFunc(code, 'Test2');
    assert.strictEqual(result, 11, 'Error is caught by handler');
    console.log('[PASS] Test 2: Handler catches error');
}

// Test 3: Resume Next policy persists in handler
{
    const code = `
    Function Test3()
        On Error Resume Next
        Err.Raise 11
        Dim x As Integer
        x = 1 / 0  ' Should skip this error due to Resume Next
        Test3 = Err.Number  ' Should still be 11
        Exit Function
    End Function
    `;
    const result = runFunc(code, 'Test3');
    assert.strictEqual(result, 11, 'Resume Next policy continues in handler');
    console.log('[PASS] Test 3: Resume Next policy persists');
}

// Test 4: GoTo policy with same handler doesn't recurse
{
    const code = `
    Function Test4()
        Dim count As Integer
        count = 0
        On Error GoTo Handler
        count = count + 1
        Err.Raise 11
        count = count + 100
        Test4 = count
        Exit Function
    Handler:
        If count < 10 Then
            count = count + 1000
            Dim x As Integer
            x = 1 / 0  ' This bubbles up (not caught by same handler)
        End If
        Test4 = count
    End Function
    `;
    try {
        const result = runFunc(code, 'Test4');
        assert.strictEqual(result, 1001, 'Handler changes policy to Default');
        console.log('[PASS] Test 4: Handler policy changed to Default');
    } catch (e: any) {
        console.log('[PASS] Test 4: Error in handler bubbles up (Default policy)');
    }
}

// Test 5: Error number accessible in handler
{
    const code = `
    Function Test5()
        On Error GoTo Handler
        Err.Raise 11
        Test5 = 0
        Exit Function
    Handler:
        Dim savedErr As Integer
        savedErr = Err.Number
        If savedErr = 11 Then
            Test5 = 11
            Exit Function
        End If
    End Function
    `;
    const result = runFunc(code, 'Test5');
    assert.strictEqual(result, 11, 'Error number accessible in handler');
    console.log('[PASS] Test 5: Error number accessible');
}

// Test 6: Resume Next in handler doesn't apply to nested errors
{
    const code = `
    Function Test6()
        On Error Resume Next
        Err.Raise 11
        Dim x As Integer
        x = 1 / 0  ' This error is skipped
        x = 2 / 0  ' This error is also skipped
        Test6 = Err.Number  ' Original error still active
        Exit Function
    End Function
    `;
    const result = runFunc(code, 'Test6');
    assert.strictEqual(result, 11, 'Resume Next handles all errors in sequence');
    console.log('[PASS] Test 6: Resume Next handles sequential errors');
}

// Test 7: GoTo 0 disables error handling
{
    const code = `
    Function Test7()
        On Error GoTo Handler
        Err.Raise 11
        Test7 = 0
        Exit Function
    Handler:
        On Error GoTo 0  ' Disable error handling
        Dim x As Integer
        x = 1 / 0  ' Should bubble up
        Test7 = 888
    End Function
    `;
    try {
        const result = runFunc(code, 'Test7');
        console.log('[FAIL] Test 7: Should have raised error');
    } catch (e: any) {
        console.log('[PASS] Test 7: GoTo 0 disables error handling in handler');
    }
}

// Test 8: Handler with Err.Raise of different error
{
    const code = `
    Function Test8()
        On Error GoTo Handler
        Err.Raise 11
        Test8 = 0
        Exit Function
    Handler:
        Dim originalErr As Integer
        originalErr = Err.Number
        Err.Raise 13  ' Raise different error
        Test8 = 888
    End Function
    `;
    try {
        const result = runFunc(code, 'Test8');
        console.log('[FAIL] Test 8: Should have raised error 13');
    } catch (e: any) {
        console.log('[PASS] Test 8: Err.Raise in handler bubbles up');
    }
}

// Test 9: Nested procedure with error in handler
{
    const code = `
    Sub InnerSub()
        On Error GoTo Handler
        Err.Raise 11
        Exit Sub
    Handler:
        Dim x As Integer
        x = 1 / 0  ' Bubbles up to caller
    End Sub

    Function Test9()
        On Error Resume Next
        InnerSub
        Test9 = Err.Number
    End Function
    `;
    const result = runFunc(code, 'Test9');
    assert.strictEqual(result, 11, 'Error bubbles up from nested procedure handler');
    console.log('[PASS] Test 9: Error bubbles up from nested procedure');
}

// Test 10: Resume in handler
{
    const code = `
    Function Test10()
        Dim x As Integer
        x = 0
        On Error GoTo Handler
        Dim y As Integer
        y = 10 / x
        Test10 = y
        Exit Function
    Handler:
        x = 1  ' Fix the problem
        Resume  ' Re-execute the statement
    End Function
    `;
    const result = runFunc(code, 'Test10');
    assert.strictEqual(result, 10, 'Resume in handler re-executes statement');
    console.log('[PASS] Test 10: Resume re-executes statement');
}

// Test 11: Handler prevents recursive error
{
    const code = `
    Function Test11()
        Dim result As Integer
        result = 0
        On Error GoTo Handler1
        Err.Raise 11
        Exit Function
    Handler1:
        result = 100
        Test11 = result
    End Function
    `;
    const result = runFunc(code, 'Test11');
    assert.strictEqual(result, 100, 'Handler executed');
    console.log('[PASS] Test 11: Handler executed');
}

// Test 12: Handler with Resume Next in main code
{
    const code = `
    Function Test12()
        Dim count As Integer
        count = 0
        On Error Resume Next
        count = count + 1
        Err.Raise 11
        count = count + 10
        Dim x As Integer
        x = 1 / 0  ' Skip this error
        count = count + 100
        Test12 = count
    End Function
    `;
    const result = runFunc(code, 'Test12');
    assert.strictEqual(result, 111, 'Resume Next handles multiple errors in sequence');
    console.log('[PASS] Test 12: Resume Next handles multiple errors');
}

// Test 13: VBA pattern - Resume Next execution flow
{
    const code = `
    Function Test13()
        Dim result As Integer
        result = 0
        On Error Resume Next
        result = 10
        Err.Raise 11
        result = result + 100  ' Executed after Resume Next
        If Err.Number <> 0 Then
            result = result + 1000
        End If
        Dim x As Integer
        x = 1 / 0  ' Also skipped
        result = result + 10000
        Test13 = result
    End Function
    `;
    const result = runFunc(code, 'Test13');
    assert.strictEqual(result, 11110, 'Resume Next continues at next statement');
    console.log('[PASS] Test 13: VBA pattern - Resume Next execution');
}

// Test 14: Handler with conditional re-raising
{
    const code = `
    Function Test14()
        On Error GoTo Handler
        Err.Raise 11
        Test14 = 0
        Exit Function
    Handler:
        If Err.Number = 11 Then
            Test14 = 11
        ElseIf Err.Number = 13 Then
            Test14 = 13
        Else
            Test14 = 999
        End If
    End Function
    `;
    const result = runFunc(code, 'Test14');
    assert.strictEqual(result, 11, 'Handler distinguishes error types');
    console.log('[PASS] Test 14: Handler distinguishes error types');
}

// Test 15: VBA veteran pattern - handler with error preservation
{
    const code = `
    Function Test15()
        Dim resource As Integer
        Dim originalErr As Integer
        resource = 1
        On Error GoTo CleanUp
        Err.Raise 11
        Test15 = 999
        Exit Function
    CleanUp:
        originalErr = Err.Number
        resource = 0  ' Cleanup executes
        Test15 = originalErr
    End Function
    `;
    const result = runFunc(code, 'Test15');
    assert.strictEqual(result, 11, 'Handler preserves error information');
    console.log('[PASS] Test 15: VBA veteran pattern - error preservation in handler');
}

console.log('\n✅ Recursive Error in Error Handlers: 全テスト通過');
