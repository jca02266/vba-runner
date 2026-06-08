import { evalVBASingle, assert } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    return evalVBASingle(code);
}

function runFunc(code: string, name: string, args: any[] = []): any {
    return evalVBA(code).callProcedure(name, args);
}

// Test 1: Resume re-executes the error-causing statement (division by zero fixed)
{
    const code = `
    Function Test1()
        Dim x As Integer
        x = 0
        On Error GoTo ErrorHandler
        Dim y As Integer
        y = 10 / x
        Test1 = y
        Exit Function
    ErrorHandler:
        x = 1  ' Fix the problem
        Resume
    End Function
    `;
    const result = runFunc(code, 'Test1');
    assert.strictEqual(result, 10, 'Resume should re-execute division after x is fixed');
    console.log('[PASS] Test 1: Resume re-executes error statement');
}

// Test 2: Resume Next skips to the next statement
{
    const code = `
    Function Test2()
        Dim result As Integer
        result = 0
        On Error GoTo ErrorHandler
        result = result + 1
        Err.Raise 11
        result = result + 10
        result = result + 100
        Test2 = result
        Exit Function
    ErrorHandler:
        Resume Next
    End Function
    `;
    const result = runFunc(code, 'Test2');
    assert.strictEqual(result, 111, 'Resume Next skips error statement, executes following statements');
    console.log('[PASS] Test 2: Resume Next skips error statement');
}

// Test 3: Resume with label jumps to the label
{
    const code = `
    Function Test3()
        Dim result As Integer
        result = 0
        On Error GoTo ErrorHandler
        Err.Raise 11
        result = 5
        Exit Function
    ErrorHandler:
        Resume SkipLabel
    SkipLabel:
        result = 10
        Test3 = result
    End Function
    `;
    const result = runFunc(code, 'Test3');
    assert.strictEqual(result, 10, 'Resume <label> should jump to the label');
    console.log('[PASS] Test 3: Resume <label> jumps to label');
}

// Test 4: Err object is reset after Resume
{
    const code = `
    Function Test4()
        On Error GoTo ErrorHandler
        Err.Raise 11
        Exit Function
    ErrorHandler:
        Dim beforeResume As Integer
        beforeResume = Err.Number
        Resume AfterLabel
    AfterLabel:
        Dim afterResume As Integer
        afterResume = Err.Number
        If beforeResume = 11 And afterResume = 0 Then
            Test4 = 1
        Else
            Test4 = 0
        End If
    End Function
    `;
    const result = runFunc(code, 'Test4');
    assert.strictEqual(result, 1, 'Err.Number should be 11 before Resume, 0 after Resume');
    console.log('[PASS] Test 4: Err object reset after Resume');
}

// Test 5: Resume without error raises error 20
{
    const code = `
    Function Test5()
        On Error GoTo Handler
        Resume
        Test5 = 0
        Exit Function
    Handler:
        Test5 = Err.Number
    End Function
    `;
    const result = runFunc(code, 'Test5');
    assert.strictEqual(result, 20, 'Resume without error should raise error 20');
    console.log('[PASS] Test 5: Resume without error raises error 20');
}

// Test 6: Resume with state variable prevents infinite retry
{
    const code = `
    Function Test6()
        Dim x As Integer
        Dim retried As Boolean
        x = 0
        retried = False
        On Error GoTo ErrorHandler
        Dim y As Integer
        y = 10 / x
        Test6 = y
        Exit Function
    ErrorHandler:
        If Not retried Then
            retried = True
            x = 2
            Resume
        End If
        Test6 = 999
    End Function
    `;
    const result = runFunc(code, 'Test6');
    assert.strictEqual(result, 5, 'Resume with retry flag should prevent infinite loop');
    console.log('[PASS] Test 6: Resume with state variable');
}

// Test 7: Resume Next with conditional
{
    const code = `
    Function Test7()
        Dim result As Integer
        result = 0
        On Error GoTo Handler
        result = 1
        Err.Raise 11
        result = result + 10
        result = result + 100
        Test7 = result
        Exit Function
    Handler:
        If Err.Number = 11 Then
            Resume Next
        End If
    End Function
    `;
    const result = runFunc(code, 'Test7');
    assert.strictEqual(result, 111, 'Resume Next in conditional should skip error statement');
    console.log('[PASS] Test 7: Resume Next in conditional');
}

// Test 8: Resume with label in conditional
{
    const code = `
    Function Test8()
        Dim result As Integer
        result = 0
        On Error GoTo Handler
        Err.Raise 11
        Exit Function
    Handler:
        If Err.Number = 11 Then
            result = 5
            Resume SkipLabel
        End If
    SkipLabel:
        result = result + 10
        Test8 = result
    End Function
    `;
    const result = runFunc(code, 'Test8');
    assert.strictEqual(result, 15, 'Resume <label> in conditional should jump correctly');
    console.log('[PASS] Test 8: Resume <label> in conditional');
}

// Test 9: Resume maintains variable state
{
    const code = `
    Function Test9()
        Dim x As Integer
        Dim y As Integer
        x = 10
        y = 0
        On Error GoTo Handler
        y = x + 5
        Dim z As Integer
        z = 1 / (x - 10)
        y = y + 100
        Test9 = x + y
        Exit Function
    Handler:
        x = x + 1
        Resume
    End Function
    `;
    const result = runFunc(code, 'Test9');
    assert.strictEqual(result, 126, 'Resume should maintain variable state (x=11, y=15+100=115, total=126)');
    console.log('[PASS] Test 9: Resume maintains variable state');
}

// Test 10: Resume Next skips multiple statements in sequence
{
    const code = `
    Function Test10()
        Dim result As Integer
        result = 0
        On Error GoTo Handler
        result = 1
        Err.Raise 11
        result = result + 10
        result = result + 100
        result = result + 1000
        Test10 = result
        Exit Function
    Handler:
        Resume Next
    End Function
    `;
    const result = runFunc(code, 'Test10');
    assert.strictEqual(result, 1111, 'Resume Next should skip error and execute all following statements');
    console.log('[PASS] Test 10: Resume Next skips to next statement');
}

// Test 11: Resume changes flow to label
{
    const code = `
    Function Test11()
        Dim result As Integer
        result = 0
        On Error GoTo Handler
        result = 1
        Err.Raise 11
        result = 100
        Exit Function
    Handler:
        Resume SkipLabel
    SkipLabel:
        result = 50
        Test11 = result
    End Function
    `;
    const result = runFunc(code, 'Test11');
    assert.strictEqual(result, 50, 'Resume <label> should bypass error recovery path');
    console.log('[PASS] Test 11: Resume <label> changes execution path');
}

// Test 12: Resume in nested procedure
{
    const code = `
    Sub ErrorSub(ByRef e As Integer)
        On Error GoTo Handler
        Err.Raise 11
        e = 0
        Exit Sub
    Handler:
        e = Err.Number
        Resume Next
        e = 100  ' This should be executed after Resume Next
    End Sub

    Function Test12()
        Dim errNum As Integer
        errNum = 0
        ErrorSub errNum
        Test12 = errNum
    End Function
    `;
    const result = runFunc(code, 'Test12');
    assert.strictEqual(result, 0, 'Resume Next in subroutine resumes at next statement (e = 0)');
    console.log('[PASS] Test 12: Resume in nested procedure');
}

// Test 13: Resume with retry counter
{
    const code = `
    Function Test13()
        Dim attempts As Integer
        Dim x As Integer
        attempts = 0
        x = 0
        On Error GoTo Handler
        attempts = attempts + 1
        Dim y As Integer
        y = 100 / x
        Test13 = y
        Exit Function
    Handler:
        If attempts < 2 Then
            x = 10
            Resume
        Else
            Test13 = 999
        End If
    End Function
    `;
    const result = runFunc(code, 'Test13');
    assert.strictEqual(result, 10, 'Resume with retry counter should work');
    console.log('[PASS] Test 13: Resume with retry counter');
}

// Test 14: Resume with multiple labels
{
    const code = `
    Function Test14()
        Dim result As Integer
        result = 0
        On Error GoTo Handler
        result = 1
        Err.Raise 11
        Exit Function
    Handler:
        result = result + 10
        Resume LabelB
    LabelB:
        result = result + 100
        Test14 = result
    End Function
    `;
    const result = runFunc(code, 'Test14');
    assert.strictEqual(result, 111, 'Resume should jump to correct label');
    console.log('[PASS] Test 14: Resume with multiple labels');
}

// Test 15: VBA veteran pattern - division by zero handling with simple retry
{
    const code = `
    Function Test15()
        Dim attempt As Integer
        Dim x As Integer
        Dim y As Integer
        attempt = 0
        x = 0
        On Error GoTo RetryHandler
        attempt = attempt + 1
        y = 1000 / x
        Test15 = y
        Exit Function
    RetryHandler:
        If attempt < 2 Then
            x = attempt
            Resume
        Else
            Test15 = 999
        End If
    End Function
    `;
    const result = runFunc(code, 'Test15');
    assert.strictEqual(result, 1000, 'Resume pattern with retry should work (1000 / 1 on first retry)');
    console.log('[PASS] Test 15: VBA veteran pattern - retry with simple increment');
}

console.log('\n✅ Resume Statement Target Determination: 全テスト通過');
