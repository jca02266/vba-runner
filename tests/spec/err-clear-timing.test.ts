import { evalVBASingle, assert } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    return evalVBASingle(code);
}

function runFunc(code: string, name: string, args: any[] = []): any {
    return evalVBA(code).callProcedure(name, args);
}

// Test 1: Err object persists after On Error Resume Next until error
{
    const code = `
    Function Test1()
        On Error Resume Next
        Err.Raise 11  ' Division by zero error
        ' After error, Err should have data
        Test1 = Err.Number
    End Function
    `;
    const result = runFunc(code, 'Test1');
    assert.strictEqual(result, 11, 'Err.Number should be set after error');
    console.log('[PASS] Test 1: Err persists after error');
}

// Test 2: Err persists in multiple statements after error
{
    const code = `
    Function Test2()
        On Error Resume Next
        Err.Raise 13  ' Type mismatch
        Dim x As Integer
        x = 1  ' Normal statement
        Test2 = Err.Number
    End Function
    `;
    const result = runFunc(code, 'Test2');
    assert.strictEqual(result, 13, 'Err should persist through normal statements');
    console.log('[PASS] Test 2: Err persists through normal statements');
}

// Test 3: Err is reset by Err.Clear method
{
    const code = `
    Function Test3()
        On Error Resume Next
        Err.Raise 11
        Err.Clear
        Test3 = Err.Number
    End Function
    `;
    const result = runFunc(code, 'Test3');
    assert.strictEqual(result, 0, 'Err.Clear should reset error number');
    console.log('[PASS] Test 3: Err.Clear resets error');
}

// Test 4: New error replaces previous error
{
    const code = `
    Function Test4()
        On Error Resume Next
        Err.Raise 11
        Err.Raise 13
        Test4 = Err.Number
    End Function
    `;
    const result = runFunc(code, 'Test4');
    assert.strictEqual(result, 13, 'New error should replace previous error');
    console.log('[PASS] Test 4: New error replaces old error');
}

// Test 5: Err persists when no new error occurs
{
    const code = `
    Function Test5()
        On Error Resume Next
        Err.Raise 11
        Dim x As Integer
        x = 5
        Dim y As Integer
        y = x + 3
        Test5 = Err.Number
    End Function
    `;
    const result = runFunc(code, 'Test5');
    assert.strictEqual(result, 11, 'Err should persist when no new error');
    console.log('[PASS] Test 5: Err persists when no new error');
}

// Test 6: Check Err.Description after error
{
    const code = `
    Function Test6()
        On Error Resume Next
        Err.Raise 11, , "Division by zero"
        Test6 = Len(Err.Description) > 0
    End Function
    `;
    const result = runFunc(code, 'Test6');
    assert.strictEqual(result, -1, 'Err.Description should be set (VBA true = -1)');
    console.log('[PASS] Test 6: Err.Description is set');
}

// Test 7: Err.Source property
{
    const code = `
    Function Test7()
        On Error Resume Next
        Err.Raise 11, "TestSource"
        Test7 = Len(Err.Source) > 0
    End Function
    `;
    const result = runFunc(code, 'Test7');
    assert.strictEqual(result, -1, 'Err.Source should be set');
    console.log('[PASS] Test 7: Err.Source is set');
}

// Test 8: Multiple error handlers in sequence
{
    const code = `
    Function Test8()
        On Error Resume Next
        Err.Raise 11
        Dim firstErr As Integer
        firstErr = Err.Number
        On Error GoTo 0
        On Error Resume Next
        Err.Raise 13
        Test8 = Err.Number
    End Function
    `;
    const result = runFunc(code, 'Test8');
    assert.strictEqual(result, 13, 'New error handler should set new error');
    console.log('[PASS] Test 8: Multiple error handlers');
}

// Test 9: Err after successful operation with previous error
{
    const code = `
    Function Test9()
        On Error Resume Next
        Err.Raise 11
        ' Now do a successful operation that might normally cause an error
        Dim x As Integer
        x = 10 / 2  ' This works fine
        Test9 = Err.Number  ' Should still be 11
    End Function
    `;
    const result = runFunc(code, 'Test9');
    assert.strictEqual(result, 11, 'Err should persist through successful operation');
    console.log('[PASS] Test 9: Err persists through successful operation');
}

// Test 10: Err values before any error
{
    const code = `
    Function Test10()
        On Error Resume Next
        Test10 = Err.Number
    End Function
    `;
    const result = runFunc(code, 'Test10');
    assert.strictEqual(result, 0, 'Err.Number should be 0 initially');
    console.log('[PASS] Test 10: Err.Number is 0 initially');
}

// Test 11: Checking Err in conditional after error
{
    const code = `
    Function Test11()
        On Error Resume Next
        Err.Raise 11
        If Err.Number <> 0 Then
            Test11 = Err.Number
        Else
            Test11 = 0
        End If
    End Function
    `;
    const result = runFunc(code, 'Test11');
    assert.strictEqual(result, 11, 'Err.Number should be accessible in conditional');
    console.log('[PASS] Test 11: Err accessible in conditional');
}

// Test 12: Err state after Resume Next (if handler exits normally)
{
    const code = `
    Sub ErrorSub(ByRef e As Integer)
        On Error Resume Next
        Err.Raise 11
        e = Err.Number
    End Sub

    Function Test12()
        Dim errNum As Integer
        ErrorSub errNum
        Test12 = errNum
    End Function
    `;
    const result = runFunc(code, 'Test12');
    assert.strictEqual(result, 11, 'Err should be captured from subroutine');
    console.log('[PASS] Test 12: Err captured from subroutine');
}

// Test 13: VBA veteran pattern - error handling with tracking
{
    const code = `
    Function Test13()
        Dim attempts As Integer
        attempts = 0
        On Error Resume Next
        Err.Raise 11
        attempts = attempts + 1  ' Increment after error
        If Err.Number <> 0 Then
            ' Error is still set
            attempts = attempts + 1
        End If
        Test13 = attempts
    End Function
    `;
    const result = runFunc(code, 'Test13');
    assert.strictEqual(result, 2, 'Error tracking should work correctly');
    console.log('[PASS] Test 13: VBA veteran pattern - error tracking');
}

// Test 14: Err.Clear vs new error
{
    const code = `
    Function Test14()
        On Error Resume Next
        Err.Raise 11
        Err.Clear
        Err.Raise 13
        Test14 = Err.Number
    End Function
    `;
    const result = runFunc(code, 'Test14');
    assert.strictEqual(result, 13, 'New error after Clear should be set');
    console.log('[PASS] Test 14: New error after Clear');
}

// Test 15: Consecutive Err.Clear calls
{
    const code = `
    Function Test15()
        On Error Resume Next
        Err.Raise 11
        Err.Clear
        Err.Clear
        Test15 = Err.Number
    End Function
    `;
    const result = runFunc(code, 'Test15');
    assert.strictEqual(result, 0, 'Multiple Clear calls should be safe');
    console.log('[PASS] Test 15: Multiple Err.Clear calls');
}

console.log('\n✅ Err.Clear Timing in Error Handlers: 全テスト通過');
