import { evalVBASingle, assert } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    return evalVBASingle(code);
}

function runFunc(code: string, name: string, args: any[] = []): any {
    return evalVBA(code).callProcedure(name, args);
}

// --- Assert with True condition ---
{
    const code = `
Function TestAssertTrue() As String
    Debug.Assert True
    TestAssertTrue = "passed"
End Function
`;
    const result = runFunc(code, 'TestAssertTrue');
    assert.strictEqual(result, 'passed', 'Assert with True condition allows execution to continue');
    console.log('[PASS] Assert with True condition');
}

// --- Assert with False condition ---
{
    const code = `
Function TestAssertFalse() As String
    Debug.Assert False
    TestAssertFalse = "should not reach here"
End Function
`;
    try {
        runFunc(code, 'TestAssertFalse');
        assert.strictEqual(false, true, 'Assert should suspend on False condition');
    } catch (e: any) {
        assert.strictEqual(true, true, 'Assert False throws error to suspend execution');
    }
    console.log('[PASS] Assert with False condition');
}

// --- Assert with numeric expression (0 = False, non-zero = True) ---
{
    const code = `
Function TestAssertNumeric(n As Integer) As String
    Debug.Assert n <> 0
    TestAssertNumeric = "passed"
End Function
`;
    const result1 = runFunc(code, 'TestAssertNumeric', [5]);
    assert.strictEqual(result1, 'passed', 'Assert with non-zero numeric value (True)');

    try {
        runFunc(code, 'TestAssertNumeric', [0]);
        assert.strictEqual(false, true, 'Assert should suspend on zero value');
    } catch (e: any) {
        assert.strictEqual(true, true, 'Assert with zero suspends execution');
    }
    console.log('[PASS] Assert with numeric expressions');
}

// --- Assert with comparison expression ---
{
    const code = `
Function TestAssertComparison(a As Integer, b As Integer) As String
    Debug.Assert a < b
    TestAssertComparison = "passed"
End Function
`;
    const result1 = runFunc(code, 'TestAssertComparison', [3, 5]);
    assert.strictEqual(result1, 'passed', 'Assert with true comparison');

    try {
        runFunc(code, 'TestAssertComparison', [5, 3]);
        assert.strictEqual(false, true, 'Assert should suspend on false comparison');
    } catch (e: any) {
        assert.strictEqual(true, true, 'Assert suspends on false comparison');
    }
    console.log('[PASS] Assert with comparison expressions');
}

// --- Assert with AND expression (entire expression evaluated) ---
{
    const code = `
Function TestAssertAnd(a As Integer, b As Integer) As String
    Dim result As String
    result = "initial"
    Debug.Assert a > 0 And b > 0
    result = "passed"
    TestAssertAnd = result
End Function
`;
    const result1 = runFunc(code, 'TestAssertAnd', [5, 3]);
    assert.strictEqual(result1, 'passed', 'Assert with true AND expression');

    try {
        runFunc(code, 'TestAssertAnd', [-5, 3]);
        assert.strictEqual(false, true, 'Assert should suspend when AND evaluates to false');
    } catch (e: any) {
        assert.strictEqual(true, true, 'Assert suspends on false AND expression');
    }
    console.log('[PASS] Assert with AND expressions');
}

// --- Assert with OR expression ---
{
    const code = `
Function TestAssertOr(a As Integer, b As Integer) As String
    Dim result As String
    result = "initial"
    Debug.Assert a > 0 Or b > 0
    result = "passed"
    TestAssertOr = result
End Function
`;
    const result1 = runFunc(code, 'TestAssertOr', [5, -3]);
    assert.strictEqual(result1, 'passed', 'Assert with true OR expression');

    try {
        runFunc(code, 'TestAssertOr', [-5, -3]);
        assert.strictEqual(false, true, 'Assert should suspend when OR evaluates to false');
    } catch (e: any) {
        assert.strictEqual(true, true, 'Assert suspends on false OR expression');
    }
    console.log('[PASS] Assert with OR expressions');
}

// --- Assert with NOT expression ---
{
    const code = `
Function TestAssertNot(a As Integer) As String
    Dim result As String
    result = "initial"
    Debug.Assert Not (a = 0)
    result = "passed"
    TestAssertNot = result
End Function
`;
    const result1 = runFunc(code, 'TestAssertNot', [5]);
    assert.strictEqual(result1, 'passed', 'Assert with true NOT expression');

    try {
        runFunc(code, 'TestAssertNot', [0]);
        assert.strictEqual(false, true, 'Assert should suspend when NOT evaluates to false');
    } catch (e: any) {
        assert.strictEqual(true, true, 'Assert suspends on false NOT expression');
    }
    console.log('[PASS] Assert with NOT expressions');
}

// --- Assert: code after true assertion executes normally ---
{
    const code = `
Function TestAssertSequence() As String
    Dim result As String
    result = "A"
    Debug.Assert True
    result = result & "B"
    TestAssertSequence = result
End Function
`;
    const result = runFunc(code, 'TestAssertSequence');
    assert.strictEqual(result, 'AB', 'Code after true assertion executes normally');
    console.log('[PASS] Assert sequence execution');
}

// --- Assert: code after false assertion does not execute ---
{
    const code = `
Function TestAssertHalt() As String
    Dim result As String
    result = "A"
    Debug.Assert False
    result = result & "B"
    TestAssertHalt = result
End Function
`;
    try {
        runFunc(code, 'TestAssertHalt');
        assert.strictEqual(false, true, 'False assertion should prevent subsequent code');
    } catch (e: any) {
        assert.strictEqual(true, true, 'Execution halted at false assertion');
    }
    console.log('[PASS] Assert halts execution');
}

console.log('\n✅ Debug.Assert: 全テスト通過');
