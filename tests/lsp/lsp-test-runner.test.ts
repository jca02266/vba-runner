import { TestRunner, TestResult } from '../../src/lsp/test-runner';
import { assert } from '../../test-libs/test-runner';

const runner = new TestRunner(() => {});

// 1. Passing test is marked PASSED
{
    const code = `
    Sub Test_Simple()
        Dim x As Long
        x = 1 + 1
    End Sub
    `;
    const results = runner.runTests(code);
    assert.strictEqual(results.length, 1, 'one result');
    assert.strictEqual(results[0].state, 'passed', 'test passed');
    console.log('[PASS] Simple test passes');
}

// 2. Failing test (Err.Raise) is marked FAILED with correct message
{
    const code = `
    Sub Test_WillFail()
        On Error GoTo ErrHandler
        Err.Raise 1, , "Intentional failure"
        Exit Sub
    ErrHandler:
        Err.Raise Err.Number, , Err.Description
    End Sub
    `;
    const results = runner.runTests(code);
    assert.strictEqual(results.length, 1, 'one result');
    assert.strictEqual(results[0].state, 'failed', 'test failed');
    assert.strictEqual(results[0].message, 'Intentional failure', 'error message is correct');
    console.log('[PASS] Failing test: state=failed, message correct');
}

// 3. Multiple tests — each is evaluated independently via shared module scope
{
    const code = `
    Sub Test_First()
        Dim x As Long
        x = 1
    End Sub
    Sub Test_Second()
        Dim y As Long
        y = 2
    End Sub
    `;
    const results = runner.runTests(code);
    assert.strictEqual(results.length, 2, 'two results');
    assert.strictEqual(results[0].name, 'Test_First', 'first test name');
    assert.strictEqual(results[1].name, 'Test_Second', 'second test name');
    assert.strictEqual(results[0].state, 'passed', 'first passed');
    assert.strictEqual(results[1].state, 'passed', 'second passed');
    console.log('[PASS] Multiple tests executed');
}

// 4. Non-Test_ procedures are not run
{
    const code = `
    Sub Test_Real()
        Dim x As Long
        x = 1
    End Sub
    Sub Helper()
    End Sub
    `;
    const results = runner.runTests(code);
    assert.strictEqual(results.length, 1, 'only Test_ executed');
    assert.strictEqual(results[0].name, 'Test_Real', 'correct name');
    console.log('[PASS] Only Test_ procedures executed');
}

// 5. Empty source returns no results
{
    const results = runner.runTests('');
    assert.strictEqual(results.length, 0, 'no tests');
    console.log('[PASS] Empty source: no test results');
}

// 6. Test result has required fields with valid duration
{
    const code = 'Sub Test_Check()\nEnd Sub';
    const results = runner.runTests(code);
    const result = results[0];
    assert.ok(result.name, 'name present');
    assert.ok(result.state, 'state present');
    assert.ok(typeof result.duration === 'number', 'duration is number');
    assert.ok(result.duration >= 0, 'duration >= 0');
    const validStates = ['passed', 'failed', 'errored', 'skipped'];
    assert.ok(validStates.includes(result.state), 'state is valid');
    console.log('[PASS] Test result fields complete');
}

// 7. runTestWithEvaluation — error message is the VBA error message, not "[object Object]"
{
    const code = `
    Sub Test_WillFail()
        On Error GoTo ErrHandler
        Err.Raise 1, , "Intentional failure"
        Exit Sub
    ErrHandler:
        Err.Raise Err.Number, , Err.Description
    End Sub
    `;
    const result = runner.runTestWithEvaluation(code, 'Test_WillFail');
    assert.strictEqual(result.state, 'failed', 'state is failed');
    assert.strictEqual(result.message, 'Intentional failure', 'error message is VBA message, not [object Object]');
    console.log('[PASS] runTestWithEvaluation: error message is correct');
}

// 8. runTestWithEvaluation — not found returns failed with message
{
    const code = 'Sub Test_Other()\nEnd Sub';
    const result = runner.runTestWithEvaluation(code, 'Test_Missing');
    assert.strictEqual(result.state, 'failed', 'state is failed');
    assert.ok(result.message?.includes('Test_Missing'), 'message mentions procedure name');
    console.log('[PASS] runTestWithEvaluation: not-found handled');
}

console.log('\n✅ LSP Test Runner: 全テスト通過');
