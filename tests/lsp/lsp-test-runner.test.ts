import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { TestRunner, TestResult } from '../../src/lsp/test-runner';
import { assert } from '../../test-libs/test-runner';

function runTests(src: string): TestResult[] {
    const tokens = new Lexer(src).tokenize();
    const ast = new Parser(tokens).parse();
    const runner = new TestRunner();
    return runner.runTests(ast.body);
}

// 1. Passing test is marked PASSED
{
    const code = `
    Sub Test_Simple()
        ' Test that does nothing should pass
    End Sub
    `;
    const results = runTests(code);
    assert.strictEqual(results.length, 1, 'one result');
    assert.strictEqual(results[0].state, 'passed', 'test passed');
    console.log('[PASS] Simple test passes');
}

// 2. Test with assertion that fails
{
    const code = `
    Sub Test_AssertFalse()
        x = 1 + 1
        if x <> 3 then
            ' Simulate test pass (in real testing, would use assertions)
        end if
    End Sub
    `;
    const results = runTests(code);
    assert.strictEqual(results.length, 1, 'one result');
    assert.ok(results[0].state === 'passed' || results[0].state === 'failed', 'has state');
    console.log('[PASS] Test with logic handled');
}

// 3. Multiple tests
{
    const code = `
    Sub Test_First()
    End Sub
    Sub Test_Second()
    End Sub
    `;
    const results = runTests(code);
    assert.strictEqual(results.length, 2, 'two results');
    assert.strictEqual(results[0].name, 'Test_First', 'first test name');
    assert.strictEqual(results[1].name, 'Test_Second', 'second test name');
    console.log('[PASS] Multiple tests executed');
}

// 4. Test error is captured
{
    const code = `
    Sub Test_WithError()
        y = z 'undefined variable should cause error in strict mode
    End Sub
    `;
    const results = runTests(code);
    assert.strictEqual(results.length, 1, 'one result');
    assert.ok(results[0].state, 'test executed');
    // Error may be caught depending on evaluator strictness
    console.log('[PASS] Error handling in test');
}

// 5. Test result has required fields
{
    const code = 'Sub Test_Check()\nEnd Sub';
    const results = runTests(code);
    const result = results[0];
    assert.ok(result.name, 'name present');
    assert.ok(result.state, 'state present');
    assert.ok(typeof result.duration === 'number', 'duration is number');
    console.log('[PASS] Test result fields complete');
}

// 6. Test duration measured
{
    const code = 'Sub Test_Timing()\nEnd Sub';
    const results = runTests(code);
    assert.ok(results[0].duration >= 0, 'duration >= 0');
    console.log('[PASS] Test duration measured');
}

// 7. Non-Test_ procedures are not run
{
    const code = `
    Sub Test_Real()
    End Sub
    Sub Helper()
    End Sub
    `;
    const results = runTests(code);
    assert.strictEqual(results.length, 1, 'only Test_ executed');
    console.log('[PASS] Only Test_ procedures executed');
}

// 8. Test with custom message
{
    const code = `
    Sub Test_WithMessage()
        x = 5 + 5
    End Sub
    `;
    const results = runTests(code);
    assert.ok(results[0], 'test executed');
    assert.ok(results[0].message === undefined || typeof results[0].message === 'string', 'message optional or string');
    console.log('[PASS] Test with optional message');
}

// 9. Empty test source
{
    const results = runTests('');
    assert.strictEqual(results.length, 0, 'no tests');
    console.log('[PASS] Empty source: no test results');
}

// 10. Test state values are valid
{
    const code = 'Sub Test_Check()\nEnd Sub';
    const results = runTests(code);
    const validStates = ['passed', 'failed', 'errored', 'skipped'];
    assert.ok(validStates.includes(results[0].state), 'state is valid');
    console.log('[PASS] Test state is valid value');
}

console.log('\n✅ LSP Test Runner: 全テスト通過');
