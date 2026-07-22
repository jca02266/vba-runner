import { evalVBASingle, assert } from '../../test-libs/test-runner';

const ev = evalVBASingle(`
    Function TestThenBranch() As Long
        If True Then 100
        TestThenBranch = 1
        Exit Function
    100 TestThenBranch = 42
    End Function

    Function TestElseBranch() As Long
        If False Then 100 Else 200
        TestElseBranch = 1
        Exit Function
    100 TestElseBranch = 42
        Exit Function
    200 TestElseBranch = 84
    End Function
`);

assert.strictEqual(ev.callProcedure('TestThenBranch', []), 42);
assert.strictEqual(ev.callProcedure('TestElseBranch', []), 84);

console.log('✅ If Then line-number branches jump to their target labels');
