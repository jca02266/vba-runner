import { evalVBASingle, assert } from '../../test-libs/test-runner';

const ev = evalVBASingle(`
    Function TestNamedLabel() As Long
        Go Sub worker
        TestNamedLabel = TestNamedLabel + 1
        Exit Function
    worker:
        TestNamedLabel = 41
        Return
    End Function

    Function TestNumericLabel() As Long
        Go Sub 100
        TestNumericLabel = TestNumericLabel + 2
        Exit Function
    100 TestNumericLabel = 40
        Return
    End Function
`);

assert.strictEqual(ev.callProcedure('TestNamedLabel', []), 42);
assert.strictEqual(ev.callProcedure('TestNumericLabel', []), 42);

console.log('✅ Two-word Go Sub statements return from named and numeric labels');
