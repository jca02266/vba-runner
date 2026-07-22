import { evalVBASingle, assert } from '../../test-libs/test-runner';

const ev = evalVBASingle(`
    Function TestNamedLabel() As Long
        Go To done
        TestNamedLabel = 1
    done:
        TestNamedLabel = 42
    End Function

    Function TestNumericLabel() As Long
        Go To 100
        TestNumericLabel = 1
    100 TestNumericLabel = 84
    End Function
`);

assert.strictEqual(ev.callProcedure('TestNamedLabel', []), 42);
assert.strictEqual(ev.callProcedure('TestNumericLabel', []), 84);

console.log('✅ Two-word Go To statements jump to named and numeric labels');
