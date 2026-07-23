import { evalVBASingle, assert } from '../../test-libs/test-runner';

const ev = evalVBASingle(`
    Function DefaultValue(ByVal Optional value As Integer = 42) As Long
        DefaultValue = value
    End Function
`);

assert.strictEqual(ev.callProcedure('DefaultValue', []), 42);
assert.strictEqual(ev.callProcedure('DefaultValue', [7]), 7);

console.log('✅ ByVal Optional parameters are accepted');
