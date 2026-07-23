import { evalVBASingle, assert } from '../../test-libs/test-runner';

const ev = evalVBASingle(`
    Function Increment(ByRef value As Long) As Long
        value = value + 1
        Increment = value
    End Function

    Function TestCallByVal() As Long
        Dim value As Long
        value = 1
        TestCallByVal = Increment(ByVal value) * 10 + value
    End Function
`);

assert.strictEqual(ev.callProcedure('TestCallByVal', []), 21);

console.log('✅ Call-site ByVal suppresses ByRef writeback');
