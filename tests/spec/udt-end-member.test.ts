import { evalVBASingle, assert } from '../../test-libs/test-runner';

const ev = evalVBASingle(`
    Type Interval
        End As Long
    End Type

    Function TestEndMember() As Long
        Dim value As Interval
        value.End = 42
        TestEndMember = value.End
    End Function
`);

assert.strictEqual(ev.callProcedure('TestEndMember', []), 42);

console.log('✅ UDT members may use End as their name');
