import { evalVBASingle, assert } from '../../test-libs/test-runner';

const ev = evalVBASingle(`
    Function TestOmittedLockStart() As Long
        Open "locks.dat" For Binary Access Read Write Shared As #1
        Lock #1, To 100
        On Error Resume Next
        Lock #1, 100 To 200
        TestOmittedLockStart = Err.Number
        Err.Clear
        On Error GoTo 0
        Unlock #1, To 100
        Close #1
    End Function
`);

assert.strictEqual(ev.callProcedure('TestOmittedLockStart', []), 75);

console.log('✅ Lock and Unlock accept an omitted range start');
