import { assert, evalVBASingle } from '../../test-libs/test-runner';

const ev = evalVBASingle(`
    Function ProbeEmptyGet() As Long
        Dim b As Byte, result As Long
        Open "empty-random-eof.dat" For Random As #1 Len = 1
        On Error Resume Next
        Get #1, , b
        result = Err.Number
        If EOF(1) Then result = result + 100
        Close #1

        Err.Clear
        Open "empty-binary-eof.dat" For Binary As #1
        Get #1, , b
        result = result + Err.Number
        If EOF(1) Then result = result + 100
        Close #1
        ProbeEmptyGet = result
    End Function
`);

assert.strictEqual(ev.callProcedure('ProbeEmptyGet', []), 324,
    'empty Random/Binary Get raises Error 62 and sets EOF');
console.log('[PASS] empty Random/Binary Get reaches EOF');
