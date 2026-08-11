import { strict as assert } from 'node:assert';
import { evalVBASingle } from '../../test-libs/test-runner';

const source = String.raw`
Public Function Probe() As String
    Dim fso As Object, stream As Object, errNo As Long
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    Err.Clear
    Set stream = fso.OpenTextFile("C:\\missing-open.txt", 1, False, -2)
    errNo = Err.Number
    Probe = "ERR=" & CStr(errNo)
    On Error GoTo 0
End Function
`;
const ev = evalVBASingle(source);
assert.equal(ev.callProcedure('Probe', []), 'ERR=53');
console.log('[PASS] FSO OpenTextFile missing-file boundary (1 case)');
