import { strict as assert } from 'node:assert';
import { evalVBASingle } from '../../test-libs/test-runner';

const source = String.raw`
Public Function Probe() As String
    Dim fso As Object, stream As Object, errNo As Long
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    Err.Clear
    Set stream = fso.CreateTextFile("C:\\overwrite.txt", True, False)
    stream.Close
    Set stream = fso.CreateTextFile("C:\\overwrite.txt", False, False)
    errNo = Err.Number
    Probe = "ERR=" & CStr(errNo)
    On Error GoTo 0
End Function
`;

const ev = evalVBASingle(source);
assert.equal(ev.callProcedure('Probe', []), 'ERR=58');
console.log('[PASS] FSO CreateTextFile overwrite boundary (1 case)');
