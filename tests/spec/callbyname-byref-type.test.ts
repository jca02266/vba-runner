import { strict as assert } from 'node:assert';
import { evalVBASingle } from '../../test-libs/test-runner';

const ev = evalVBASingle(String.raw`
Class Worker
    Public Sub TakeLong(ByRef value As Long)
        value = value + 1
    End Sub
End Class
Function Probe() As String
    Dim worker As New Worker, value As Integer
    value = 10
    On Error Resume Next
    Err.Clear: CallByName worker, "TakeLong", VbMethod, value
    Probe = CStr(Err.Number) & ":" & CStr(value)
End Function`);

assert.equal(ev.callProcedure('Probe', []), '13:10');
console.log('[PASS] CallByName typed ByRef boundary runner behavior');
