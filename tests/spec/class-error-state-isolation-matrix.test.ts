/** クラス局所On Errorと呼出し元Err状態の分離。 */
import { evalVBAModules, assert } from '../../test-libs/test-runner';

const ev = evalVBAModules([{
    name: 'Module1',
    code: String.raw`
Class Worker
    Public Function Recover() As String
        Dim localErr As Long
        On Error Resume Next
        Err.Raise 5
        localErr = Err.Number
        Recover = CStr(localErr) & ":" & CStr(Err.Number)
    End Function

    Public Function Rethrow() As Long
        On Error GoTo Handler
        Err.Raise 6
        Exit Function
Handler:
        Err.Raise 7
    End Function
End Class

Public Function ProbeLocal() As String
    Dim worker As New Worker
    On Error GoTo Handler
    ProbeLocal = worker.Recover()
    Exit Function
Handler:
    ProbeLocal = "H" & CStr(Err.Number)
End Function

Public Function ProbeRethrow() As String
    Dim worker As New Worker
    On Error GoTo Handler
    worker.Rethrow
    Exit Function
Handler:
    ProbeRethrow = "H" & CStr(Err.Number)
End Function
`,
}]);

assert.strictEqual(ev.callProcedure('ProbeLocal', []), '5:5',
    'class-local Resume Next handles its own error without caller handler');
assert.strictEqual(ev.callProcedure('ProbeRethrow', []), 'H7',
    'class handler re-raise reaches the caller handler');
console.log('[PASS] Class error-state isolation matrix');
