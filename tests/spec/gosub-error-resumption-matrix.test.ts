/** GoSub/ReturnとOn Error復帰境界。 */
import { evalVBASingle, assert } from '../../test-libs/test-runner';

const ev = evalVBASingle(String.raw`
Function ProbeNormal() As String
    Dim value As Long
    value = 1
    GoSub Worker
    ProbeNormal = CStr(value)
    Exit Function
Worker:
    value = value + 2
    Return
End Function

Function ProbeError() As String
    On Error GoTo Handler
    GoSub Worker
    ProbeError = "bad"
    Exit Function
Worker:
    Err.Raise 5
    Return
Handler:
    ProbeError = "H" & CStr(Err.Number)
End Function

Function ProbeResume() As String
    Dim value As Long
    On Error Resume Next
    GoSub Worker
    value = value + 1
    ProbeResume = CStr(Err.Number) & ":" & CStr(value)
    Exit Function
Worker:
    Err.Raise 6
    Return
End Function
`);

assert.strictEqual(ev.callProcedure('ProbeNormal', []), '3',
    'GoSub and Return resume the caller at the next statement');
assert.strictEqual(ev.callProcedure('ProbeError', []), 'H5',
    'GoSub errors reach the caller On Error handler');
assert.strictEqual(ev.callProcedure('ProbeResume', []), '6:1',
    'Resume Next handles an error raised inside a GoSub body');
console.log('[PASS] GoSub error resumption matrix');
