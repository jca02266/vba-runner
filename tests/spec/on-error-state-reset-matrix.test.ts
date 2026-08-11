/** On Error状態解除・Resume・ハンドラ再設定境界。 */
import { evalVBASingle, assert } from '../../test-libs/test-runner';

const ev = evalVBASingle(String.raw`
Function ProbeReset() As String
    Dim first As Long
    On Error Resume Next
    Err.Raise 5
    first = Err.Number
    On Error GoTo 0
    On Error GoTo Handler
    Err.Raise 6
    Exit Function
Handler:
    ProbeReset = CStr(first) & ":" & CStr(Err.Number)
End Function

Function ProbeResume() As String
    Dim value As Long
    On Error GoTo Handler
    value = 1
    Err.Raise 7
    value = 99
Done:
    ProbeResume = CStr(value)
    Exit Function
Handler:
    Resume Done
End Function

Function ProbeReenable() As String
    Dim first As Long, second As Long
    On Error Resume Next
    Err.Raise 8
    first = Err.Number
    On Error GoTo Handler
    Err.Raise 9
    Exit Function
Handler:
    second = Err.Number
    ProbeReenable = CStr(first) & ":" & CStr(second)
End Function
`);

assert.strictEqual(ev.callProcedure('ProbeReset', []), '5:6',
    'GoTo 0 followed by a new handler resets and rehandles errors');
assert.strictEqual(ev.callProcedure('ProbeResume', []), '1',
    'Resume jumps to the requested label without executing the failed path');
assert.strictEqual(ev.callProcedure('ProbeReenable', []), '8:9',
    'A GoTo handler can be re-enabled after Resume Next');
console.log('[PASS] On Error state reset and Resume matrix');
