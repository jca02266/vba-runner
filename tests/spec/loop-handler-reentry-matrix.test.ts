/** For/Doループ内のエラー処理とResume再入防止。 */
import { evalVBASingle, assert } from '../../test-libs/test-runner';

const ev = evalVBASingle(String.raw`
Function ProbeFor() As String
    Dim i As Long, errors As Long, visits As Long
    On Error GoTo Handler
    For i = 1 To 3
        visits = visits + 1
        If i = 2 Then Err.Raise 5
ContinueFor:
    Next i
    ProbeFor = CStr(visits) & ":" & CStr(errors)
    Exit Function
Handler:
    errors = errors + 1
    Resume ContinueFor
End Function

Function ProbeDo() As String
    Dim i As Long, errors As Long
    i = 0
    On Error GoTo Handler
LoopStart:
    i = i + 1
    If i < 3 Then Err.Raise 6
    ProbeDo = CStr(i) & ":" & CStr(errors)
    Exit Function
Handler:
    errors = errors + 1
    Resume LoopStart
End Function
`);

assert.strictEqual(ev.callProcedure('ProbeFor', []), '3:1',
    'For loop resumes at its continuation without re-entering the handler');
assert.strictEqual(ev.callProcedure('ProbeDo', []), '3:2',
    'Do-style loop resumes repeatedly without infinite handler re-entry');
console.log('[PASS] Loop error-handler re-entry matrix');
