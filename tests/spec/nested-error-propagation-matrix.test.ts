/** ネスト手続き呼出しのErr伝播・Resume Next境界。 */
import { evalVBASingle, assert } from '../../test-libs/test-runner';

const ev = evalVBASingle(String.raw`
Sub RaiseFive()
    Err.Raise 5
End Sub

Sub RaiseSix()
    On Error GoTo LocalHandler
    Err.Raise 6
    Exit Sub
LocalHandler:
    Err.Raise 7
End Sub

Function ProbeCaller() As String
    Dim first As Long, second As Long
    On Error Resume Next
    RaiseFive
    first = Err.Number
    Err.Clear
    RaiseFive
    second = Err.Number
    ProbeCaller = CStr(first) & ":" & CStr(second)
End Function

Function ProbePropagation() As String
    On Error GoTo Handler
    RaiseSix
    Exit Function
Handler:
    ProbePropagation = CStr(Err.Number)
End Function
`);

assert.strictEqual(ev.callProcedure('ProbeCaller', []), '5:5',
    'Resume Next caller receives repeated nested procedure errors');
assert.strictEqual(ev.callProcedure('ProbePropagation', []), '7',
    'a nested handler re-raise propagates the replacement error');
console.log('[PASS] Nested error propagation matrix');
