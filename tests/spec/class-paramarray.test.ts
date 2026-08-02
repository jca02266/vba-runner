import { evalVBASingle, assert } from '../../test-libs/test-runner';

// Class members must pack every argument at and after ParamArray into the
// Variant array, just like standard-module procedures do.
const code = `
Class Box
    Public Function Sum(ParamArray values()) As Long
        Sum = values(0) + values(1)
    End Function

    Public Sub Increment(ParamArray values())
        values(0) = values(0) + 5
    End Sub
End Class

Function Probe() As String
    Dim box As New Box, value As Long
    value = 7
    box.Increment value
    Probe = CStr(value) & "|" & CStr(box.Sum(1, 2))
End Function

Function ProbeCallByName() As String
    Dim box As New Box
    On Error Resume Next
    ProbeCallByName = CStr(CallByName(box, "Sum", VbMethod, 1, 2)) & "|" & CStr(Err.Number)
End Function
`;

const ev = evalVBASingle(code);
assert.strictEqual(ev.callProcedure('Probe', []), '12|3',
    'Class ParamArray packs multiple arguments and preserves ByRef writeback');
assert.strictEqual(ev.callProcedure('ProbeCallByName', []), '3|0',
    'CallByName class method packs ParamArray arguments');
console.log('[PASS] Class member ParamArray packing and ByRef');
