import { evalVBASingle, assert } from '../../test-libs/test-runner';

// A dynamic array member passed ByRef must be rebound when the callee ReDims it.
// The same operation on a local array is the reference behavior; MemberExpression
// writeback must not mistake the replacement for an illegal whole-array Let.
const code = String.raw`
Class Holder
    Public Values() As Long
End Class

Type Box
    Values() As Long
End Type

Sub Resize(ByRef values() As Long)
    ReDim values(0 To 1)
    values(0) = 9
    values(1) = 10
End Sub

Function Probe() As String
    Dim holder As New Holder, box As Box, errNo As Long
    ReDim holder.Values(0 To 0)
    ReDim box.Values(0 To 0)
    On Error Resume Next
    Resize holder.Values
    errNo = Err.Number
    Probe = CStr(errNo) & ":" & CStr(UBound(holder.Values)) & ":" & _
        CStr(holder.Values(0)) & ":" & CStr(holder.Values(1))
    Err.Clear
    Resize box.Values
    Probe = Probe & ";" & CStr(Err.Number) & ":" & CStr(UBound(box.Values)) & ":" & _
        CStr(box.Values(0)) & ":" & CStr(box.Values(1))
End Function
`;

const evaluator = evalVBASingle(code);
assert.strictEqual(
    evaluator.callProcedure('Probe', []),
    '0:1:9:10;0:1:9:10',
    'Class and UDT dynamic array members preserve ByRef ReDim writeback',
);
console.log('[PASS] ByRef ReDim writeback for Class/UDT array members');
