import { evalVBASingle, assert } from '../../test-libs/test-runner';

const code = String.raw`
Class Payload
    Public Name As String
End Class

Type Holder
    Values(0 To 1) As Object
End Type

Function Probe() As String
    Dim holder As Holder, payload As New Payload, errNo As Long
    payload.Name = "payload"
    On Error Resume Next
    Err.Clear
    Set holder.Values(0) = payload
    errNo = Err.Number
    Probe = CStr(errNo) & ":" & TypeName(holder.Values(0))
    Err.Clear
    Set holder.Values(1) = Nothing
    Probe = Probe & ";" & CStr(Err.Number) & ":" & TypeName(holder.Values(1))
End Function
`;

const evaluator = evalVBASingle(code);
assert.strictEqual(evaluator.callProcedure('Probe', []), '0:Payload;0:Nothing',
    'UDT Object-array Set assignments retain objects and Nothing');
console.log('[PASS] UDT Object-array Set assignment');
