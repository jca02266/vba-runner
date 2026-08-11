import { strict as assert } from 'node:assert';
import { evalVBASingle } from '../../test-libs/test-runner';

const ev = evalVBASingle(String.raw`
Class Payload
    Public Id As Long
End Class

Class Box
    Private saved As Long
    Private refValue As Object
    Public Property Let Value(ByRef value As Long)
        saved = value
        value = value + 10
    End Property
    Public Property Get Value() As Long
        Value = saved
    End Property
    Public Property Set Reference(ByRef value As Object)
        Set refValue = value
    End Property
    Public Property Get Reference() As Object
        Set Reference = refValue
    End Property
End Class

Type Holder
    Items(0 To 1) As Object
End Type

Function Probe() As String
    Dim box As New Box, payload As New Payload, holder As Holder
    Dim value As Long, values() As Long, fixed(0 To 0) As Long, errNo As Long
    payload.Id = 7
    value = 3
    box.Value = value
    Probe = CStr(box.Value) & ":" & CStr(value)
    Set box.Reference = payload
    Probe = Probe & ":" & CStr(box.Reference.Id)
    Set holder.Items(0) = payload
    Set holder.Items(1) = Nothing
    Probe = Probe & ":" & TypeName(holder.Items(0)) & ":" & TypeName(holder.Items(1))
    ReDim values(0 To 0)
    values(0) = 1
    On Error Resume Next
    Err.Clear
    ReDim values(0 To 1)
    errNo = Err.Number
    Probe = Probe & ":" & CStr(errNo) & ":" & CStr(UBound(values))
    Err.Clear
    fixed(0) = 4
    ReDim fixed(0 To 1)
    Probe = Probe & ":" & CStr(Err.Number) & ":" & CStr(fixed(0))
End Function`);

assert.equal(ev.callProcedure('Probe', []), '3:3:7:Payload:Nothing:0:1:10:4');
console.log('[PASS] Property/ByRef object-array matrix (8 boundaries)');
