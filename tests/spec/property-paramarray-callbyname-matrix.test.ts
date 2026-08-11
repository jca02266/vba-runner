import { strict as assert } from 'node:assert';
import { evalVBASingle } from '../../test-libs/test-runner';

const ev = evalVBASingle(String.raw`
Class ValueObject
    Public Id As Long
End Class
Class Bag
    Private seen As String
    Public Property Get Item(ParamArray indexes()) As String
        Item = seen
    End Property
    Public Property Let Item(ParamArray indexes(), value As Variant)
        seen = CStr(UBound(indexes) - LBound(indexes) + 1) & ":" & CStr(indexes(0)) & ":" & CStr(value)
    End Property
    Public Property Set ObjectItem(ParamArray indexes(), value As Object)
        seen = CStr(UBound(indexes) - LBound(indexes) + 1) & ":" & CStr(indexes(0)) & ":" & CStr(value.Id)
    End Property
End Class
Function Probe() As String
    Dim bag As New Bag, value As New ValueObject
    value.Id = 4
    On Error Resume Next
    bag.Item(1, 2) = 8
    Probe = CStr(Err.Number) & "|" & bag.Item
    Err.Clear
    Set bag.ObjectItem(1, 2) = value
    Probe = Probe & ";" & CStr(Err.Number) & "|" & bag.Item
    Err.Clear
    CallByName bag, "Item", VbLet, 1, 2, 9
    Probe = Probe & ";" & CStr(Err.Number) & "|" & bag.Item
    Err.Clear
    CallByName bag, "ObjectItem", VbSet, 1, 2, value
    Probe = Probe & ";" & CStr(Err.Number) & "|" & bag.Item
End Function`);

assert.equal(ev.callProcedure('Probe', []), '0|2:1:8;0|2:1:4;0|2:1:9;0|2:1:4');
console.log('[PASS] Property ParamArray/CallByName matrix (4 boundaries)');
