import { evalVBASingle, assert } from '../../test-libs/test-runner';

const code = String.raw`
Class ValueObject
    Public Property Get Value() As Long
        Value = 4
    End Property
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
        seen = CStr(UBound(indexes) - LBound(indexes) + 1) & ":" & CStr(indexes(0)) & ":" & CStr(value.Value)
    End Property
End Class

Function ProbeLet() As String
    Dim bag As New Bag
    On Error Resume Next
    bag.Item(1, 2) = 7
    ProbeLet = CStr(Err.Number) & "|" & bag.Item
End Function

Function ProbeSet() As String
    Dim bag As New Bag, value As New ValueObject
    On Error Resume Next
    Set bag.ObjectItem(1, 2) = value
    ProbeSet = CStr(Err.Number) & "|" & bag.Item
End Function

Function ProbeCallByName() As String
    Dim bag As New Bag, value As New ValueObject
    On Error Resume Next
    CallByName bag, "Item", VbLet, 1, 2, 8
    ProbeCallByName = CStr(Err.Number) & "|" & CStr(CallByName(bag, "Item", VbGet)) & "|"
    Err.Clear
    CallByName bag, "ObjectItem", VbSet, 1, 2, value
    ProbeCallByName = ProbeCallByName & CStr(Err.Number) & "|" & CStr(CallByName(bag, "Item", VbGet))
End Function
`;

const evaluator = evalVBASingle(code);
assert.strictEqual(evaluator.callProcedure('ProbeLet', []), '0|2:1:7',
    'Property Let excludes its RHS from ParamArray');
assert.strictEqual(evaluator.callProcedure('ProbeSet', []), '0|2:1:4',
    'Property Set excludes its Object RHS from ParamArray');
assert.strictEqual(evaluator.callProcedure('ProbeCallByName', []), '0|2:1:8|0|2:1:4',
    'CallByName preserves Property Let/Set ParamArray and value-tail binding');
console.log('[PASS] Property Let/Set ParamArray value-tail binding');
