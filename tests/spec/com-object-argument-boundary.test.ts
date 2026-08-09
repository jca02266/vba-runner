import { evalVBASingle, MockApplication, assert } from '../../test-libs/test-runner';

const evaluator = evalVBASingle(String.raw`
Function VariantName(ByVal value As Variant) As String
    VariantName = TypeName(value)
End Function

Function ObjectName(ByVal value As Object) As String
    ObjectName = TypeName(value)
End Function

Function OptionalObjectName(Optional ByVal value As Object) As String
    If value Is Nothing Then
        OptionalObjectName = "Nothing"
    Else
        OptionalObjectName = TypeName(value)
    End If
End Function

Function ProbeObjectArgumentBoundaries() As String
    Dim sheet As Object, errNull As Long, errMissing As Long
    Set sheet = ActiveSheet
    ProbeObjectArgumentBoundaries = VariantName(sheet) & "|" & _
        VariantName(value:=sheet) & "|" & ObjectName(sheet) & "|" & _
        ObjectName(value:=sheet)

    On Error Resume Next
    ObjectName Null
    errNull = Err.Number
    Err.Clear
    ProbeObjectArgumentBoundaries = ProbeObjectArgumentBoundaries & "|N=" & CStr(errNull)
    ProbeObjectArgumentBoundaries = ProbeObjectArgumentBoundaries & "|M=" & OptionalObjectName()
    errMissing = Err.Number
    ProbeObjectArgumentBoundaries = ProbeObjectArgumentBoundaries & "|ME=" & CStr(errMissing)
End Function
`, {
    setup: (ev) => ev.setDefaultBindingObject(new MockApplication()),
});

assert.strictEqual(
    evaluator.callProcedure('ProbeObjectArgumentBoundaries', []),
    'Worksheet|Worksheet|Worksheet|Worksheet|N=424|M=Nothing|ME=0',
    'Module Object/Variant binders preserve host objects, reject Null, and default Optional Object to Nothing',
);
console.log('[PASS] Module Object/Variant/Optional Object argument boundaries');
