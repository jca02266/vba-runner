import { evalVBASingle, assert } from '../../test-libs/test-runner';

const code = String.raw`
Class Payload
    Public Name As String
End Class

Class Probe
    Private stored As Long
    Public Property Let ExplicitValue(ByRef value As Long)
        stored = value
        value = value + 9
    End Property
    Public Property Get ExplicitValue() As Long
        ExplicitValue = stored
    End Property
    Public Property Let ImplicitValue(value As Long)
        stored = value
        value = value + 7
    End Property
    Public Property Get ImplicitValue() As Long
        ImplicitValue = stored
    End Property
    Public Property Set ObjectValue(ByRef value As Object)
        Set value = Nothing
    End Property
End Class

Function VerifyPropertyValueByVal() As String
    Dim probe As New Probe, payload As New Payload, value As Long
    value = 5
    probe.ExplicitValue = value
    VerifyPropertyValueByVal = CStr(probe.ExplicitValue) & ":" & CStr(value)
    value = 6
    probe.ImplicitValue = value
    VerifyPropertyValueByVal = VerifyPropertyValueByVal & ":" & CStr(probe.ImplicitValue) & ":" & CStr(value)
    Set probe.ObjectValue = payload
    VerifyPropertyValueByVal = VerifyPropertyValueByVal & ":" & CStr(payload Is Nothing)
End Function
`;

const evaluator = evalVBASingle(code);
assert.strictEqual(
    evaluator.callProcedure('VerifyPropertyValueByVal', []),
    '5:5:6:6:False',
    'Property Let/Set value-param is always ByVal, including explicit and implicit ByRef syntax',
);
console.log('[PASS] Property Let/Set value-parameter ByVal boundary');
