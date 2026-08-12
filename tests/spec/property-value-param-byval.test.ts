import { evalVBAModules, evalVBASingle, assert } from '../../test-libs/test-runner';

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

const propertyBoundaryModules = [
  {
    name: 'PropertyBoundaryBox',
    code: String.raw`Class PropertyBoundaryBox
Private seen As Long
Public Property Let Item(ByRef value As Long)
  seen = value
  value = value + 1
End Property
Public Property Let Converted(ByRef value As Long)
  seen = value
End Property
Public Function Read() As Long
  Read = seen
End Function
End Class`,
  },
  {
    name: 'PropertyBoundaryModule',
    code: String.raw`Function VerifyPropertyBoundaries() As String
Dim box As New PropertyBoundaryBox, values(0 To 0) As Long, source As Integer
values(0) = 7
box.Item = values(0)
source = 12
box.Converted = source
VerifyPropertyBoundaries = CStr(box.Read()) & ":" & CStr(values(0)) & ":" & CStr(source)
End Function`,
  },
];
const propertyBoundaryRunner = evalVBAModules(propertyBoundaryModules);
assert.strictEqual(
  propertyBoundaryRunner.callProcedure('VerifyPropertyBoundaries', []),
  '12:7:12',
  'Property Let value arguments must not write back to array elements and must accept numeric conversion',
);
console.log('[PASS] Property Let array-element and conversion boundaries');
