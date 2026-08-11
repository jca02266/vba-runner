import { strict as assert } from 'node:assert';
import { evalVBASingle } from '../../test-libs/test-runner';

const evaluator = evalVBASingle(String.raw`
Class Bag
    Public Property Let Item(ParamArray indexes(), value As Variant)
    End Property
    Public Seen As Long
    Public Property Let Keyed(ByVal key As Long, ParamArray indexes(), value As Variant)
        Seen = key
    End Property
    Public Property Get Read(ParamArray indexes()) As Long
        Read = 1
    End Property
End Class

Sub Plain(ParamArray values())
End Sub

Function Probe() As Long
    Dim bag As New Bag
    On Error Resume Next
    bag.Item(1, value:=9)
    Probe = Err.Number
End Function

Function ProbeGet() As Long
    Dim bag As New Bag
    On Error Resume Next
    ProbeGet = bag.Read(indexes:=1)
    ProbeGet = Err.Number
End Function

Function ProbePlain() As Long
    On Error Resume Next
    Plain(value:=1)
    ProbePlain = Err.Number
End Function

Function ProbePositional() As Long
    Dim bag As New Bag
    On Error Resume Next
    Err.Clear
    bag.Item(1) = 9
    ProbePositional = Err.Number
End Function

Function ProbeLeadingNamed() As Long
    Dim bag As New Bag
    On Error Resume Next
    bag.Keyed(key:=1) = 9
    ProbeLeadingNamed = Err.Number
End Function
`);

assert.equal(
    evaluator.callProcedure('Probe', []),
    448,
    'Property Let with ParamArray must reject named arguments',
);
assert.equal(evaluator.callProcedure('ProbeGet', []), 448,
    'Property Get with ParamArray must reject named arguments');
assert.equal(evaluator.callProcedure('ProbePlain', []), 448,
    'Module ParamArray must reject named arguments');
assert.equal(evaluator.callProcedure('ProbePositional', []), 0,
    'Positional Property ParamArray remains valid');
assert.equal(evaluator.callProcedure('ProbeLeadingNamed', []), 448,
    'Leading named Property ParamArray arguments are rejected');
console.log('[PASS] Property ParamArray named-argument rejection');
