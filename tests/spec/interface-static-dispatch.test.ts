import { evalVBASingle, assert } from '../../test-libs/test-runner';

const code = String.raw`
Class IThing
    Public Sub Mutate(ByRef value As Long)
    End Sub
End Class

Class Thing
    Implements IThing
    Public Sub IThing_Mutate(ByRef value As Long)
        value = value + 1
    End Sub
End Class

Function ConcreteCall() As String
    Dim concrete As New Thing, value As Long
    value = 1
    On Error Resume Next
    concrete.Mutate value
    ConcreteCall = CStr(Err.Number) & ":" & CStr(value)
End Function

Function InterfaceCall() As String
    Dim concrete As New Thing, target As IThing, value As Long
    Set target = concrete
    value = 1
    On Error Resume Next
    Err.Clear
    target.Mutate value
    InterfaceCall = CStr(Err.Number) & ":" & CStr(value)
End Function
`;

const evaluator = evalVBASingle(code);
assert.strictEqual(evaluator.callProcedure('ConcreteCall', []), '438:1',
    'concrete class references must not expose Interface-only members');
assert.strictEqual(evaluator.callProcedure('InterfaceCall', []), '0:2',
    'Interface-typed references dispatch to Implements members');
console.log('[PASS] static class/interface dispatch boundary');
