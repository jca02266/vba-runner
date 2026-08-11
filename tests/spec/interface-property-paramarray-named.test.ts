import { strict as assert } from 'node:assert';
import { evalVBASingle } from '../../test-libs/test-runner';

const evaluator = evalVBASingle(String.raw`
Class IStore
    Public Property Let Item(ParamArray indexes(), value As Variant)
    End Property
End Class

Class Store
    Implements IStore
    Public Seen As Long
    Public Property Let IStore_Item(ParamArray indexes(), value As Variant)
        Seen = CInt(value)
    End Property
End Class

Function ProbeNamed() As Long
    Dim concrete As New Store, target As IStore
    Set target = concrete
    On Error Resume Next
    target.Item(1, value:=9)
    ProbeNamed = Err.Number
End Function

Function ProbePositional() As Long
    Dim concrete As New Store, target As IStore
    Set target = concrete
    On Error Resume Next
    Err.Clear
    target.Item(1) = 9
    ProbePositional = Err.Number
End Function
`);

assert.equal(evaluator.callProcedure('ProbeNamed', []), 448,
    'Interface Property ParamArray rejects named value arguments');
assert.equal(evaluator.callProcedure('ProbePositional', []), 0,
    'Interface Property ParamArray preserves positional value-tail assignment');
console.log('[PASS] Interface Property ParamArray named matrix');
