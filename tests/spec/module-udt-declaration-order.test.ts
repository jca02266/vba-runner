import { evalVBASingle, assert } from '../../test-libs/test-runner';

const source = String.raw`Option Explicit

Private state As ForwardState

Private Type ForwardState
    Value As Long
End Type

Public Function Probe() As String
    Dim initial As Long
    initial = state.Value
    state.Value = 42
    Probe = CStr(initial) & ":" & CStr(state.Value)
End Function
`;

const evaluator = evalVBASingle(source);
assert.strictEqual(evaluator.callProcedure('Probe', []), '0:42');

console.log('✅ module UDT variables are declaration-order invariant');
