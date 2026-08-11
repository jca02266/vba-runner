import { strict as assert } from 'node:assert';
import { evalVBASingle } from '../../test-libs/test-runner';

const evaluator = evalVBASingle(String.raw`
Function ProbeArray() As Long
    On Error Resume Next
    Array(Arglist:=1)
    ProbeArray = Err.Number
End Function

Function ProbeChoose() As Long
    On Error Resume Next
    Choose(Index:=1, Choice:=2)
    ProbeChoose = Err.Number
End Function

Function ProbeSwitch() As Long
    On Error Resume Next
    Switch(VarExpr:=True, Value:=2)
    ProbeSwitch = Err.Number
End Function
`);

assert.equal(evaluator.callProcedure('ProbeArray', []), 448,
    'Array ParamArray rejects named arguments');
assert.equal(evaluator.callProcedure('ProbeChoose', []), 448,
    'Choose ParamArray rejects named arguments');
assert.equal(evaluator.callProcedure('ProbeSwitch', []), 448,
    'Switch ParamArray rejects named arguments');
console.log('[PASS] Builtin ParamArray named-argument rejection');
