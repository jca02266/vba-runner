import { strict as assert } from 'node:assert';
import { evalVBASingle } from '../../test-libs/test-runner';

const evaluator = evalVBASingle(String.raw`
Sub Plain(ParamArray values())
End Sub

Function WithResult(ParamArray values()) As Long
    WithResult = 1
End Function

Function ProbeSub() As Long
    On Error Resume Next
    Plain(values:=1)
    ProbeSub = Err.Number
End Function

Function ProbeFunction() As Long
    On Error Resume Next
    WithResult(values:=1)
    ProbeFunction = Err.Number
End Function
`);

assert.equal(evaluator.callProcedure('ProbeSub', []), 448,
    'Module Sub ParamArray rejects a named ParamArray argument');
assert.equal(evaluator.callProcedure('ProbeFunction', []), 448,
    'Module Function ParamArray rejects a named ParamArray argument');
console.log('[PASS] Module ParamArray named-argument rejection');
