import { strict as assert } from 'node:assert';
import { evalVBASingle } from '../../test-libs/test-runner';

const ev = evalVBASingle(String.raw`
Sub TakeLong(ByRef value As Long)
    value = value + 1
End Sub
Function ProbeFunction() As Long
    On Error Resume Next
    Err.Clear: TakeLong CLng(10)
    ProbeFunction = Err.Number
End Function
Function ProbeLiteral() As Long
    On Error Resume Next
    Err.Clear: TakeLong 10
    ProbeLiteral = Err.Number
End Function
Function ProbeParenthesized() As Long
    On Error Resume Next
    Err.Clear: TakeLong (10)
    ProbeParenthesized = Err.Number
End Function`);

assert.equal(ev.callProcedure('ProbeFunction', []), 0);
assert.equal(ev.callProcedure('ProbeLiteral', []), 0);
assert.equal(ev.callProcedure('ProbeParenthesized', []), 0);
console.log('[PASS] ByRef expression boundary runner behavior');
