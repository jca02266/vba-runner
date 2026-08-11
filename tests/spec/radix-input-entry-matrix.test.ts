import { strict as assert } from 'node:assert';
import { evalVBASingle } from '../../test-libs/test-runner';

const evaluator = evalVBASingle(String.raw`
Function Probe() As String
    Dim value As Variant, result As String
    On Error Resume Next
    value = CByte("&H100")
    result = CStr(Err.Number)
    Err.Clear
    value = CSng("&HFFFFFFFF")
    result = result & ":" & CStr(Err.Number)
    Err.Clear
    value = CDbl("&HFFFFFFFFFFFFFFFF")
    result = result & ":" & CStr(Err.Number)
    Err.Clear
    value = CDate("&H10")
    result = result & ":" & CStr(Err.Number)
    Err.Clear
    result = result & ":" & Hex("&H100000000")
    Err.Clear
    result = result & ":" & Oct("&H100000000")
    Probe = result
End Function
`);

assert.equal(evaluator.callProcedure('Probe', []), '6:0:0:0:100000000:40000000000',
    'radix conversion entry points preserve width, coercion, and formatting contracts');
console.log('[PASS] radix conversion input-entry matrix (6 boundaries)');
