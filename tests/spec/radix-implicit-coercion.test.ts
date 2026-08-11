import { strict as assert } from 'node:assert';
import { evalVBASingle } from '../../test-libs/test-runner';

const evaluator = evalVBASingle(String.raw`
Function Probe() As String
    Dim value As Variant, number As Long, result As String
    On Error Resume Next
    value = "&H10"
    number = value
    result = CStr(number) & ":" & CStr(Err.Number)
    Err.Clear
    value = "&HFFFFFFFF"
    number = value
    result = result & ";" & CStr(number) & ":" & CStr(Err.Number)
    Err.Clear
    value = "&O37777777777"
    number = value
    result = result & ";" & CStr(number) & ":" & CStr(Err.Number)
    Err.Clear
    value = "&H10"
    number = value + 1
    result = result & ";" & CStr(number) & ":" & CStr(Err.Number)
    Probe = result
End Function
`);

assert.equal(evaluator.callProcedure('Probe', []), '16:0;-1:0;-1:0;-1:13',
    'implicit Variant radix coercion preserves assignment and arithmetic boundaries');
console.log('[PASS] implicit radix coercion matrix (4 boundaries)');
