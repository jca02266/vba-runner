import { strict as assert } from 'node:assert';
import { evalVBASingle } from '../../test-libs/test-runner';

const cases = [
    ['TypeName(&HFFFFFFFF)', 'Long'],
    ['TypeName(&HFFFFFFFFFFFFFFFF)', 'Double'],
    ['TypeName(&HFFFFFFFFFFFFFFFF^)', 'LongLong'],
    ['CLngLng("&HFFFFFFFFFFFFFFFF")', '-1'],
    ['CLngLng("&H8000000000000000")', '-9223372036854775808'],
    ['CLngPtr("&HFFFFFFFF")', '4294967295'],
    ['TypeName(CLngLng("&HFFFFFFFFFFFFFFFF"))', 'LongLong'],
    ['TypeName(CLngPtr("&HFFFFFFFF"))', 'LongPtr'],
] as const;

for (const [expression, expected] of cases) {
    const ev = evalVBASingle(String.raw`Option Explicit
Public Function Probe() As String
    Dim value As Variant, errNo As Long
    On Error Resume Next
    Err.Clear
    value = ${expression}
    errNo = Err.Number
    If errNo <> 0 Then
        Probe = "ERR=" & CStr(errNo)
    Else
        Probe = CStr(value)
    End If
End Function`);
    assert.equal(ev.callProcedure('Probe', []), expected, `LongLong subtype ${expression}`);
}

assert.equal(cases.length, 8);
console.log(`[PASS] radix LongLong/LongPtr subtype matrix (${cases.length} cases)`);
