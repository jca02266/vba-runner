import { strict as assert } from 'node:assert';
import { evalVBASingle } from '../../test-libs/test-runner';

const cases = [
    ['CInt("&H10000")', 'ERR=6'],
    ['CInt("&O200000")', 'ERR=6'],
    ['CLng("&H100000000")', 'ERR=6'],
    ['CLng("&O40000000000")', 'ERR=6'],
    ['CLngLng("&H10000000000000000")', 'ERR=6'],
    ['CLngLng("&O2000000000000000000000")', 'ERR=6'],
    ['CDec("&H10000000000000000")', 'TYPE=Decimal:VALUE=18446744073709551616'],
    ['CDec("&O2000000000000000000000")', 'TYPE=Decimal:VALUE=18446744073709551616'],
    ['CCur("&H10000000000000000")', 'ERR=6'],
    ['CCur("&O2000000000000000000000")', 'ERR=6'],
    ['Val("&H10000000000000000")', 'TYPE=Double:VALUE=1.84467440737096E+19'],
    ['Val("&O2000000000000000000000")', 'TYPE=Double:VALUE=1.84467440737096E+19'],
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
        Probe = "TYPE=" & TypeName(value) & ":VALUE=" & CStr(value)
    End If
End Function`);
    assert.equal(ev.callProcedure('Probe', []), expected, `Radix width overflow ${expression}`);
}

assert.equal(cases.length, 12);
console.log(`[PASS] radix width overflow matrix (${cases.length} cases)`);
