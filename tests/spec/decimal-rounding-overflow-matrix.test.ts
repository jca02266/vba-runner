import { strict as assert } from 'node:assert';
import { evalVBASingle } from '../../test-libs/test-runner';

const cases = [
    ['CInt(2.5)', 'TYPE=Integer:VALUE=2'],
    ['CInt(3.5)', 'TYPE=Integer:VALUE=4'],
    ['CLng(-2.5)', 'TYPE=Long:VALUE=-2'],
    ['CLng(-3.5)', 'TYPE=Long:VALUE=-4'],
    ['CCur("0.00005")', 'TYPE=Currency:VALUE=0'],
    ['CCur("0.00015")', 'TYPE=Currency:VALUE=0.0002'],
    ['CCur("922337203685477.5807")', 'TYPE=Currency:VALUE=922337203685477.5807'],
    ['CCur("922337203685477.5808")', 'ERR=6'],
    ['CDec("79228162514264337593543950335")', 'TYPE=Decimal:VALUE=79228162514264337593543950335'],
    ['CDec("79228162514264337593543950336")', 'ERR=6'],
    ['CInt("32767.5")', 'ERR=6'],
    ['CLng("2147483647.5")', 'ERR=6'],
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
    assert.equal(ev.callProcedure('Probe', []), expected, `Decimal rounding/overflow ${expression}`);
}

assert.equal(cases.length, 12);
console.log(`[PASS] Decimal/Currency rounding overflow matrix (${cases.length} cases)`);
