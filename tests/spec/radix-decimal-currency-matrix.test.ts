import { strict as assert } from 'node:assert';
import { evalVBASingle } from '../../test-libs/test-runner';

const cases = [
    ['CDec("+1,234")', 'TYPE=Decimal:VALUE=1234'],
    ['CDec("-1,234")', 'TYPE=Decimal:VALUE=-1234'],
    ['CDec("1.2345")', 'TYPE=Decimal:VALUE=1.2345'],
    ['CDec("9.007199254740993125E+14")', 'TYPE=Decimal:VALUE=900719925474099.3125'],
    ['CDec("900,719,925,474,099.3125")', 'TYPE=Decimal:VALUE=900719925474099.3125'],
    ['CDec("-&H10")', 'ERR=13'],
    ['CCur("+1,234")', 'TYPE=Currency:VALUE=1234'],
    ['CCur("-1,234")', 'TYPE=Currency:VALUE=-1234'],
    ['CCur("1.23456")', 'TYPE=Currency:VALUE=1.2346'],
    ['CCur("9.007199254740993125E+14")', 'TYPE=Currency:VALUE=900719925474099.3125'],
    ['CCur("900,719,925,474,099.3125")', 'TYPE=Currency:VALUE=900719925474099.3125'],
    ['CCur("-&H10")', 'TYPE=Currency:VALUE=-16'],
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
    assert.equal(ev.callProcedure('Probe', []), expected, `Decimal/Currency ${expression}`);
}

assert.equal(cases.length, 12);
console.log(`[PASS] radix Decimal/Currency matrix (${cases.length} cases)`);
