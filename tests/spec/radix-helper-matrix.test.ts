import { strict as assert } from 'node:assert';
import { evalVBASingle } from '../../test-libs/test-runner';

const cases = [
    ['CByte("&HFF")', 'TYPE=Byte:VALUE=255'],
    ['CByte("&H100")', 'ERR=6'],
    ['CSng("&H80000000")', 'TYPE=Single:VALUE=2147483648'],
    ['CSng("&H100000000")', 'TYPE=Single:VALUE=4294967296'],
    ['CDbl("&HFFFFFFFF")', 'TYPE=Double:VALUE=4294967295'],
    ['CDate("&H10")', 'TYPE=Date:VALUE=1900/01/15'],
    ['Hex("&HFFFF")', 'TYPE=String:VALUE=FFFF'],
    ['Hex("&H100000000")', 'TYPE=String:VALUE=100000000'],
    ['Oct("&O77777777777")', 'TYPE=String:VALUE=77777777777'],
    ['Oct("&O100000000000")', 'TYPE=String:VALUE=100000000000'],
    ['Val("&H100000000")', 'TYPE=Double:VALUE=4294967296'],
    ['Val("&O40000000000")', 'TYPE=Double:VALUE=4294967296'],
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
    assert.equal(ev.callProcedure('Probe', []), expected, `Radix helper ${expression}`);
}

assert.equal(cases.length, 12);
console.log(`[PASS] radix helper matrix (${cases.length} cases)`);
