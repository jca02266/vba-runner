import { strict as assert } from 'node:assert';
import { evalVBASingle } from '../../test-libs/test-runner';

const source = String.raw`
Public Function Probe(ByVal text As String) As String
    Dim value As Variant, errNo As Long
    On Error Resume Next
    Err.Clear
    value = Val(text)
    errNo = Err.Number
    If errNo <> 0 Then
        Probe = "ERR=" & CStr(errNo)
    Else
        Probe = CStr(value)
    End If
    On Error GoTo 0
End Function
`;
const ev = evalVBASingle(source);
const cases: Array<[string, string]> = [
    ['1.23456@', '1.2346'],
    ['1.23456789!', '1.2345678806304932'],
    ['32767%', '32767'],
    ['32768%', 'ERR=6'],
    ['2147483647&', '2147483647'],
    ['2147483648&', 'ERR=6'],
    ['1E40!', 'ERR=6'],
    ['1E309#', 'ERR=6'],
    ['922337203685478@', 'ERR=6'],
    [' +32767%', '32767'],
    [' -32768%', '-32768'],
    [' +2147483647&', '2147483647'],
    [' -2147483648&', '-2147483648'],
    [' 1.23456@', '1.2346'],
    [' +1.5^', 'ERR=13'],
];
for (const [input, expected] of cases) assert.equal(ev.callProcedure('Probe', [input]), expected, input);
console.log(`[PASS] Val type suffix boundary (${cases.length} cases)`);
