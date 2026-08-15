import { strict as assert } from 'node:assert';
import { evalVBASingle } from '../../test-libs/test-runner';

const source = String.raw`Function Probe(ByVal kind As String, ByVal text As String) As String
Dim value As Variant, errNo As Long
On Error Resume Next
Err.Clear
If kind = "date" Then
    value = DateValue(text)
Else
    value = TimeValue(text)
End If
errNo = Err.Number
If errNo <> 0 Then
    Probe = "ERR=" & CStr(errNo)
Else
    Probe = CStr(Year(value)) & "/" & CStr(Month(value)) & "/" & CStr(Day(value)) & " " & CStr(Hour(value)) & ":" & CStr(Minute(value)) & ":" & CStr(Second(value))
End If
On Error GoTo 0
End Function`;

const cases: Array<[string, string, string]> = [
    ['date', '2024-02-29', '2024/2/29 0:0:0'],
    ['date', '2024/02/29 12:34:56', '2024/2/29 0:0:0'],
    ['date', '2024/02/30', 'ERR=13'],
    ['date', '02/29/2023', 'ERR=13'],
    ['time', '00:00:00', '1899/12/30 0:0:0'],
    ['time', '14:30:45', '1899/12/30 14:30:45'],
    ['time', '23:59:59', '1899/12/30 23:59:59'],
];
const ev = evalVBASingle(source);
for (const [kind, text, expected] of cases) {
    const actual = ev.callProcedure('Probe', [kind, text]);
    assert.equal(actual, expected, `${kind} ${text}`);
    console.log(`DATE-VALUE ${kind}|${text} => ${actual}`);
}
