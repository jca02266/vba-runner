import { strict as assert } from 'node:assert';
import { evalVBASingle } from '../../test-libs/test-runner';

const source = String.raw`
Public Function Probe(ByVal y As Long, ByVal m As Long, ByVal d As Long) As String
    On Error Resume Next
    Err.Clear
    Probe = CStr(DateSerial(y, m, d)) & ":ERR=" & CStr(Err.Number)
    On Error GoTo 0
End Function

Public Function ProbeYear(ByVal y As Long) As String
    Dim value As Variant, errNo As Long
    On Error Resume Next
    Err.Clear
    value = DateSerial(y, 1, 1)
    errNo = Err.Number
    If errNo <> 0 Then
        ProbeYear = "ERR=" & CStr(errNo)
    Else
        ProbeYear = CStr(Year(value))
    End If
    On Error GoTo 0
End Function
`;
const ev = evalVBASingle(source);
const cases: Array<[number, number, number, string]> = [
    [2024, 13, 1, '2025/01/01:ERR=0'],
    [2024, 0, 15, '2023/12/15:ERR=0'],
    [2024, 1, 32, '2024/02/01:ERR=0'],
    [2024, 1, 0, '2023/12/31:ERR=0'],
    [2024, 2, 29, '2024/02/29:ERR=0'],
    [2023, 2, 29, '2023/03/01:ERR=0'],
];
for (const [y, m, d, expected] of cases) assert.equal(ev.callProcedure('Probe', [y, m, d]), expected);
const yearCases: Array<[number, string]> = [[0, '2000'], [29, '2029'], [30, '1930'], [99, '1999'], [100, '100'], [-1, 'ERR=6']];
for (const [year, expected] of yearCases) assert.equal(ev.callProcedure('ProbeYear', [year]), expected);
console.log(`[PASS] DateSerial rollover boundary (${cases.length + yearCases.length} cases)`);
