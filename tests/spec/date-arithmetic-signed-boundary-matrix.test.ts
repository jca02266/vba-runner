import { strict as assert } from 'node:assert';
import { evalVBASingle } from '../../test-libs/test-runner';

const cases = [
    ['DateAdd("d", -1, #1/1/2024#)', '0|Date|31/12/2023'],
    ['DateAdd("d", 1, #12/31/2023#)', '0|Date|1/1/2024'],
    ['DateAdd("q", -1, #3/31/2024#)', '0|Date|31/12/2023'],
    ['DateAdd("yyyy", -1, #2/29/2024#)', '0|Date|28/2/2023'],
    ['DateAdd("m", -1, #3/31/2024#)', '0|Date|29/2/2024'],
    ['DateDiff("q", #3/31/2024#, #12/31/2023#)', '0|Long|-1'],
    ['DateDiff("yyyy", #2/29/2024#, #2/28/2023#)', '0|Long|-1'],
    ['DateDiff("d", #1/1/2024 00:00:00#, #12/31/2023 23:59:59#)', '0|Long|-1'],
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
        Probe = CStr(errNo) & "|"
    ElseIf TypeName(value) = "Date" Then
        Probe = CStr(errNo) & "|Date|" & CStr(Day(value)) & "/" & CStr(Month(value)) & "/" & CStr(Year(value))
    Else
        Probe = CStr(errNo) & "|" & TypeName(value) & "|" & CStr(value)
    End If
End Function`);
    const actual = ev.callProcedure('Probe', []);
    assert.equal(actual, expected, `Date arithmetic ${expression}`);
    console.log(`DATE-MATRIX ${expression} => ${actual}`);
}

assert.equal(cases.length, 8);
console.log(`[PASS] Date arithmetic signed boundary matrix (${cases.length} cases)`);
