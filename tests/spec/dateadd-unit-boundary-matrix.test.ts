import { strict as assert } from 'node:assert';
import { evalVBASingle } from '../../test-libs/test-runner';

const cases = [
    ['DateAdd("ww", -1, #1/7/2024#)', '0|31/12/2023 0:0:0'],
    ['DateAdd("ww", 1, #12/31/2023#)', '0|7/1/2024 0:0:0'],
    ['DateAdd("h", -1, #1/1/2024 00:30:00#)', '0|31/12/2023 23:30:0'],
    ['DateAdd("h", 1, #12/31/2023 23:30:00#)', '0|1/1/2024 0:30:0'],
    ['DateAdd("n", -31, #1/1/2024 00:30:00#)', '0|31/12/2023 23:59:0'],
    ['DateAdd("s", -1, #1/1/2024 00:00:00#)', '0|31/12/2023 23:59:59'],
    ['DateAdd("s", 1, #12/31/2023 23:59:59#)', '0|1/1/2024 0:0:0'],
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
    Else
        Probe = CStr(errNo) & "|" & CStr(Day(value)) & "/" & CStr(Month(value)) & "/" & CStr(Year(value)) & " " & CStr(Hour(value)) & ":" & CStr(Minute(value)) & ":" & CStr(Second(value))
    End If
End Function`);
    assert.equal(ev.callProcedure('Probe', []), expected, `DateAdd unit ${expression}`);
}

assert.equal(cases.length, 7);
console.log(`[PASS] DateAdd week/time unit matrix (${cases.length} cases)`);
