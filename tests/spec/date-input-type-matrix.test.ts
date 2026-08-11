import { strict as assert } from 'node:assert';
import { evalVBASingle } from '../../test-libs/test-runner';

const cases = [
    ['DatePart("yyyy", CDate("2024-01-01"))', '0|Integer|2024'],
    ['DatePart("yyyy", "2024-01-01")', '0|Integer|2024'],
    ['DatePart("yyyy", 45292)', '0|Integer|2024'],
    ['DatePart("yyyy", Empty)', '0|Integer|1899'],
    ['DatePart("yyyy", Null)', '0|Null|'],
    ['Weekday(CDate("2024-01-01"))', '0|Integer|2'],
    ['Weekday(Null)', '0|Null|'],
    ['DateDiff("d", CDate("2024-01-01"), "2024-01-02")', '0|Long|1'],
    ['DateDiff("d", Empty, CDate("2024-01-01"))', '0|Long|45292'],
] as const;

for (const [expression, expected] of cases) {
    const ev = evalVBASingle(String.raw`Option Explicit
Public Function Probe() As String
    Dim value As Variant, errNo As Long
    On Error Resume Next
    Err.Clear
    value = ${expression}
    errNo = Err.Number
    If IsNull(value) Then
        Probe = CStr(errNo) & "|Null|"
    Else
        Probe = CStr(errNo) & "|" & TypeName(value) & "|" & CStr(value)
    End If
End Function`);
    assert.equal(ev.callProcedure('Probe', []), expected, `Date input type ${expression}`);
}

assert.equal(cases.length, 9);
console.log(`[PASS] Date input type matrix (${cases.length} cases)`);
