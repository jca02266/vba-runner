import { strict as assert } from 'node:assert';
import { evalVBASingle } from '../../test-libs/test-runner';

const cases = [
    ['Year(DateValue(Empty))', '0|Integer|1899'],
    ['Hour(TimeValue(Empty))', '0|Integer|0'],
    ['DatePart("yyyy", CDec(Empty))', '0|Integer|1899'],
    ['DatePart("yyyy", CCur(Empty))', '0|Integer|1899'],
    ['DateDiff("d", CDec(Empty), CDate("2024-01-01"))', '0|Long|45292'],
    ['DateDiff("d", CCur(Empty), CDate("2024-01-01"))', '0|Long|45292'],
    ['DatePart("yyyy", Null)', '0|Null|'],
    ['DateDiff("d", Null, Null)', '0|Null|'],
    ['Year(Nothing)', '91|Empty|'],
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
    ElseIf errNo <> 0 Then
        Probe = CStr(errNo) & "|" & TypeName(value) & "|"
    Else
        Probe = CStr(errNo) & "|" & TypeName(value) & "|" & CStr(value)
    End If
End Function`);
    assert.equal(ev.callProcedure('Probe', []), expected, `Date cross-type ${expression}`);
}

assert.equal(cases.length, 9);
console.log(`[PASS] Date conversion cross-type matrix (${cases.length} cases)`);
