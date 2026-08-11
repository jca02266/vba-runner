import { strict as assert } from 'node:assert';
import { evalVBASingle } from '../../test-libs/test-runner';

const cases = [
    ['DateValue("2024/02/30")', '13|Empty|'],
    ['DateValue("not-a-date")', '13|Empty|'],
    ['TimeValue("25:00")', '13|Empty|'],
    ['TimeValue("not-a-time")', '13|Empty|'],
    ['DatePart("yyyy", True)', '13|Empty|'],
    ['DatePart("yyyy", False)', '13|Empty|'],
    ['DatePart("yyyy", CVErr(2042))', '13|Empty|'],
    ['DatePart("yyyy", Array(1, 2))', '13|Empty|'],
    ['Year(Nothing)', '91|Empty|'],
    ['DateDiff("d", Nothing, CDate("2024-01-01"))', '91|Empty|'],
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
    assert.equal(ev.callProcedure('Probe', []), expected, `Date invalid input ${expression}`);
}

assert.equal(cases.length, 10);
console.log(`[PASS] Date invalid input matrix (${cases.length} cases)`);
