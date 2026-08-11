import { strict as assert } from 'node:assert';
import { evalVBASingle } from '../../test-libs/test-runner';

const cases = [
    ['Year(Empty)', '0|Integer|1899'],
    ['Month(Empty)', '0|Integer|12'],
    ['Day(Empty)', '0|Integer|30'],
    ['Hour(Empty)', '0|Integer|0'],
    ['Minute(Empty)', '0|Integer|0'],
    ['Second(Empty)', '0|Integer|0'],
    ['Year(Null)', '0|Null|'],
    ['Month(Null)', '0|Null|'],
    ['Day(Null)', '0|Null|'],
    ['Hour(Null)', '0|Null|'],
    ['Minute(Null)', '0|Null|'],
    ['Second(Null)', '0|Null|'],
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
    assert.equal(ev.callProcedure('Probe', []), expected, `Date scalar ${expression}`);
}

assert.equal(cases.length, 12);
console.log(`[PASS] Date scalar Empty/Null matrix (${cases.length} cases)`);
