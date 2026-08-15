import { strict as assert } from 'node:assert';
import { evalVBASingle } from '../../test-libs/test-runner';

const cases = [
    ['DateAdd("zz", 1, #1/1/2024#)', '5|Empty|'],
    ['DateDiff("zz", #1/1/2024#, #1/2/2024#)', '5|Empty|'],
    ['DatePart("zz", #1/1/2024#)', '5|Empty|'],
    ['DatePart("ww", #1/1/2024#, 0, 0)', '0|Integer|1'],
    ['DatePart("ww", #1/1/2024#, 8, 1)', '5|Empty|'],
    ['DatePart("ww", #1/1/2024#, 1, 4)', '5|Empty|'],
    ['Weekday(#1/1/2024#, 8)', '5|Empty|'],
    ['DateDiff("d", #1/1/2024#, #1/2/2024#, 8, 1)', '5|Empty|'],
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
    const actual = ev.callProcedure('Probe', []);
    assert.equal(actual, expected, `Date option ${expression}`);
    console.log(`DATE-OPTION ${expression} => ${actual}`);
}

assert.equal(cases.length, 8);
console.log(`[PASS] Date option validation matrix (${cases.length} cases)`);
