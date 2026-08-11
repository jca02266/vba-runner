import { strict as assert } from 'node:assert';
import { evalVBASingle } from '../../test-libs/test-runner';

const cases = [
    ['DatePart("yyyy", CDate(-1))', '0|Integer|1899'],
    ['DatePart("m", CDate(-1))', '0|Integer|12'],
    ['DatePart("d", CDate(-1))', '0|Integer|29'],
    ['DateDiff("d", CDate(0), CDate(-1))', '0|Long|-1'],
    ['Hour(CDate(-0.5))', '0|Integer|12'],
    ['Minute(CDate(-0.5))', '0|Integer|0'],
    ['Second(CDate(-0.00001))', '0|Integer|0'],
    ['DateValue("1899/12/29")', '0|Date|1899/12/29'],
    ['DateValue("1899-12-29")', '0|Date|1899/12/29'],
    ['TimeValue("12:34:56.789")', '13|Empty|'],
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
    assert.equal(ev.callProcedure('Probe', []), expected, `Date text/serial ${expression}`);
}

assert.equal(cases.length, 10);
console.log(`[PASS] Date text/negative serial matrix (${cases.length} cases)`);
