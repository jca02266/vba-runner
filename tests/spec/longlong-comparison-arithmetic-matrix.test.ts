import { strict as assert } from 'node:assert';
import { evalVBASingle } from '../../test-libs/test-runner';

const cases = [
    ['1^', 'E1G1C1|0|2:0'],
    ['&H20000000000001^', 'E1G1C1|9007199254740992|9007199254740994:0'],
    ['9223372036854775807^', 'E1G1C1:6'],
] as const;

for (const [literal, expected] of cases) {
    const ev = evalVBASingle(String.raw`Option Explicit
Public Function Probe() As String
    Dim value As LongLong, output As String, errNo As Long
    value = ${literal}
    If value = CLngLng(CStr(value)) Then output = output & "E1" Else output = output & "E0"
    If value >= CLngLng(CStr(value)) Then output = output & "G1" Else output = output & "G0"
    Select Case value
    Case CLngLng(CStr(value))
        output = output & "C1"
    Case Else
        output = output & "C0"
    End Select
    On Error Resume Next
    Err.Clear
    output = output & "|" & CStr(value - 1) & "|" & CStr(value + 1)
    errNo = Err.Number
    Probe = output & ":" & CStr(errNo)
End Function`);
    const actual = ev.callProcedure('Probe', []);
    assert.equal(actual, expected,
        `LongLong comparison/arithmetic ${literal}`);
}

assert.equal(cases.length, 3);
console.log(`[PASS] LongLong comparison and arithmetic matrix (${cases.length} cases)`);
