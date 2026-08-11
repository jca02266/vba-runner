import { strict as assert } from 'node:assert';
import { evalVBAModules } from '../../test-libs/test-runner';

const cases = [
    ['1^', '1|2|2|0'],
    ['-1^', '-1|0|0|0'],
    ['&H20000000000001^', '9007199254740993|9007199254740994|9007199254740994|0'],
    ['9223372036854775807^', '9223372036854775807|9223372036854775807|9223372036854775807|6'],
] as const;

for (const [literal, expected] of cases) {
    const ev = evalVBAModules([
        { name: 'ProbeModule', code: String.raw`Option Explicit
Public Function Probe() As String
    Dim c As LongLongHolder, value As LongLong, values(0 To 0) As LongLong, errNo As Long
    Set c = New LongLongHolder
    value = ${literal}
    c.Value = value
    On Error Resume Next
    Err.Clear
    Mutate value
    errNo = Err.Number
    values(0) = value
    Probe = CStr(c.Value) & "|" & CStr(value) & "|" & CStr(values(0)) & "|" & CStr(errNo)
End Function
Public Sub Mutate(ByRef value As LongLong)
    value = value + 1
End Sub` },
        { name: 'LongLongHolder', parseAsClass: 'LongLongHolder', code: String.raw`Option Explicit
Private stored As LongLong
Public Property Let Value(ByVal value As LongLong)
    stored = value
End Property
Public Property Get Value() As LongLong
    Value = stored
End Property` },
    ]);
    assert.equal(ev.callProcedure('Probe', []), expected,
        `LongLong ByRef/Property/array writeback ${literal}`);
}

assert.equal(cases.length, 4);
console.log(`[PASS] LongLong reference writeback matrix (${cases.length} cases)`);
