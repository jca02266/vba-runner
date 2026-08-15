import { strict as assert } from 'node:assert';
import { evalVBASingle } from '../../test-libs/test-runner';

const source = String.raw`
Public Function Probe(ByVal h As Long, ByVal m As Long, ByVal s As Long) As String
    On Error Resume Next
    Err.Clear
    Probe = CStr(TimeSerial(h, m, s)) & ":ERR=" & CStr(Err.Number)
    On Error GoTo 0
End Function
`;
const ev = evalVBASingle(source);
const cases: Array<[number, number, number, string]> = [
    [0, 0, -1, '00:00:01:ERR=0'],
    [0, -1, 0, '00:01:00:ERR=0'],
    [0, 60, 0, '01:00:00:ERR=0'],
    [24, 0, 0, '1899/12/31:ERR=0'],
    [-1, 0, 0, '01:00:00:ERR=0'],
    [0, 0, 60, '00:01:00:ERR=0'],
];
for (const [h, m, s, expected] of cases) {
    const actual = ev.callProcedure('Probe', [h, m, s]);
    assert.equal(actual, expected);
    console.log(`TIME-SERIAL ${h},${m},${s} => ${actual}`);
}
console.log(`[PASS] TimeSerial normalization boundary (${cases.length} cases)`);
