import { strict as assert } from 'node:assert';
import { evalVBASingle } from '../../test-libs/test-runner';

const cases = [
    ['2024-01-01 00:00:00', '2024-01-02 01:02:03', '1,0,0,0,25,1502,90123,0,0|-90123'],
    ['2020-02-29 23:59:00', '2020-03-01 00:01:00', '1,1,0,0,1,2,120,0,1|-120'],
    ['2020-01-01 00:00:00', '2020-04-01 00:00:00', '91,3,1,0,2184,131040,7862400,13,13|-7862400'],
    ['2024-01-01 00:00:01', '2024-01-01 00:59:59', '0,0,0,0,0,59,3598,0,0|-3598'],
    ['2024-01-01 00:59:59', '2024-01-01 01:00:00', '0,0,0,0,1,1,1,0,0|-1'],
] as const;

for (const [from, to, expected] of cases) {
    const ev = evalVBASingle(String.raw`Option Explicit
Public Function Probe() As String
    Dim first As Date, second As Date
    first = CDate("${from}")
    second = CDate("${to}")
    Probe = CStr(DateDiff("d", first, second)) & "," & CStr(DateDiff("m", first, second)) & "," & _
        CStr(DateDiff("q", first, second)) & "," & CStr(DateDiff("yyyy", first, second)) & "," & _
        CStr(DateDiff("h", first, second)) & "," & CStr(DateDiff("n", first, second)) & "," & _
        CStr(DateDiff("s", first, second)) & "," & CStr(DateDiff("w", first, second)) & "," & _
        CStr(DateDiff("ww", first, second)) & "|" & CStr(DateDiff("s", second, first))
End Function`);
    assert.equal(ev.callProcedure('Probe', []), expected,
        `DateDiff interval/time ${from} -> ${to}`);
}

assert.equal(cases.length, 5);
console.log(`[PASS] DateDiff interval/time matrix (${cases.length} cases)`);
