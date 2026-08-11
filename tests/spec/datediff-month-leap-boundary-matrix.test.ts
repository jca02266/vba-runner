import { strict as assert } from 'node:assert';
import { evalVBASingle } from '../../test-libs/test-runner';

const cases = [
    ['2020-02-28', '2020-03-01', 'd', '2'],
    ['2020-03-01', '2020-02-28', 'd', '-2'],
    ['2020-01-31', '2020-02-29', 'm', '1'],
    ['2020-02-29', '2020-03-31', 'm', '1'],
    ['2019-12-31', '2020-01-01', 'yyyy', '1'],
    ['2020-12-27', '2021-01-03', 'ww', '1'],
] as const;

for (const [from, to, interval, expected] of cases) {
    const weekArgs = interval === 'ww' ? ', vbSunday, vbFirstJan1' : '';
    const ev = evalVBASingle(String.raw`Option Explicit
Public Function Probe() As String
    Probe = CStr(DateDiff("${interval}", CDate("${from}"), CDate("${to}")${weekArgs}))
End Function`);
    assert.equal(String(ev.callProcedure('Probe', [])), expected,
        `DateDiff ${interval} ${from} -> ${to}`);
}

assert.equal(cases.length, 6);
console.log(`[PASS] DateDiff month/leap boundary matrix (${cases.length} cases)`);
