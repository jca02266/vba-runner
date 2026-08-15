import { strict as assert } from 'node:assert';
import { evalVBASingle } from '../../test-libs/test-runner';

const cases = [
    [1, 1, 1], [1, 2, 53], [1, 3, 52],
    [2, 1, 1], [2, 2, 53], [2, 3, 52],
    [4, 1, 1], [4, 2, 1], [4, 3, 53],
    [6, 1, 1], [6, 2, 1], [6, 3, 1],
    [7, 1, 1], [7, 2, 52], [7, 3, 52],
] as const;

for (const [firstDay, firstWeek, expected] of cases) {
    const ev = evalVBASingle(String.raw`Option Explicit
Public Function Probe() As String
    Dim value As Date
    value = CDate("2021-01-01")
    Probe = CStr(DatePart("ww", value, ${firstDay}, ${firstWeek})) & "," & _
        CStr(Weekday(value, ${firstDay}))
End Function`);
    const weekday = firstDay === 1 ? 6 : firstDay === 2 ? 5 : firstDay === 4 ? 3 : firstDay === 6 ? 1 : 7;
    const actual = ev.callProcedure('Probe', []);
    assert.equal(actual, `${expected},${weekday}`,
        `Date week options firstday=${firstDay} firstweek=${firstWeek}`);
    console.log(`DATE-WEEK-OPTION ${firstDay},${firstWeek} => ${actual}`);
}

assert.equal(cases.length, 15);
console.log(`[PASS] Date week option matrix (${cases.length} cases)`);
