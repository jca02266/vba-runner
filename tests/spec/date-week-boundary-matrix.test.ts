import { strict as assert } from 'node:assert';
import { evalVBASingle } from '../../test-libs/test-runner';

const cases = [
    ['2020-12-28', '53,53,1,2'],
    ['2021-01-01', '53,1,5,6'],
    ['2021-01-03', '53,2,7,1'],
    ['2021-01-04', '1,2,1,2'],
    ['2022-01-01', '52,1,6,7'],
    ['2022-01-03', '1,2,1,2'],
] as const;

for (const [date, expected] of cases) {
    const ev = evalVBASingle(String.raw`Option Explicit
Public Function Probe() As String
    Dim value As Date
    value = CDate("${date}")
    Probe = CStr(DatePart("ww", value, vbMonday, vbFirstFourDays)) & "," & _
        CStr(DatePart("ww", value, vbSunday, vbFirstJan1)) & "," & _
        CStr(Weekday(value, vbMonday)) & "," & CStr(Weekday(value, vbSunday))
End Function`);
    assert.equal(ev.callProcedure('Probe', []), expected, `date week boundary ${date}`);
}

assert.equal(cases.length, 6);
console.log(`[PASS] Date week boundary matrix (${cases.length} cases)`);
