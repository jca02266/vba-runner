import { strict as assert } from 'node:assert';
import { evalVBASingle } from '../../test-libs/test-runner';

const ev = evalVBASingle(String.raw`Function WeekdayDiff(ByVal date1 As String, ByVal date2 As String) As Long
    WeekdayDiff = DateDiff("w", CDate(date1), CDate(date2))
End Function`);

assert.equal(ev.callProcedure('WeekdayDiff', ['2024-01-01 00:00:00', '2024-01-07 23:00:00']), 0,
    'DateDiff w ignores time-of-day when counting the date1 weekday');
assert.equal(ev.callProcedure('WeekdayDiff', ['2024-01-01 12:00:00', '2024-01-08 00:00:00']), 1,
    'DateDiff w counts a complete calendar-week boundary');
assert.equal(ev.callProcedure('WeekdayDiff', ['2024-01-08 23:00:00', '2024-01-01 00:00:00']), -1,
    'DateDiff w preserves the reverse-direction sign');
assert.equal(ev.callProcedure('WeekdayDiff', ['2024-01-01 00:00:00', '2024-01-01 23:59:59']), 0,
    'DateDiff w returns zero within one calendar day');

console.log('✅ DateDiff weekday calendar boundaries');
