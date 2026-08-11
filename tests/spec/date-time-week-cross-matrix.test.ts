/** 時刻を含むDateDiff、週境界、TimeValue正規化。 */
import { evalVBASingle, assert } from '../../test-libs/test-runner';

const ev = evalVBASingle(String.raw`
Public sameDay, hourSpan, minuteSpan, secondSpan, weekSpan
Public timeHour, timeMinute, timeSecond, invalidTimeErr
Sub Test()
    sameDay = DateDiff("d", #2024/01/01 00:00:01#, #2024/01/01 23:59:59#)
    hourSpan = DateDiff("h", #2024/01/01 00:00:01#, #2024/01/01 02:00:00#)
    minuteSpan = DateDiff("n", #2024/01/01 00:00:01#, #2024/01/01 00:02:00#)
    secondSpan = DateDiff("s", #2024/01/01 00:00:01#, #2024/01/01 00:02:00#)
    weekSpan = DateDiff("ww", #2024/01/06#, #2024/01/08#)
    timeHour = Hour(TimeValue("12:34:56"))
    timeMinute = Minute(TimeValue("12:34:56"))
    timeSecond = Second(TimeValue("12:34:56"))
    On Error Resume Next
    Dim ignored As Date
    ignored = TimeValue("24:00:00")
    invalidTimeErr = Err.Number
End Sub
`);
ev.callProcedure('Test', []);

assert.strictEqual(ev.env.get('sameday'), 0, 'DateDiff day ignores time-of-day');
assert.strictEqual(ev.env.get('hourspan'), 2, 'DateDiff hour counts the hour boundary');
assert.strictEqual(ev.env.get('minutespan'), 2, 'DateDiff minute counts minute boundaries');
assert.strictEqual(ev.env.get('secondspan'), 119, 'DateDiff second preserves sub-minute offset');
assert.strictEqual(ev.env.get('weekspan'), 1, 'DateDiff week counts the Sunday boundary');
assert.strictEqual(ev.env.get('timehour'), 12, 'TimeValue hour');
assert.strictEqual(ev.env.get('timeminute'), 34, 'TimeValue minute');
assert.strictEqual(ev.env.get('timesecond'), 56, 'TimeValue second');
assert.strictEqual(ev.env.get('invalidtimeerr'), 13, 'TimeValue rejects 24:00:00');
console.log('[PASS] Date time and week cross-function matrix');
