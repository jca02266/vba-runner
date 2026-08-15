/** DateSerial繰り上げ・負数とCDate/DateValue変換。 */
import { evalVBASingle, assert } from '../../test-libs/test-runner';

const ev = evalVBASingle(String.raw`
Public d32, dMonth0, dMonthNeg, dDay0, cdateLeap, dateValueDay, dateValueHour
Sub Test()
    d32 = Day(DateSerial(2024, 1, 32))
    dMonth0 = Month(DateSerial(2024, 0, 1))
    dMonthNeg = Month(DateSerial(2024, -1, 1))
    dDay0 = Day(DateSerial(2024, 1, 0))
    cdateLeap = Day(CDate("2024/02/29"))
    dateValueDay = Day(DateValue("2024/02/29 12:34:56"))
    dateValueHour = Hour(DateValue("2024/02/29 12:34:56"))
End Sub
`);
ev.callProcedure('Test', []);

assert.strictEqual(ev.env.get('d32'), 1, 'DateSerial day overflow advances the month');
assert.strictEqual(ev.env.get('dmonth0'), 12, 'DateSerial month zero rolls back one month');
assert.strictEqual(ev.env.get('dmonthneg'), 11, 'DateSerial negative month rolls back two months');
assert.strictEqual(ev.env.get('dday0'), 31, 'DateSerial day zero returns the prior month end');
assert.strictEqual(ev.env.get('cdateleap'), 29, 'CDate accepts a leap-day string');
assert.strictEqual(ev.env.get('datevalueday'), 29, 'DateValue preserves the calendar date');
assert.strictEqual(ev.env.get('datevaluehour'), 0, 'DateValue strips the time component');
assert.throwsMatch(() => ev.evalExpression('DateSerial(2024, 1, 32768)'), /error '6'/,
    'DateSerial rejects day values outside the Integer argument range');
assert.throwsMatch(() => ev.evalExpression('DateSerial(2024, 32768, 1)'), /error '6'/,
    'DateSerial rejects month values outside the Integer argument range');
assert.throwsMatch(() => ev.evalExpression('DateSerial(2024, 1, -32769)'), /error '6'/,
    'DateSerial rejects negative day values outside the Integer argument range');
console.log('[PASS] DateSerial and date conversion matrix');
