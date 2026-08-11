/** TimeSerial負数とDateAdd時分秒・四半期単位。 */
import { evalVBASingle, assert } from '../../test-libs/test-runner';

const ev = evalVBASingle(String.raw`
Public negSecond, negMinute, carryHour, addSeconds, addMinutes, addHours, addQuarter
Sub Test()
    negSecond = Second(TimeSerial(0, 0, -1))
    negMinute = Minute(TimeSerial(0, -1, 0))
    carryHour = Hour(TimeSerial(24, 0, 0))
    addSeconds = Second(DateAdd("s", -1, #2024/01/01 00:00:00#))
    addMinutes = Minute(DateAdd("n", -1, #2024/01/01 00:00:00#))
    addHours = Hour(DateAdd("h", -1, #2024/01/01 00:00:00#))
    addQuarter = Month(DateAdd("q", -1, #2024/05/31#))
End Sub
`);
ev.callProcedure('Test', []);

assert.strictEqual(ev.env.get('negsecond'), 1, 'TimeSerial negative second normalizes to one second');
assert.strictEqual(ev.env.get('negminute'), 1, 'TimeSerial negative minute normalizes to one minute');
assert.strictEqual(ev.env.get('carryhour'), 0, 'TimeSerial 24 hours wraps to midnight');
assert.strictEqual(ev.env.get('addseconds'), 59, 'DateAdd negative second crosses midnight');
assert.strictEqual(ev.env.get('addminutes'), 59, 'DateAdd negative minute crosses midnight');
assert.strictEqual(ev.env.get('addhours'), 23, 'DateAdd negative hour crosses midnight');
assert.strictEqual(ev.env.get('addquarter'), 2, 'DateAdd negative quarter preserves month-end clamp');
console.log('[PASS] TimeSerial signed-unit matrix');
