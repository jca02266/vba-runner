/** DateDiff/DatePart/DateAddの暦境界横断マトリクス。 */
import { evalVBASingle, assert } from '../../test-libs/test-runner';

const ev = evalVBASingle(String.raw`
Public addLeapForward, addLeapBack, diffLeap, diffWeek
Public partIsoWeek, addYearFromLeap, partQuarter
Sub Test()
    addLeapForward = Day(DateAdd("m", 1, #2024/01/31#))
    addLeapBack = Day(DateAdd("m", -1, #2024/03/31#))
    diffLeap = DateDiff("d", #2024/02/28#, #2024/03/01#)
    diffWeek = DateDiff("ww", #2023/12/30#, #2024/01/01#)
    partIsoWeek = DatePart("ww", #2024/01/01#, 2, 2)
    addYearFromLeap = Day(DateAdd("yyyy", 1, #2024/02/29#))
    partQuarter = DatePart("q", #2024/12/31#)
End Sub
`);
ev.callProcedure('Test', []);

assert.strictEqual(ev.env.get('addleapforward'), 29,
    'DateAdd month clamps January 31 to leap-year February 29');
assert.strictEqual(ev.env.get('addleapback'), 29,
    'DateAdd negative month clamps March 31 to February 29');
assert.strictEqual(ev.env.get('diffleap'), 2,
    'DateDiff day counts leap-day crossing');
assert.strictEqual(ev.env.get('diffweek'), 1,
    'DateDiff week counts a Sunday boundary');
assert.strictEqual(ev.env.get('partisoweek'), 1,
    'DatePart week honors Monday and first-four-days options');
assert.strictEqual(ev.env.get('addyearfromleap'), 28,
    'DateAdd year clamps February 29 to February 28');
assert.strictEqual(ev.env.get('partquarter'), 4, 'DatePart quarter at year end');
console.log('[PASS] Date calendar cross-function matrix');
