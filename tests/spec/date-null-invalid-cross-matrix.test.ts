/** 日付関数のNull・Empty・不正引数横断境界。 */
import { evalVBASingle, assert, vbaNull } from '../../test-libs/test-runner';

const ev = evalVBASingle(String.raw`
Public addNull, addEmptyDay, serialMonth13, serialNegativeDay
Public diffNothingErr, partInvalidErr
Sub Test()
    addNull = DateAdd("d", 1, Null)
    addEmptyDay = Day(DateAdd("d", 1, Empty))
    serialMonth13 = Month(DateSerial(2024, 13, 1))
    serialNegativeDay = Day(DateSerial(2024, 1, -1))
    On Error Resume Next
    Dim ignored As Long
    ignored = DateDiff("d", Nothing, #2024/01/01#): diffNothingErr = Err.Number
    Err.Clear
    ignored = DatePart("invalid", #2024/01/01#): partInvalidErr = Err.Number
End Sub
`);
ev.callProcedure('Test', []);

assert.strictEqual(ev.env.get('addnull'), vbaNull, 'DateAdd Null propagates Null');
assert.strictEqual(ev.env.get('addemptyday'), 31, 'DateAdd Empty uses the empty date serial');
assert.strictEqual(ev.env.get('serialmonth13'), 1, 'DateSerial month 13 advances to January');
assert.strictEqual(ev.env.get('serialnegativeday'), 30, 'DateSerial negative day rolls back');
assert.strictEqual(ev.env.get('diffnothingerr'), 91, 'DateDiff Nothing raises Error 91');
assert.strictEqual(ev.env.get('partinvaliderr'), 5, 'DatePart invalid interval raises Error 5');
console.log('[PASS] Date Null, Empty, and invalid-input matrix');
