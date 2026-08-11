/** 配列戻り値とByRef再束縛の組み合わせ。 */
import { evalVBASingle, assert } from '../../test-libs/test-runner';

const ev = evalVBASingle(String.raw`
Function MakeValues() As Long()
    ReDim MakeValues(1 To 2)
    MakeValues(1) = 11
    MakeValues(2) = 22
End Function

Sub ResizeValues(ByRef values() As Long)
    ReDim Preserve values(1 To 3)
    values(3) = 33
End Sub

Public returnedFirst, returnedSecond, reboundUpper, reboundThird
Sub Test()
    Dim values() As Long, result() As Long
    result = MakeValues()
    returnedFirst = result(1): returnedSecond = result(2)
    ReDim values(1 To 2)
    values(1) = 4: values(2) = 5
    ResizeValues values
    reboundUpper = UBound(values): reboundThird = values(3)
End Sub
`);
ev.callProcedure('Test', []);

assert.strictEqual(ev.env.get('returnedfirst'), 11, 'Function array return keeps its lower element');
assert.strictEqual(ev.env.get('returnedsecond'), 22, 'Function array return keeps its upper element');
assert.strictEqual(ev.env.get('reboundupper'), 3, 'ByRef ReDim Preserve writes back the new bound');
assert.strictEqual(ev.env.get('reboundthird'), 33, 'ByRef ReDim writes back new elements');
console.log('[PASS] Array return and ByRef rebinding matrix');
