/** 文字列関数のNull・Empty・Compare省略境界。 */
import { evalVBASingle, assert, vbaNull } from '../../test-libs/test-runner';

const ev = evalVBASingle(String.raw`
Public instrNull, instrEmpty, replaceNullErr, replaceEmpty
Public strNull, strEmpty, filterNullErr, filterEmptyCount
Sub Test()
    instrNull = InStr("abc", Null)
    instrEmpty = InStr("abc", Empty)
    replaceEmpty = Replace("abc", Empty, "x")
    strNull = StrComp(Null, "")
    strEmpty = StrComp(Empty, "")
    On Error Resume Next
    Dim values As Variant
    values = Filter(Array("", "a"), Null, True)
    filterNullErr = Err.Number
    Err.Clear
    values = Filter(Array("", "a"), Empty, True)
    filterEmptyCount = UBound(values) + 1
End Sub
`);
ev.callProcedure('Test', []);

assert.strictEqual(ev.env.get('instrnull'), vbaNull, 'InStr Null propagates Null');
assert.strictEqual(ev.env.get('instrempty'), 1, 'InStr Empty matches an empty substring');
assert.strictEqual(ev.env.get('replaceempty'), 'abc', 'Replace Empty leaves the source unchanged');
assert.strictEqual(ev.env.get('strnull'), vbaNull, 'StrComp Null propagates Null');
assert.strictEqual(ev.env.get('strempty'), 0, 'StrComp Empty equals an empty string');
assert.strictEqual(ev.env.get('filternullerr'), 94, 'Filter Null raises Error 94');
assert.strictEqual(ev.env.get('filteremptycount'), 2, 'Filter Empty matches both entries');
console.log('[PASS] String Null, Empty, and default Compare matrix');
