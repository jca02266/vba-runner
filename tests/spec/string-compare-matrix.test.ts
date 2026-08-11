/** 文字列組み込み関数のCompareモード・Unicode境界。 */
import { evalVBASingle, assert } from '../../test-libs/test-runner';

const ev = evalVBASingle(String.raw`
Public instrText, instrBinary, replaceText, replaceBinary
Public strText, strBinary, filterTextCount, filterBinaryCount
Public instrInvalidErr
Sub Test()
    instrText = InStr(1, "AbC", "B", 1)
    instrBinary = InStr(1, "AbC", "B", 0)
    replaceText = Replace("aBab", "B", "x", 1, -1, 1)
    replaceBinary = Replace("aBab", "B", "x", 1, -1, 0)
    strText = StrComp("A", "a", 1)
    strBinary = StrComp("A", "a", 0)
    filterTextCount = UBound(Filter(Array("Alpha", "beta", "GAMMA"), "A", True, 1)) + 1
    filterBinaryCount = UBound(Filter(Array("Alpha", "beta", "GAMMA"), "A", True, 0)) + 1
    On Error Resume Next
    Dim ignored As Long
    ignored = InStr(1, "AbC", "b", 2)
    instrInvalidErr = Err.Number
End Sub
`);
ev.callProcedure('Test', []);

assert.strictEqual(ev.env.get('instrtext'), 2, 'InStr TextCompare ignores case');
assert.strictEqual(ev.env.get('instrbinary'), 0, 'InStr BinaryCompare preserves case');
assert.strictEqual(ev.env.get('instrinvaliderr'), 5, 'InStr rejects invalid Compare values');
assert.strictEqual(ev.env.get('replacetext'), 'axax', 'Replace TextCompare ignores case');
assert.strictEqual(ev.env.get('replacebinary'), 'axab', 'Replace BinaryCompare preserves case');
assert.strictEqual(ev.env.get('strtext'), 0, 'StrComp TextCompare treats case as equal');
assert.ok((ev.env.get('strbinary') as number) !== 0,
    'StrComp BinaryCompare distinguishes case');
assert.strictEqual(ev.env.get('filtertextcount'), 3, 'Filter TextCompare finds all case variants');
assert.strictEqual(ev.env.get('filterbinarycount'), 2, 'Filter BinaryCompare finds exact case');
console.log('[PASS] String Compare and Option Compare matrix');
