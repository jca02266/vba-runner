/** Unicode比較とCompare列挙値検証の共通マトリクス。 */
import { evalVBASingle, assert } from '../../test-libs/test-runner';

const ev = evalVBASingle(String.raw`
Public instrAccentText, instrAccentBinary
Public replaceAccentText, replaceAccentBinary
Public instrInvalid, replaceInvalid, strInvalid, filterInvalid
Sub Test()
    instrAccentText = InStr(1, "ÀBC", "à", 1)
    instrAccentBinary = InStr(1, "ÀBC", "à", 0)
    replaceAccentText = Replace("Àà", "à", "x", 1, -1, 1)
    replaceAccentBinary = Replace("Àà", "à", "x", 1, -1, 0)
    On Error Resume Next
    Dim ignored As Long, values As Variant
    ignored = InStr(1, "abc", "a", 2): instrInvalid = Err.Number
    Err.Clear
    replaceAccentText = Replace("abc", "a", "x", 1, -1, 2): replaceInvalid = Err.Number
    Err.Clear
    replaceAccentText = StrComp("a", "a", 2): strInvalid = Err.Number
    Err.Clear
    values = Filter(Array("abc"), "a", True, 2): filterInvalid = Err.Number
End Sub
`);
ev.callProcedure('Test', []);

assert.strictEqual(ev.env.get('instraccenttext'), 1,
    'InStr TextCompare folds accented case');
assert.strictEqual(ev.env.get('instraccentbinary'), 0,
    'InStr BinaryCompare preserves accented case');
assert.strictEqual(ev.env.get('replaceaccenttext'), 'xx',
    'Replace TextCompare folds accented case');
assert.strictEqual(ev.env.get('replaceaccentbinary'), 'Àx',
    'Replace BinaryCompare preserves accented case');
for (const key of ['instrinvalid', 'replaceinvalid', 'strinvalid', 'filterinvalid']) {
    assert.strictEqual(ev.env.get(key), 5, `${key} rejects invalid Compare`);
}
console.log('[PASS] Unicode comparison and Compare validation matrix');
