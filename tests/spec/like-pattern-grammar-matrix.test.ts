/** Likeパターンの空クラス・降順範囲・未終端クラス境界。 */
import { evalVBASingle, assert, vbaTrue } from '../../test-libs/test-runner';

const ev = evalVBASingle(String.raw`
Public emptyMatch, emptyErr, reversedErr, unterminatedErr
Sub Test()
    Dim value As Boolean
    On Error Resume Next
    value = ("" Like "[]")
    emptyErr = Err.Number
    emptyMatch = value
    Err.Clear
    value = ("A" Like "[Z-A]")
    reversedErr = Err.Number
    Err.Clear
    value = ("A" Like "[")
    unterminatedErr = Err.Number
End Sub
`);
ev.callProcedure('Test', []);

assert.strictEqual(ev.env.get('emptymatch'), vbaTrue,
    'an empty Like character class matches an empty string');
assert.strictEqual(ev.env.get('emptyerr'), 0,
    'an empty Like character class does not raise an error');
assert.strictEqual(ev.env.get('reversederr'), 93,
    'a descending Like range is an invalid pattern');
assert.strictEqual(ev.env.get('unterminatederr'), 93,
    'an unterminated Like character class is an invalid pattern');
console.log('[PASS] Like pattern grammar matrix');
