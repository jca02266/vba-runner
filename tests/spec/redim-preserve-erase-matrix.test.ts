/** 動的配列のReDim Preserve・Erase・下限変更境界。 */
import { evalVBASingle, assert } from '../../test-libs/test-runner';

const ev = evalVBASingle(String.raw`
Public kept, expanded, rebuilt, lowerChangeErr, firstDimErr
Sub Test()
    Dim values() As Long, matrix() As Long
    ReDim values(2 To 3)
    values(2) = 20: values(3) = 30
    ReDim Preserve values(2 To 5)
    kept = values(2): expanded = UBound(values)
    Erase values
    ReDim values(0 To 1)
    values(0) = 7
    rebuilt = values(0)
    On Error Resume Next
    ReDim Preserve values(1 To 2)
    lowerChangeErr = Err.Number
    Err.Clear
    ReDim matrix(0 To 1, 0 To 1)
    ReDim Preserve matrix(0 To 2, 0 To 1)
    firstDimErr = Err.Number
End Sub
`);
ev.callProcedure('Test', []);

assert.strictEqual(ev.env.get('kept'), 20, 'ReDim Preserve retains values at the lower bound');
assert.strictEqual(ev.env.get('expanded'), 5, 'ReDim Preserve expands the last dimension');
assert.strictEqual(ev.env.get('rebuilt'), 7, 'Erase permits a fresh ReDim');
assert.strictEqual(ev.env.get('lowerchangeerr'), 9, 'ReDim Preserve rejects a lower-bound change');
assert.strictEqual(ev.env.get('firstdimerr'), 9, 'ReDim Preserve rejects changing a non-final dimension');
console.log('[PASS] ReDim Preserve and Erase matrix');
