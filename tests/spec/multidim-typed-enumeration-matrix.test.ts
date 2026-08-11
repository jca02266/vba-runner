/** Option Base・多次元配列・型付きFor Each境界。 */
import { evalVBASingle, assert } from '../../test-libs/test-runner';

const ev = evalVBASingle(String.raw`
Option Base 1
Public lower1, upper1, lower2, upper2, sumValues, countValues
Sub Test()
    Dim values(2) As Long, matrix(2, 3) As Integer, item As Variant
    values(1) = 4: values(2) = 6
    matrix(1, 1) = 1: matrix(2, 3) = 9
    lower1 = LBound(values): upper1 = UBound(values)
    lower2 = LBound(matrix, 2): upper2 = UBound(matrix, 2)
    For Each item In values
        sumValues = sumValues + item
        countValues = countValues + 1
    Next item
End Sub
`);
ev.callProcedure('Test', []);

assert.strictEqual(ev.env.get('lower1'), 1, 'Option Base 1 sets the vector lower bound');
assert.strictEqual(ev.env.get('upper1'), 2, 'Option Base 1 preserves the declared upper bound');
assert.strictEqual(ev.env.get('lower2'), 1, 'Option Base 1 applies to the second dimension');
assert.strictEqual(ev.env.get('upper2'), 3, 'Multidimensional upper bound is preserved');
assert.strictEqual(ev.env.get('sumvalues'), 10, 'For Each visits all declared elements');
assert.strictEqual(ev.env.get('countvalues'), 2, 'For Each does not include an Option Base filler');
console.log('[PASS] Multidimensional typed enumeration matrix');
