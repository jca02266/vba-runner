/** UDT配列のReDim Preserve、Erase、全体代入境界。 */
import { evalVBASingle, assert } from '../../test-libs/test-runner';

const ev = evalVBASingle(String.raw`
Private Type PointRecord
    X As Long
    Y As Long
End Type
Public preserveX, preserveY, newElement, erasedBoundErr
Sub Test()
    Dim points() As PointRecord
    ReDim points(1 To 2)
    points(1).X = 10: points(1).Y = 20
    points(2).X = 30: points(2).Y = 40
    ReDim Preserve points(1 To 3)
    preserveX = points(1).X
    preserveY = points(2).Y
    points(3).X = 50
    newElement = points(3).X
    Erase points
    On Error Resume Next
    erasedBoundErr = Err.Number
    Err.Clear
    erasedBoundErr = LBound(points)
    erasedBoundErr = Err.Number
End Sub
`);
ev.callProcedure('Test', []);

assert.strictEqual(ev.env.get('preservex'), 10, 'ReDim Preserve keeps the first UDT element');
assert.strictEqual(ev.env.get('preservey'), 40, 'ReDim Preserve keeps later UDT fields');
assert.strictEqual(ev.env.get('newelement'), 50, 'ReDim Preserve creates a writable new element');
assert.strictEqual(ev.env.get('erasedbounderr'), 9, 'LBound after Erase raises subscript error');
const invalid = evalVBASingle(String.raw`
Private Type PointRecord
    X As Long
End Type
Sub Test()
    Dim left() As PointRecord, right() As PointRecord
    left = right
End Sub
`);
let rejected = false;
try {
    invalid.callProcedure('Test', []);
} catch (error: any) {
    rejected = error?.number === 13;
}
assert.strictEqual(rejected, true, 'UDT array whole assignment is rejected at compile time');
console.log('[PASS] UDT array resize and assignment matrix');
