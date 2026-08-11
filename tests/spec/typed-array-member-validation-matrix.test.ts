/** UDT・クラス配列メンバーの型検証と要素変換。 */
import { evalVBASingle, assert } from '../../test-libs/test-runner';

const ev = evalVBASingle(String.raw`
Type IntegerBox
    Values(0 To 1) As Integer
End Type

Class Holder
    Public Values(0 To 1) As Integer
End Class

Function MakeLongs() As Long()
    ReDim MakeLongs(0 To 1)
    MakeLongs(0) = 1: MakeLongs(1) = 2
End Function

Public udtWholeErr, classWholeErr, udtElementType, classElementType
Sub Test()
    Dim box As IntegerBox, holder As New Holder
    On Error Resume Next
    box.Values = MakeLongs()
    udtWholeErr = Err.Number
    Err.Clear
    holder.Values = MakeLongs()
    classWholeErr = Err.Number
    Err.Clear
    box.Values(0) = "12"
    udtElementType = TypeName(box.Values(0))
    holder.Values(0) = "13"
    classElementType = TypeName(holder.Values(0))
End Sub
`);
ev.callProcedure('Test', []);

assert.strictEqual(ev.env.get('udtwholeerr'), 13, 'UDT member rejects a mismatched whole-array type');
assert.strictEqual(ev.env.get('classwholeerr'), 13, 'Class member rejects a mismatched whole-array type');
assert.strictEqual(ev.env.get('udtelementtype'), 'Double', 'UDT Integer element coercion follows VBA numeric rules');
assert.strictEqual(ev.env.get('classelementtype'), 'Double', 'Class Integer element coercion follows VBA numeric rules');
console.log('[PASS] Typed array member validation matrix');
