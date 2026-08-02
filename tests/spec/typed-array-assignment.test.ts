import { evalVBASingle, assert } from '../../test-libs/test-runner';

// A typed array cannot receive another whole array in VBA.  The destination
// must be a Variant; typed-array element assignment and ByRef ReDim remain
// valid and are covered by the existing array tests.
const code = `
Function MakeLongs() As Long()
    Dim values(0 To 1) As Long
    values(0) = 10
    values(1) = 20
    MakeLongs = values
End Function

Sub AssignTypedArray()
    Dim source() As Long, target() As Long
    source = MakeLongs()
    target = source
End Sub

Function AssignToVariant() As Long
    Dim source() As Long, value As Variant
    source = MakeLongs()
    value = source
    AssignToVariant = value(1)
End Function
`;

const ev = evalVBASingle(code);
assert.throwsMatch(
    () => ev.callProcedure('AssignTypedArray', []),
    /error '13'/,
    'typed whole-array assignment must be rejected',
);
assert.strictEqual(ev.callProcedure('AssignToVariant', []), 20,
    'Variant whole-array assignment remains valid');
console.log('[PASS] typed array whole-assignment boundary');
