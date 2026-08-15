import { evalVBASingle, assert } from '../../test-libs/test-runner';

const mismatch = evalVBASingle(String.raw`Function TakeLong(ByRef value As Long) As Long
    value = value + 1
    TakeLong = value
End Function
Function Probe() As Long
    Dim value As Integer
    value = 10
    Probe = TakeLong(value)
End Function`);

assert.throws(
    () => mismatch.callProcedure('Probe', []),
    'ByRef scalar mismatch must be rejected',
);

const parenthesized = evalVBASingle(String.raw`Function TakeLong(ByRef value As Long) As Long
    value = value + 1
    TakeLong = value
End Function
Function Probe() As String
    Dim value As Integer
    value = 10
    Probe = CStr(TakeLong((value))) & ":" & CStr(value)
End Function`);
assert.strictEqual(parenthesized.callProcedure('Probe', []), '11:10',
    'parenthesized ByRef argument is coerced through a temporary');

const classMismatch = evalVBASingle(String.raw`Class Worker
Public Sub TakeLong(ByRef value As Long)
    value = value + 1
End Sub
End Class
Function Probe() As Long
    Dim worker As New Worker, value As Integer
    value = 10
    worker.TakeLong value
    Probe = value
End Function`);
assert.throws(
    () => classMismatch.callProcedure('Probe', []),
    'class ByRef scalar mismatch must be rejected',
);

console.log('[PASS] Scalar ByRef argument type identity');
