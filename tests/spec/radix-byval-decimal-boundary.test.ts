import { strict as assert } from 'node:assert';
import { evalVBASingle } from '../../test-libs/test-runner';

const source = String.raw`
Option Explicit
Public Sub TakeLongLong(ByVal value As LongLong)
End Sub

Public Function ProbeByVal(ByVal text As String) As String
    Dim source As Variant, errNo As Long
    source = CDec(text)
    On Error Resume Next
    Err.Clear
    TakeLongLong source
    errNo = Err.Number
    ProbeByVal = "ERR=" & CStr(errNo)
    On Error GoTo 0
End Function
`;

const ev = evalVBASingle(source);
assert.equal(ev.callProcedure('ProbeByVal', ['-9223372036854775809']), 'ERR=6');
assert.equal(ev.callProcedure('ProbeByVal', ['9223372036854775807']), 'ERR=0');

console.log('[PASS] Decimal ByVal parameter width boundaries (2 cases)');
