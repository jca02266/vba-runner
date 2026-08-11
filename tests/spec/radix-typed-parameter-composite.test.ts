import { strict as assert } from 'node:assert';
import { evalVBASingle } from '../../test-libs/test-runner';

const source = String.raw`
Option Explicit

Public Sub UpdateLongLong(ByRef value As LongLong, Optional ByVal pointer As LongPtr = 2)
    value = value + pointer
End Sub

Public Sub UpdateLongLongArray(ByRef values() As LongLong)
    values(0) = values(0) + 1
End Sub

Public Function ProbeComposite() As String
    Dim value As LongLong, errNo As Long
    value = 10
    On Error Resume Next
    Err.Clear
    UpdateLongLong pointer:=2, value:=value
    errNo = Err.Number
    ProbeComposite = "ERR=" & CStr(errNo) & ":VALUE=" & CStr(value)
    On Error GoTo 0
End Function

Public Function ProbeArray() As String
    Dim values(0 To 0) As LongLong, errNo As Long
    values(0) = 10
    On Error Resume Next
    Err.Clear
    UpdateLongLongArray values
    errNo = Err.Number
    ProbeArray = "ERR=" & CStr(errNo) & ":VALUE=" & CStr(values(0))
    On Error GoTo 0
End Function
`;

const ev = evalVBASingle(source);
assert.equal(ev.callProcedure('ProbeComposite', []), 'ERR=0:VALUE=12');
assert.equal(ev.callProcedure('ProbeArray', []), 'ERR=0:VALUE=11');

console.log('[PASS] typed parameter composite boundary (2 cases)');
