import { strict as assert } from 'node:assert';
import { evalVBASingle } from '../../test-libs/test-runner';

const ev = evalVBASingle(String.raw`
Class Counter
    Public Sub Add(Optional ByRef amount As Long = 1, Optional ByRef threshold As Long = 10)
        amount = amount + 1
        threshold = threshold + amount
    End Sub
End Class
Function ProbeNamed() As String
    Dim c As New Counter, amount As Long, threshold As Long
    amount = 3: threshold = 5
    c.Add threshold:=threshold, amount:=amount
    ProbeNamed = CStr(amount) & ":" & CStr(threshold)
End Function
Function ProbeGap() As String
    Dim c As New Counter, amount As Long, threshold As Long
    amount = 12
    c.Add amount:=amount
    ProbeGap = CStr(amount) & ":" & CStr(threshold)
End Function
Function ProbeOmitted() As String
    Dim c As New Counter, amount As Long, threshold As Long
    threshold = 12
    c.Add threshold:=threshold
    ProbeOmitted = CStr(amount) & ":" & CStr(threshold)
End Function
Function ProbeArrayElements() As String
    Dim c As New Counter, values(0 To 1) As Long
    values(0) = 5: values(1) = 3
    c.Add threshold:=values(0), amount:=values(1)
    ProbeArrayElements = CStr(values(0)) & ":" & CStr(values(1))
End Function`);

assert.equal(ev.callProcedure('ProbeNamed', []), '4:9');
assert.equal(ev.callProcedure('ProbeGap', []), '13:0');
assert.equal(ev.callProcedure('ProbeOmitted', []), '0:14');
assert.equal(ev.callProcedure('ProbeArrayElements', []), '9:4');
console.log('[PASS] Property named optional ByRef matrix (4 boundaries)');
