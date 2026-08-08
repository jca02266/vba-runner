import { evalVBAModules } from '../../test-libs/test-runner';

const modules = [
  {
    name: 'Box',
    code: `Class Box
Public Seen As Long
Public Property Let Values(ByRef incoming() As Long)
  Seen = incoming(0)
End Property
End Class`,
  },
  {
    name: 'Module1',
    code: `Function Probe() As Long
Dim b As New Box, values(0 To 0) As Long
values(0) = 7
b.Values = values
Probe = b.Seen
End Function`,
  },
];

const runner = evalVBAModules(modules);
const result = runner.callProcedure('Probe', []);
if (result !== 7) throw new Error(`expected Property Let array value 7, got ${result}`);
console.log('[PASS] Property Let accepts typed array argument');

const mismatchModules = [
  {
    name: 'TypedArrayBox',
    code: String.raw`Class TypedArrayBox
Public Property Let Values(ByRef incoming() As Long)
End Property
End Class`,
  },
  {
    name: 'MismatchModule',
    code: String.raw`Sub AcceptLongs(ByRef incoming() As Long)
End Sub
Function ProbeMismatch() As String
Dim b As New TypedArrayBox, wrong(0 To 0) As Integer, values As Variant
On Error Resume Next
b.Values = wrong
Dim propertyErr As Long
propertyErr = Err.Number
Err.Clear
AcceptLongs wrong
Dim moduleErr As Long
moduleErr = Err.Number
Err.Clear
values = Array(1)
b.Values = values
Dim variantErr As Long
variantErr = Err.Number
ProbeMismatch = CStr(propertyErr) & ":" & CStr(moduleErr) & ":" & CStr(variantErr)
End Function`,
  },
];
const mismatchRunner = evalVBAModules(mismatchModules);
const mismatchResult = mismatchRunner.callProcedure('ProbeMismatch', []);
if (mismatchResult !== '13:13:13') {
  throw new Error(`expected typed-array mismatches to reject, got ${mismatchResult}`);
}
console.log('[PASS] Property Let and ByRef reject mismatched array element types');

const eraseModules = [
  {
    name: 'EraseArrayModule',
    code: String.raw`Sub AcceptLongs(ByRef incoming() As Long)
End Sub
Function ProbeErase() As String
Dim values() As Long, valuesAsVariant As Variant
ReDim values(0 To 0)
Erase values
On Error Resume Next
AcceptLongs values
Dim directErr As Long
directErr = Err.Number
Err.Clear
valuesAsVariant = values
AcceptLongs valuesAsVariant
Dim variantErr As Long
variantErr = Err.Number
Err.Clear
Dim ignored As Long
ignored = values(0)
Dim readErr As Long
readErr = Err.Number
Err.Clear
values(0) = 4
Dim writeErr As Long
writeErr = Err.Number
Err.Clear
ignored = valuesAsVariant(0)
Dim variantReadErr As Long
variantReadErr = Err.Number
ProbeErase = CStr(directErr) & ":" & CStr(variantErr) & ":" & CStr(readErr) & ":" & CStr(writeErr) & ":" & CStr(variantReadErr)
End Function`,
  },
];
const eraseRunner = evalVBAModules(eraseModules);
const eraseResult = eraseRunner.callProcedure('ProbeErase', []);
if (eraseResult !== '0:0:9:9:9') {
  throw new Error(`expected Erase to preserve typed array binding, got ${eraseResult}`);
}
console.log('[PASS] Erase preserves typed array parameter compatibility');

const optionalPropertyModules = [
  {
    name: 'OptionalPropertyBox',
    code: String.raw`Class OptionalPropertyBox
Private total As Long
Public Property Let Item(Optional index1 As Long = 0, Optional index2 As Long = 0, ByVal value As Long)
  total = index1 * 100 + index2 * 10 + value
End Property
Public Function Read() As Long
  Read = total
End Function
End Class`,
  },
  {
    name: 'OptionalPropertyModule',
    code: String.raw`Function ProbeOptionalProperty() As String
Dim b As New OptionalPropertyBox, e As Long
On Error Resume Next
b.Item(, 2) = 3
e = Err.Number
ProbeOptionalProperty = "pos=" & CStr(e) & ":" & CStr(b.Read())
Err.Clear
b.Item(index2:=2) = 3
e = Err.Number
ProbeOptionalProperty = ProbeOptionalProperty & "|named=" & CStr(e) & ":" & CStr(b.Read())
End Function`,
  },
];
const optionalPropertyRunner = evalVBAModules(optionalPropertyModules);
const optionalPropertyResult = optionalPropertyRunner.callProcedure('ProbeOptionalProperty', []);
if (optionalPropertyResult !== 'pos=0:23|named=0:23') {
  throw new Error(`expected Property Let Optional holes to bind defaults, got ${optionalPropertyResult}`);
}
console.log('[PASS] Property Let binds named Optional holes before value');
