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
