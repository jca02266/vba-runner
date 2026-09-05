import { evalVBASingle, assert } from '../../test-libs/test-runner';

const source = String.raw`Option Explicit

Private color As ColorKind

Private Enum ColorKind
    Red = 3
End Enum

Public Function Probe() As String
    Probe = CStr(VarType(color)) & ":" & CStr(color)
End Function
`;

const evaluator = evalVBASingle(source);
assert.strictEqual(evaluator.callProcedure('Probe', []), '3:0');

const localSource = String.raw`Option Explicit
Private Enum ColorKind
    Red = 3
End Enum

Public Function ProbeLocal() As String
    Dim localColor As ColorKind
    ProbeLocal = CStr(VarType(localColor)) & ":" & CStr(localColor)
End Function
`;
const localEvaluator = evalVBASingle(localSource);
assert.strictEqual(localEvaluator.callProcedure('ProbeLocal', []), '3:0');

console.log('✅ Enum variables use their Long default value');
