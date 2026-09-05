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

const priorSource = String.raw`Option Explicit
Private Enum ColorKind
    Red = 3
End Enum

Private color As ColorKind

Public Function ProbePrior() As String
    ProbePrior = CStr(VarType(color)) & ":" & CStr(color)
End Function
`;
const priorEvaluator = evalVBASingle(priorSource);
assert.strictEqual(priorEvaluator.callProcedure('ProbePrior', []), '3:0');

const arraySource = String.raw`Option Explicit
Private Enum ColorKind
    Red = 3
End Enum
Private colors(0 To 1) As ColorKind

Public Function ProbeArray() As String
    ProbeArray = CStr(VarType(colors)) & ":" & CStr(VarType(colors(0))) & ":" & CStr(colors(0))
End Function
`;
const arrayEvaluator = evalVBASingle(arraySource);
assert.strictEqual(arrayEvaluator.callProcedure('ProbeArray', []), '8195:3:0');

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
