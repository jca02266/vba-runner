import { evalVBAModules, assert } from '../../test-libs/test-runner';

const moduleA = String.raw`Option Explicit
Private Enum Status
    Current = 3
End Enum
Private value As Status
Public Function ProbeA() As String
    value = Current
    ProbeA = CStr(value) & ":" & TypeName(value) & ":" & CStr(VarType(value))
End Function
`;

const moduleB = String.raw`Option Explicit
Private Enum Status
    Current = 7
End Enum
Private value As Status
Public Function ProbeB() As String
    value = Current
    ProbeB = CStr(value) & ":" & TypeName(value) & ":" & CStr(VarType(value))
End Function
`;

const evaluator = evalVBAModules([
    { name: 'ModuleA', code: moduleA },
    { name: 'ModuleB', code: moduleB },
]);
assert.strictEqual(evaluator.callProcedure('ProbeA', []), '3:Long:3');
assert.strictEqual(evaluator.callProcedure('ProbeB', []), '7:Long:3');

console.log('✅ Private Enum members remain isolated by module');
