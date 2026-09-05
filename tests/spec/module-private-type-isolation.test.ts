import { evalVBAModules, assert } from '../../test-libs/test-runner';

const moduleA = String.raw`Option Explicit
Private x As Long
Public Function ProbeA() As String
    ProbeA = CStr(VarType(x)) & ":" & TypeName(x)
End Function
`;

const moduleB = String.raw`Option Explicit
Private x As String
Public Function ProbeB() As String
    ProbeB = CStr(VarType(x)) & ":" & TypeName(x)
End Function
`;

function probe(order: Array<{ name: string; code: string }>): [string, string] {
    const evaluator = evalVBAModules(order);
    return [
        evaluator.callProcedure('ProbeA', []),
        evaluator.callProcedure('ProbeB', []),
    ];
}

assert.deepStrictEqual(probe([
    { name: 'ModuleA', code: moduleA },
    { name: 'ModuleB', code: moduleB },
]), ['3:Long', '8:String']);
assert.deepStrictEqual(probe([
    { name: 'ModuleB', code: moduleB },
    { name: 'ModuleA', code: moduleA },
]), ['3:Long', '8:String']);

console.log('✅ Module-private variable type metadata is isolated by module');
