import { evalVBAModules, assert } from '../../test-libs/test-runner';

const classA = { name: 'ClassA', parseAsClass: 'ClassA' as const, code: 'Public Name As String' };
const classB = { name: 'ClassB', parseAsClass: 'ClassB' as const, code: 'Public Name As String' };
const moduleA = String.raw`Private item As New ClassA
Public Function ProbeA() As String
    Dim errNo As Long
    On Error Resume Next
    Set item = Nothing
    errNo = Err.Number
    ProbeA = CStr(errNo) & ":" & TypeName(item)
End Function`;
const moduleB = String.raw`Private item As New ClassB
Public Function ProbeB() As String
    Dim errNo As Long
    On Error Resume Next
    Set item = Nothing
    errNo = Err.Number
    ProbeB = CStr(errNo) & ":" & TypeName(item)
End Function`;

for (const modules of [
    [moduleA, moduleB, classA, classB],
    [moduleB, moduleA, classA, classB],
]) {
    const ev = evalVBAModules(modules.map((module, index) => ({
        name: ['ModuleA', 'ModuleB', 'ClassA', 'ClassB'][index],
        code: typeof module === 'string' ? module : module.code,
        parseAsClass: index >= 2 ? ['ClassA', 'ClassB'][index - 2] as 'ClassA' | 'ClassB' : undefined,
    })));
    assert.strictEqual(String(ev.callProcedure('ProbeA', [])), '0:ClassA');
    assert.strictEqual(String(ev.callProcedure('ProbeB', [])), '0:ClassB');
}

console.log('[PASS] As New module state is isolated by owning environment');
