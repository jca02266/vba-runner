import { evalVBAModules, assert } from '../../test-libs/test-runner';

const modules = [
    {
        name: 'Holder',
        code: `
Class Holder
    Public Values(0 To 1) As Long
End Class
`,
    },
    {
        name: 'Module1',
        code: `
Sub AssignClassArray()
    Dim holder As New Holder, source(0 To 1) As Long
    source(0) = 1
    holder.Values = source
End Sub
`,
    },
];

const ev = evalVBAModules(modules);
assert.throwsMatch(
    () => ev.callProcedure('AssignClassArray', []),
    /error '13'/,
    'class typed-array field whole assignment must be rejected',
);
console.log('[PASS] typed class-array whole-assignment boundary');
