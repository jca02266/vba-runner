import { evalVBAModules, assert } from '../../test-libs/test-runner';

const modules = [
    { name: 'Holder', code: `
Class Holder
    Public Values(0 To 1) As Long
End Class
` },
    { name: 'Source', code: `
Class Source
    Public Other(0 To 1) As Long
End Class
` },
    { name: 'Module1', code: `
Sub CopyMemberArray()
    Dim h As New Holder, source As New Source
    h.Values = source.Other
End Sub
` },
];

const ev = evalVBAModules(modules);
assert.throwsMatch(() => ev.callProcedure('CopyMemberArray', []), /error '13'/,
    'a plain member array must not be accepted as an array return');
console.log('[PASS] member array RHS assignment boundary');
