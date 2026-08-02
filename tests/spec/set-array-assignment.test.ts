import { evalVBAModules, assert } from '../../test-libs/test-runner';

const modules = [
    { name: 'Holder', code: `
Class Holder
    Public Values(0 To 1) As Long
End Class
` },
    { name: 'Module1', code: `
Sub SetArray()
    Dim holder As New Holder, source(0 To 1) As Long
    Set holder.Values = source
End Sub
Sub SetIdentifier()
    Dim source(0 To 1) As Long, target() As Long
    Set target = source
End Sub
` },
];

const ev = evalVBAModules(modules);
assert.throwsMatch(() => ev.callProcedure('SetArray', []), /error '13'/,
    'Set to a typed class array must be rejected');
assert.throwsMatch(() => ev.callProcedure('SetIdentifier', []), /error '13'/,
    'Set to a typed array identifier must be rejected');
console.log('[PASS] Set typed-array assignment boundary');
