import { evalVBAModules, assert } from '../../test-libs/test-runner';

const modules = [
    { name: 'Inner', code: `
Class Inner
    Private data(0 To 1) As Long
    Public Property Get Values() As Long()
        data(0) = 7
        Values = data
    End Property
End Class
` },
    { name: 'Outer', code: `
Class Outer
    Private innerChild As Inner
    Private Sub Class_Initialize()
        Set innerChild = New Inner
    End Sub
    Public Property Get Child() As Inner
        Set Child = innerChild
    End Property
End Class
` },
    { name: 'Module1', code: `
Function ReadWith() As Long
    Dim outer As New Outer, values() As Long
    With outer.Child
        values = .Values
    End With
    ReadWith = values(0)
End Function
` },
];

const ev = evalVBAModules(modules);
assert.strictEqual(ev.callProcedure('ReadWith', []), 7,
    'With Property Get array returns remain assignable');
console.log('[PASS] With Property Get array return boundary');
