import { evalVBAModules, assert } from '../../test-libs/test-runner';

const modules = [
    { name: 'Inner', code: `
Class Inner
    Public Property Get Values() As Long()
        Dim items(0 To 0) As Long
        items(0) = 7
        Values = items
    End Property
End Class
` },
    { name: 'Outer', code: `
Class Outer
    Private innerChild As Inner
    Private Sub Class_Initialize()
        Set innerChild = New Inner
    End Sub
    Public Property Get Child() As Object
        Set Child = innerChild
    End Property
End Class
` },
    { name: 'Module1', code: `
Function ReadDynamic() As Long
    Dim outer As New Outer, values() As Long
    values = outer.Child.Values
    ReadDynamic = values(0)
End Function
` },
];

const ev = evalVBAModules(modules);
assert.strictEqual(ev.callProcedure('ReadDynamic', []), 7,
    'dynamic owner still exposes a valid array getter result');
console.log('[PASS] dynamic array getter return boundary');
