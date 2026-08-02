import { evalVBAModules, assert } from '../../test-libs/test-runner';

const modules = [
    { name: 'Inner', code: `
Class Inner
    Public Property Get Values() As Long()
        Dim items(0 To 0) As Long
        items(0) = 9
        Values = items
    End Property
End Class
` },
    { name: 'Outer', code: `
Class Outer
    Private innerChild As Inner
    Public Hits As Long
    Private Sub Class_Initialize()
        Set innerChild = New Inner
        Hits = 0
    End Sub
    Public Property Get Child() As Inner
        Hits = Hits + 1
        Set Child = innerChild
    End Property
End Class
` },
    { name: 'Module1', code: `
Function CountGetter() As Long
    Dim outer As New Outer, values() As Long
    values = outer.Child.Values
    CountGetter = outer.Hits * 100 + values(0)
End Function
` },
];

const ev = evalVBAModules(modules);
assert.strictEqual(ev.callProcedure('CountGetter', []), 109,
    'nested getter must be evaluated once');
console.log('[PASS] nested array getter evaluation count');
