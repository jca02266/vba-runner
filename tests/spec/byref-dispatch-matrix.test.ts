import { evalVBASingle, assert } from '../../test-libs/test-runner';

const ev = evalVBASingle(`
    Type Pair
        Left As Long
        Right As Long
    End Type

    Class Box
        Public Value As Long
        Public Sub Bump(ByRef value As Long)
            value = value + 1
        End Sub
    End Class

    Sub Bump(ByRef value As Long)
        value = value + 1
    End Sub

    Function RunMatrix() As Long
        Dim n As Long
        Dim p As Pair
        Dim b As Box
        Dim values(0) As Long
        n = 1
        values(0) = 2
        p.Left = 3
        Set b = New Box
        Bump n
        b.Bump n
        Bump ByVal n
        Bump (n)
        RunMatrix = n * 1000 + values(0) * 100 + p.Left * 10 + b.Value
    End Function
`);

// Standard/class calls, UDT member, array element, named/forced ByVal
// arguments all share the same argument binding and l-value paths. Property
// dispatch is covered by the existing class-module regression suite.
assert.strictEqual(ev.callProcedure('RunMatrix', []), 3230);
console.log('✅ ByRef dispatch matrix: module/class/UDT/array/ByVal');
