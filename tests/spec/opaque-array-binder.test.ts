import { evalVBASingle, assert } from '../../test-libs/test-runner';

const code = String.raw`
Sub Accept(ByRef incoming() As Long)
    Dim n As Long
    n = UBound(incoming)
End Sub

Class Holder
    Private mSize As Long
    Public Property Let Values(ByRef incoming() As Long)
        mSize = UBound(incoming)
    End Property
    Public Property Get Size() As Long
        Size = mSize
    End Property
End Class

Function ProbeBinders() As String
    Dim source As Object, holder As New Holder
    Dim col As New Collection, dict As Object, errNo As Long
    Set source = CreateObject("Test.Array")
    Set dict = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    Accept source.Values
    errNo = Err.Number
    ProbeBinders = "B" & CStr(errNo)
    Err.Clear
    holder.Values = source.Values
    ProbeBinders = ProbeBinders & ";P" & CStr(Err.Number)
    Err.Clear
    col.Add source.Values
    Accept col(1)
    ProbeBinders = ProbeBinders & ";C" & CStr(Err.Number)
    Err.Clear
    dict.Add "k", source.Values
    Accept dict("k")
    ProbeBinders = ProbeBinders & ";D" & CStr(Err.Number)
End Function
`;

const evaluator = evalVBASingle(code);
evaluator.registerComObject(() => ({
    __progId__: 'Test.Array',
    Values: [7],
}) as any);

assert.strictEqual(
    evaluator.callProcedure('ProbeBinders', []),
    'B13;P13;C13;D13',
    'Opaque arrays require element metadata at ByRef, Property, and container binders',
);
console.log('[PASS] Opaque array binder element validation');
