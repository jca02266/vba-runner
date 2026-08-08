import { evalVBASingle, assert } from '../../test-libs/test-runner';

const code = String.raw`
Class Payload
    Public Name As String
End Class

Class Probe
    Private mSeen As String

    Public Function Method(a As Long, Optional middle As Long = 7, Optional tail As Long = 9) As Long
        Method = a + middle + tail
    End Function

    Public Property Get Value(a As Long, Optional middle As Long = 7, Optional tail As Long = 9) As Long
        Value = a + middle + tail
    End Property

    Public Property Let Label(a As Long, Optional middle As Long = 7, Optional tail As Long = 9, value As String)
        mSeen = CStr(a + middle + tail) & ":" & value
    End Property

    Public Property Set Item(a As Long, Optional middle As Long = 7, Optional tail As Long = 9, value As Object)
        mSeen = CStr(a + middle + tail) & ":" & value.Name
    End Property

    Public Property Get Seen() As String
        Seen = mSeen
    End Property

    Public Sub Mutate(ByRef target As Long, Optional middle As Long = 7, Optional tail As Long = 9)
        target = target + middle + tail
    End Sub
End Class

Function VerifyCallByNameOptional() As String
    Dim p As New Probe, payload As New Payload, target As Long
    payload.Name = "payload"
    target = 1
    VerifyCallByNameOptional = CStr(CallByName(p, "Method", VbMethod, 1, , 3)) & "|"
    VerifyCallByNameOptional = VerifyCallByNameOptional & CStr(CallByName(p, "Value", VbGet, 1, , 3)) & "|"
    CallByName p, "Label", VbLet, 1, , 3, "label"
    VerifyCallByNameOptional = VerifyCallByNameOptional & p.Seen & "|"
    CallByName p, "Item", VbSet, 1, , 3, payload
    VerifyCallByNameOptional = VerifyCallByNameOptional & p.Seen & "|"
    CallByName p, "Mutate", VbMethod, target, , 3
    VerifyCallByNameOptional = VerifyCallByNameOptional & CStr(target)
End Function
`;

const evaluator = evalVBASingle(code);
assert.strictEqual(
    evaluator.callProcedure('VerifyCallByNameOptional', []),
    '11|11|11:label|11:payload|11',
    'CallByNameの各呼出し種別でOptional省略とByRef書戻しを共通契約どおり処理する',
);
console.log('✅ CallByName Optional/ByRef binder 横展開: 全テスト通過');
