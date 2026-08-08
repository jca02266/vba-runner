import { evalVBASingle, assert } from '../../test-libs/test-runner';

const code = String.raw`
Class IStore
    Public Property Let Item(ByVal index As Long, Optional middle As Long = 7, Optional tail As Long = 9, value As String)
    End Property
End Class

Class Store
    Implements IStore
    Private mSeen As String
    Public Property Let IStore_Item(ByVal index As Long, Optional middle As Long = 7, Optional tail As Long = 9, value As String)
        mSeen = CStr(index + middle + tail) & ":" & value
    End Property
    Public Property Get Seen() As String
        Seen = mSeen
    End Property
End Class

Class IObjectStore
    Public Property Set Item(ByVal index As Long, Optional middle As Long = 7, Optional tail As Long = 9, value As Object)
    End Property
End Class

Class ObjectStore
    Implements IObjectStore
    Private mSeen As String
    Public Property Set IObjectStore_Item(ByVal index As Long, Optional middle As Long = 7, Optional tail As Long = 9, value As Object)
        mSeen = CStr(index + middle + tail) & ":" & value.Name
    End Property
    Public Property Get Seen() As String
        Seen = mSeen
    End Property
End Class

Class Payload
    Public Name As String
End Class

Function VerifyInterfaceCallByName() As String
    Dim s As New Store, i As IStore, o As New ObjectStore, oi As IObjectStore
    Dim payload As New Payload, errNo As Long, setErr As Long
    Set i = s
    Set oi = o
    payload.Name = "payload"
    On Error Resume Next
    CallByName i, "Item", VbLet, 1, , 3, "ok"
    errNo = Err.Number
    Err.Clear
    CallByName oi, "Item", VbSet, 1, , 3, payload
    setErr = Err.Number
    VerifyInterfaceCallByName = CStr(errNo) & ":" & s.Seen & ":" & CStr(setErr) & ":" & o.Seen
End Function
`;

const evaluator = evalVBASingle(code);
assert.strictEqual(
    evaluator.callProcedure('VerifyInterfaceCallByName', []),
    '0:11:ok:0:11:payload',
    'Interface参照のCallByName VbLet/VbSetがProperty value-tailへdispatchする',
);
console.log('✅ Interface CallByName Property value-tail: 全テスト通過');
