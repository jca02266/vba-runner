import { evalVBASingle, assert } from '../../test-libs/test-runner';

const code = String.raw`
Type Point
    X As Long
End Type

Class Payload
    Public Name As String
End Class

Sub AcceptPoints(ByRef values() As Point)
    values(0).X = values(0).X + 1
End Sub

Function MakePoints() As Point()
    ReDim MakePoints(0)
    MakePoints(0).X = 42
End Function

Function MakeIntegers() As Integer()
    ReDim MakeIntegers(0)
    MakeIntegers(0) = 7
End Function

Class ArrayFactory
    Public Property Get Points() As Point()
        ReDim Points(0)
        Points(0).X = 42
    End Property

    Public Property Get Integers() As Integer()
        ReDim Integers(0)
        Integers(0) = 7
    End Property
End Class

Type UdtIntegerArray
    Values(0 To 1) As Integer
End Type

Class PrimitiveArrayHolder
    Public Values(0 To 1) As Integer
End Class

Function MakeLongs() As Long()
    Dim values(0 To 1) As Long
    MakeLongs = values
End Function

Sub AcceptObjects(ByRef values() As Object)
    If values(0) Is Nothing Then
        Err.Raise 5
    End If
End Sub

Function VerifyTypedArrays() As String
    Dim points(0 To 0) As Point, returned() As Point, wrong() As Point
    Dim objects(0 To 0) As Object, integers(0 To 0) As Integer
    Dim payload As New Payload, factory As New ArrayFactory
    Dim objectError As Long, returnError As Long, propertyError As Long
    points(0).X = 1
    Call AcceptPoints(points)
    returned = MakePoints()
    Set payload = New Payload
    payload.Name = "payload"
    Set objects(0) = payload
    Call AcceptObjects(objects)
    On Error Resume Next
    Call AcceptObjects(integers)
    objectError = Err.Number
    Err.Clear
    wrong = MakeIntegers()
    returnError = Err.Number
    Err.Clear
    wrong = factory.Integers
    propertyError = Err.Number
    VerifyTypedArrays = CStr(points(0).X) & ":" & CStr(returned(0).X) & ":" & _
        TypeName(returned) & ":" & TypeName(objects(0)) & ":" & CStr(objectError) & ":" & _
        CStr(returnError) & ":" & CStr(propertyError)
End Function

Function VerifyMemberArrayContracts() As String
    Dim u As UdtIntegerArray, h As New PrimitiveArrayHolder
    Dim errU As Long, errH As Long
    On Error Resume Next
    u.Values = MakeLongs()
    errU = Err.Number
    Err.Clear
    h.Values = MakeLongs()
    errH = Err.Number
    Err.Clear
    u.Values(0) = "12"
    h.Values(0) = "12"
    VerifyMemberArrayContracts = CStr(errU) & ":" & CStr(errH) & ":" & _
        TypeName(u.Values(0)) & ":" & TypeName(h.Values(0)) & ":" & _
        CStr(u.Values(0)) & ":" & CStr(h.Values(0))
End Function
`;

const evaluator = evalVBASingle(code);
assert.strictEqual(
    evaluator.callProcedure('VerifyTypedArrays', []),
    '2:42:Point():Payload:13:13:13',
    'UDT/Object配列の戻り値・ByRef・異型要素検証を共通binderで維持する',
);
assert.strictEqual(
    evaluator.callProcedure('VerifyMemberArrayContracts', []),
    '13:13:Double:Double:12:12',
    'UDT/Class member arrays enforce whole-array type checks and element coercion',
);
console.log('✅ UDT/Object typed-array root-cause boundary: 全テスト通過');
