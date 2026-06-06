' __mocks__/ExcelObjects.bas: Excel オブジェクトスタブ

Class MockWorksheet
    Private mName As String

    Public Sub SetName(n As String)
        mName = n
    End Sub

    Public Property Get Name() As String
        Name = mName
    End Property
End Class

Function Sheets(nameOrIndex)
    Dim ws As New MockWorksheet
    ws.SetName "MockSheet"
    Set Sheets = ws
End Function
