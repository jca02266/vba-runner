Function GetSheetName() As String
    Dim ws As Object
    Set ws = Sheets("Data")
    GetSheetName = ws.Name
End Function

Function GetMsgResult() As Long
    GetMsgResult = MsgBox("test")
End Function
