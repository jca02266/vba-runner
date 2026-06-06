Function CallMsgBox() As Long
    CallMsgBox = MsgBox("Hello", 0, "Test")
End Function

Function CallInputBox() As String
    CallInputBox = InputBox("Enter value", , "default")
End Function

Function UseSheet() As String
    Dim ws As Object
    Set ws = ActiveSheet()
    UseSheet = ws.Name
End Function
