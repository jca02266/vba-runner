Function CallMsgBox() As Long
    CallMsgBox = MsgBox("Hello", 0, "Test")
End Function

Function UseRegExp() As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "\d+"
    re.Global = True
    If re.Test("abc123") Then
        UseRegExp = "matched"
    Else
        UseRegExp = "no match"
    End If
End Function
