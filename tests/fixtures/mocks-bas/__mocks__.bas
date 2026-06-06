' __mocks__.bas: MsgBox / InputBox / ActiveSheet のスタブ

Function MsgBox(prompt, Optional buttons, Optional title)
    Debug.Print "[MsgBox] " & prompt
    MsgBox = 42
End Function

Function InputBox(prompt, Optional title, Optional defaultValue)
    Debug.Print "[InputBox] " & prompt
    InputBox = "mock-input"
End Function

Class MockWorksheet
    Public Property Get Name() As String
        Name = "MockSheet"
    End Property
End Class

Function ActiveSheet()
    Set ActiveSheet = New MockWorksheet
End Function
