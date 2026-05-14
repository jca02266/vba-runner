' TestHelper.cls
' Helper class for tracking test state

Class TestHelper
    Public TerminateCount As Integer
    Public TerminateLog As String

    Sub Reset()
        TerminateCount = 0
        TerminateLog = ""
    End Sub
End Class
