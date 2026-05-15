Option Explicit

' Helper class for tracking circular reference test state

Public TerminateCount As Integer
Public TerminateLog As String

Sub Reset()
    TerminateCount = 0
    TerminateLog = ""
End Sub
