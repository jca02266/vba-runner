Option Explicit

Public ObjA As Object
Public Helper As Object

Sub Class_Terminate()
    If Helper Is Nothing Then Exit Sub
    Helper.TerminateCount = Helper.TerminateCount + 1
    Helper.TerminateLog = Helper.TerminateLog & "B;"
    Set ObjA = Nothing
End Sub
