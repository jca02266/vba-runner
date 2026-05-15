Option Explicit

' Test class for circular reference testing (Class B)

Public RefA As Object
Public Helper As Object

Sub Class_Terminate()
    If Helper Is Nothing Then Exit Sub
    Helper.TerminateCount = Helper.TerminateCount + 1
    Helper.TerminateLog = Helper.TerminateLog & "B;"
    Set RefA = Nothing
End Sub
