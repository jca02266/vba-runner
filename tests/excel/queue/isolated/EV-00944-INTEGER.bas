Attribute VB_Name = "P944Integer"
Option Explicit

Public Function RunCompileProbe() As String
    Dim flows(0 To 2) As Integer, value As Variant, errNo As Long
    flows(0) = -100: flows(1) = 50: flows(2) = 100
    On Error Resume Next
    Err.Clear: value = Empty
    value = MIRR(flows, 0.1, 0.12)
    errNo = Err.Number
    RunCompileProbe = "XL-246 CASE=INTEGER COMPILE=0 ERR=" & CStr(errNo) & _
        " VALUE=[" & CStr(value) & "] TYPE=" & TypeName(value)
    On Error GoTo 0
End Function
