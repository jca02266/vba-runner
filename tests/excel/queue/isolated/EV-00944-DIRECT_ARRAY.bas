Attribute VB_Name = "P944DirectArray"
Option Explicit

Public Function RunCompileProbe() As String
    Dim value As Variant, errNo As Long
    On Error Resume Next
    Err.Clear: value = Empty
    value = MIRR(Array(-100, 50, 100), 0.1, 0.12)
    errNo = Err.Number
    RunCompileProbe = "XL-246 CASE=DIRECT_ARRAY COMPILE=0 ERR=" & CStr(errNo) & _
        " VALUE=[" & CStr(value) & "] TYPE=" & TypeName(value)
    On Error GoTo 0
End Function
