Attribute VB_Name = "P946OneDim"
Option Explicit

Public Function RunCompileProbe() As String
    Dim flows(0 To 3) As Double, value As Variant, errNo As Long
    flows(0) = -100: flows(1) = 50: flows(2) = 100: flows(3) = 0
    On Error Resume Next
    Err.Clear: value = Empty
    value = MIRR(flows, 0.1, 0.12)
    errNo = Err.Number
    RunCompileProbe = "XL-248 CASE=ONE_DIM COMPILE=0 ERR=" & CStr(errNo) & _
        " VALUE=[" & CStr(value) & "] TYPE=" & TypeName(value)
    On Error GoTo 0
End Function
