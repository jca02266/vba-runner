Attribute VB_Name = "P946Shifted2D"
Option Explicit

Public Function RunCompileProbe() As String
    Dim flows(1 To 2, 5 To 6) As Double, value As Variant, errNo As Long
    flows(1, 5) = -100: flows(2, 5) = 50
    flows(1, 6) = 100: flows(2, 6) = 0
    On Error Resume Next
    Err.Clear: value = Empty
    value = MIRR(flows, 0.1, 0.12)
    errNo = Err.Number
    RunCompileProbe = "XL-248 CASE=SHIFTED_2D COMPILE=0 ERR=" & CStr(errNo) & _
        " VALUE=[" & CStr(value) & "] TYPE=" & TypeName(value)
    On Error GoTo 0
End Function
