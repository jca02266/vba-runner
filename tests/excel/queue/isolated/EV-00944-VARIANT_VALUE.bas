Attribute VB_Name = "P944VariantValue"
Option Explicit

Public Function RunCompileProbe() As String
    Dim flows As Variant, value As Variant, errNo As Long
    flows = Array(-100, 50, 100)
    On Error Resume Next
    Err.Clear: value = Empty
    value = MIRR(flows, 0.1, 0.12)
    errNo = Err.Number
    RunCompileProbe = "XL-246 CASE=VARIANT_VALUE COMPILE=0 ERR=" & CStr(errNo) & _
        " VALUE=[" & CStr(value) & "] TYPE=" & TypeName(value)
    On Error GoTo 0
End Function
