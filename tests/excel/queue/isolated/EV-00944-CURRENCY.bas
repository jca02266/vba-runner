Attribute VB_Name = "P944Currency"
Option Explicit

Public Function RunCompileProbe() As String
    ' Expected Excel result: compile error 13, array or user-defined type required.
    ' The procedure is intentionally not executable; the isolated harness records
    ' the compile failure before Application.Run can enter this function.
    Dim flows(0 To 2) As Currency, value As Variant
    flows(0) = -100: flows(1) = 50: flows(2) = 100
    value = MIRR(flows, 0.1, 0.12)
    RunCompileProbe = "XL-246 CASE=CURRENCY COMPILE=0 VALUE=[" & CStr(value) & "]"
End Function
