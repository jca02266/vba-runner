Attribute VB_Name = "EV00947Currency"
Option Explicit
Public Function RunCompileProbe() As String
    Dim flows(0 To 2) As Currency, value As Variant
    flows(0) = -100: flows(1) = 50: flows(2) = 100
    value = IRR(flows)
    RunCompileProbe = "XL-251 OK"
End Function
