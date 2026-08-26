Attribute VB_Name = "EV00947Variant"
Option Explicit
Public Function RunCompileProbe() As String
    Dim flows As Variant, value As Variant
    flows = Array(-100, 50, 100)
    value = IRR(flows)
    RunCompileProbe = "XL-253 OK"
End Function
