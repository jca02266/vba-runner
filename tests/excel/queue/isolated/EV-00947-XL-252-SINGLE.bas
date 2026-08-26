Attribute VB_Name = "EV00947Single"
Option Explicit
Public Function RunCompileProbe() As String
    Dim flows(0 To 2) As Single, value As Variant
    flows(0) = -100: flows(1) = 50: flows(2) = 100
    value = NPV(0.1, flows)
    RunCompileProbe = "XL-252 OK"
End Function
