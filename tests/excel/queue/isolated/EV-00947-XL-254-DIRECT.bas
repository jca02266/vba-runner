Attribute VB_Name = "EV00947Direct"
Option Explicit
Public Function RunCompileProbe() As String
    Dim value As Variant
    value = NPV(0.1, Array(-100, 50, 100))
    RunCompileProbe = "XL-254 OK"
End Function
