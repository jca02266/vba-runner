Attribute VB_Name = "ExcelQueueQualifiedProperty"
Option Explicit

Public Property Get Item(ByVal index As Long) As Long
    Item = index * 10 + 5
End Property

Public Property Get Name() As String
    Name = "module-a"
End Property

Public Property Get Values() As Variant
    Values = Array(10, 20, 30)
End Property
