VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelQueueDefaultValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mValue As Variant

Public Property Get Value() As Variant
Attribute Value.VB_UserMemId = 0
    Value = mValue
End Property

Public Property Let Value(ByVal newValue As Variant)
    mValue = newValue
End Property
