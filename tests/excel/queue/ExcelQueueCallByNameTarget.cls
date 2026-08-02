VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelQueueCallByNameTarget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mLetHits As Long
Private mSetHits As Long
Private mReference As Object

Public Property Let Scalar(ByVal value As Long)
    mLetHits = mLetHits + 1
End Property

Public Property Set Reference(ByVal value As Object)
    mSetHits = mSetHits + 1
    Set mReference = value
End Property

Public Property Get LetHits() As Long
    LetHits = mLetHits
End Property

Public Property Get SetHits() As Long
    SetHits = mSetHits
End Property
