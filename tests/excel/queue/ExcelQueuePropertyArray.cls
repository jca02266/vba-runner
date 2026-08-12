VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelQueuePropertyArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private seen As Long

Public Property Let Value(ByRef incoming As Long)
    seen = incoming
    incoming = incoming + 1
End Property

Public Property Get SeenValue() As Long
    SeenValue = seen
End Property
