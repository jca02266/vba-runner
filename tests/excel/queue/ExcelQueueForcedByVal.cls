VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelQueueForcedByVal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ExcelQueueForcedByValInterface

Public Sub ExcelQueueForcedByValInterface_Mutate(ByRef value As Long)
    value = value + 1
End Sub
