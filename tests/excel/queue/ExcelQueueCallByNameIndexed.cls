VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelQueueCallByNameIndexed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mValue As Object
Private mIndex As Long

Public Property Set Item(ByVal index As Long, ByVal value As Object)
    mIndex = index
    Set mValue = value
End Property

Public Property Get Seen() As String
    If mValue Is Nothing Then
        Seen = CStr(mIndex) & ":Nothing"
    Else
        Seen = CStr(mIndex) & ":" & TypeName(mValue)
    End If
End Property
