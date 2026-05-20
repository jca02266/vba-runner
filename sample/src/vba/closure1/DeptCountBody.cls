VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DeptCountBody"
Option Explicit

Private mCount As Long
Private mResult As String

Public Function GetKey(rec As Object) As Variant
    GetKey = rec.Department
End Function

Public Sub OnRecord(rec As Object)
    mCount = mCount + 1
End Sub

Public Sub OnBreak(currentKey As Variant)
    mResult = mResult & currentKey & ": " & mCount & "件" & vbLf
    mCount = 0
End Sub

Public Function GetResult() As String
    GetResult = mResult
End Function
