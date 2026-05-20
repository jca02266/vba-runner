VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DeptSubtotalBody"
Option Explicit

Private mSubtotal As Long
Private mResult As String

Public Function GetKey(rec As Object) As Variant
    GetKey = rec.Department
End Function

Public Sub OnRecord(rec As Object)
    mSubtotal = mSubtotal + rec.Amount
End Sub

Public Sub OnBreak(currentKey As Variant)
    mResult = mResult & currentKey & ": " & mSubtotal & vbLf
    mSubtotal = 0
End Sub

Public Function GetResult() As String
    GetResult = mResult
End Function
