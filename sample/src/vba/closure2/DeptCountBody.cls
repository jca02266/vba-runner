VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DeptCountBody"
Option Explicit
Implements IKeyBreakBody

Private mCount As Long
Private mResult As String

Private Function IKeyBreakBody_GetKey(rec As Object) As Variant
    IKeyBreakBody_GetKey = rec.Department
End Function

Private Sub IKeyBreakBody_OnRecord(rec As Object)
    mCount = mCount + 1
End Sub

Private Sub IKeyBreakBody_OnBreak(currentKey As Variant)
    mResult = mResult & currentKey & ": " & mCount & "件" & vbLf
    mCount = 0
End Sub

Public Function GetResult() As String
    GetResult = mResult
End Function
