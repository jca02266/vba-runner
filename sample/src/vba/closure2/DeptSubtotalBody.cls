VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DeptSubtotalBody"
Option Explicit
Implements IKeyBreakBody

Private mSubtotal As Long
Private mResult As String

Private Function IKeyBreakBody_GetKey(rec As Object) As Variant
    IKeyBreakBody_GetKey = rec.Department
End Function

' 各レコード: 金額を累積
Private Sub IKeyBreakBody_OnRecord(rec As Object)
    mSubtotal = mSubtotal + rec.Amount
End Sub

' グループ終了: 小計を出力してリセット
Private Sub IKeyBreakBody_OnBreak(currentKey As Variant)
    mResult = mResult & currentKey & ": " & mSubtotal & vbLf
    mSubtotal = 0
End Sub

Public Function GetResult() As String
    GetResult = mResult
End Function
