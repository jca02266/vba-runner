VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IKeyBreakBody"
Option Explicit

' グループキーを返す
Public Function GetKey(rec As Object) As Variant
End Function

' 各レコードを処理する
Public Sub OnRecord(rec As Object)
End Sub

' キーブレイク発生時（グループ終了時）の処理
Public Sub OnBreak(currentKey As Variant)
End Sub
