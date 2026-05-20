Attribute VB_Name = "KeyBreakModule"
Option Explicit

' ループ構造だけを担う汎用プロシージャ
' records : Object の Collection（キーでソート済み）
' body    : GetKey(rec) / OnRecord(rec) / OnBreak(key) メソッドを持つ任意のオブジェクト
Public Sub RunKeyBreak(records As Collection, body As Object)
    Dim currentKey As Variant
    Dim isFirst As Boolean
    isFirst = True

    Dim rec As Object
    For Each rec In records
        Dim key As Variant
        key = CallByName(body, "GetKey", VbMethod, rec)

        If isFirst Or key <> currentKey Then
            If Not isFirst Then CallByName body, "OnBreak", VbMethod, currentKey
            currentKey = key
            isFirst = False
        End If
        CallByName body, "OnRecord", VbMethod, rec
    Next rec

    If Not isFirst Then CallByName body, "OnBreak", VbMethod, currentKey
End Sub
