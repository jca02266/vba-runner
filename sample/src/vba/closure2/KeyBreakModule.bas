Attribute VB_Name = "KeyBreakModule"
Option Explicit

' ループ構造だけを担う汎用プロシージャ
' records : Object の Collection（キーでソート済み）
' body    : IKeyBreakBody の実装クラス
Public Sub RunKeyBreak(records As Collection, body As IKeyBreakBody)
    Dim currentKey As Variant
    Dim isFirst As Boolean
    isFirst = True

    Dim rec As Object
    For Each rec In records
        Dim key As Variant
        key = body.GetKey(rec)

        If isFirst Or key <> currentKey Then
            If Not isFirst Then body.OnBreak currentKey
            currentKey = key
            isFirst = False
        End If
        body.OnRecord rec
    Next rec

    If Not isFirst Then body.OnBreak currentKey
End Sub
