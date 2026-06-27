Option Explicit

' 売上データを集計して合計行と商品別サマリーを追記するマクロ
' 入力: Sales シート（ヘッダー行 + データ行）
'   A列: 担当者名、B列: 商品名、C列: 売上金額

Sub Main()
    Dim ws As Worksheet
    Set ws = Worksheets("Sales")

    ' --- データ末尾行を検索（xlUp パターン） ---
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(-4162).Row  ' xlUp = -4162

    If lastRow < 2 Then
        Debug.Print "データがありません"
        Exit Sub
    End If

    Debug.Print "データ行数: " & (lastRow - 1)

    ' --- 合計を計算 ---
    Dim total As Double
    Dim i As Long
    For i = 2 To lastRow
        total = total + ws.Cells(i, 3).Value
    Next i

    ' --- 合計行を追記 ---
    Dim totalRow As Long
    totalRow = lastRow + 2
    ws.Cells(totalRow, 1).Value = "合計"
    ws.Cells(totalRow, 3).Value = total
    ws.Cells(totalRow, 1).Font.Bold = True
    ws.Cells(totalRow, 3).Font.Bold = True

    Debug.Print "合計: " & total

    ' --- 商品別サマリーを Summary シートに出力 ---
    Dim wsSummary As Worksheet
    Set wsSummary = Worksheets("Summary")

    ' ヘッダー
    wsSummary.Cells(1, 1).Value = "商品名"
    wsSummary.Cells(1, 2).Value = "売上合計"
    wsSummary.Cells(1, 1).Font.Bold = True
    wsSummary.Cells(1, 2).Font.Bold = True

    ' Scripting.Dictionary で商品別集計
    Dim dic As Object
    Set dic = CreateObject("Scripting.Dictionary")

    For i = 2 To lastRow
        Dim product As String
        Dim amount As Double
        product = ws.Cells(i, 2).Value
        amount = ws.Cells(i, 3).Value
        If dic.Exists(product) Then
            dic(product) = dic(product) + amount
        Else
            dic.Add product, amount
        End If
    Next i

    ' サマリーシートへ書き出し
    Dim outRow As Long
    outRow = 2
    Dim key As Variant
    For Each key In dic.Keys
        wsSummary.Cells(outRow, 1).Value = key
        wsSummary.Cells(outRow, 2).Value = dic(key)
        outRow = outRow + 1
    Next key

    Debug.Print "商品種別数: " & dic.Count
    Debug.Print "Summary シートへの書き出し完了"
End Sub
