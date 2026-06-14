Attribute VB_Name = "ExcelMain"
Option Explicit

' ════════════════════════════════════════════════════════════
' ExcelMain — Excel ワークブックのデモ実行エントリーポイント
'
' Auto_Open:         ワークブックオープン時の旧式自動実行マクロ (標準モジュール方式)
' RunWorkbookDemo:   ワークブック/シートイベントのデモ
' RunDataDemo:       データ操作・集計・レポートのデモ
' RunValidationDemo: 入力検証イベントのデモ
' ════════════════════════════════════════════════════════════

' ─────────────────────────────────────────────────────────
' Auto_Open — 標準モジュールに書く旧式の起動マクロ
' 実際の Excel では ThisWorkbook の Workbook_Open より先に呼ばれる
' ─────────────────────────────────────────────────────────
Public Sub Auto_Open()
    Debug.Print "[Auto_Open] ワークブックが開かれました"
    Debug.Print "[Auto_Open] 標準モジュールの Auto_Open は Excel VBA の旧式自動実行方式"
End Sub

' ─────────────────────────────────────────────────────────
' RunWorkbookDemo
' ワークブックとシートを作成し、イベント発火を確認する
' ─────────────────────────────────────────────────────────
Public Sub RunWorkbookDemo()
    Debug.Print "=========================================="
    Debug.Print " Workbook/Sheet イベントデモ"
    Debug.Print "=========================================="

    ' ── ワークブック・シートを構築 ──────────────────────────
    Dim wb As SimWorkbook
    Set wb = New SimWorkbook
    wb.Name = "SalesReport.xlsm"
    wb.Path = "C:\Users\user\Documents"

    ' シートを追加
    Dim ws1 As SimSheet
    Dim ws2 As SimSheet
    Set ws1 = wb.AddSheet("売上データ")
    Set ws2 = wb.AddSheet("集計")

    ' ── ThisWorkbook (ワークブックイベントハンドラー) ──────
    ' 実際の Excel では ThisWorkbook モジュールへの参照を持つが、
    ' ここではクラスとしてインスタンス化して接続する
    Dim twb As ThisWorkbook
    Set twb = New ThisWorkbook
    twb.AttachWorkbook wb

    ' ── Sheet1 (シートイベントハンドラー) ─────────────────
    ' Worksheet_Change / Worksheet_Activate 等を受信する
    Dim sh1 As Sheet1
    Set sh1 = New Sheet1
    sh1.AttachSheet ws1

    ' ── Auto_Open 相当の起動処理 ─────────────────────────
    ' (実際の Excel では Workbook.Open で自動発火)
    Debug.Print ""
    Auto_Open

    ' ── Workbook_Open イベント発火 ────────────────────────
    Debug.Print ""
    wb.FireOpenEvent

    ' ── シートをアクティブ化 (Worksheet_Activate) ─────────
    Debug.Print ""
    ws1.Activate

    ' ── データ入力 (シートの Change イベントが連鎖して発火) ──
    Debug.Print ""
    Debug.Print "--- データ入力 ---"

    ' SetValue 経由で入力すると Change イベントが発火する。
    ' ThisWorkbook の Workbook_SheetChange が金額を自動計算する。
    ' Sheet1 の Worksheet_Change が入力検証を行う。
    ws1.SetValue 2, 1, "2024/01/15"
    ws1.SetValue 2, 2, "緑茶"
    ws1.SetValue 2, 3, 10        ' 数量 → Workbook_SheetChange で金額計算
    ws1.SetValue 2, 4, 120       ' 単価 → 金額 = 10×120 = 1200

    ws1.SetValue 3, 1, "2024/01/15"
    ws1.SetValue 3, 2, "クッキー"
    ws1.SetValue 3, 3, 5
    ws1.SetValue 3, 4, 250       ' 金額 = 5×250 = 1250

    ws1.SetValue 4, 1, "2024/01/16"
    ws1.SetValue 4, 2, "コーヒー"
    ws1.SetValue 4, 3, 8
    ws1.SetValue 4, 4, 180       ' 金額 = 8×180 = 1440

    ' ── シート切り替え (Worksheet_Deactivate/Activate) ────
    Debug.Print ""
    ws1.Deactivate
    ws2.Activate

    ' ── 集計シートに書き込み ──────────────────────────────
    Debug.Print ""
    Debug.Print "--- 集計シートへ転記 ---"
    Dim totalSales As Double
    totalSales = ws1.SumRange(2, 5, 4, 5)
    ws2.SetValue 2, 1, "2024/01/16"
    ws2.SetValue 2, 2, totalSales
    ws2.SetValue 2, 3, ws1.CountRange(2, 2, 4, 2)

    Debug.Print "総売上: " & totalSales
    Debug.Print "件数: " & ws1.CountRange(2, 2, 4, 2)

    ' ── ワークブックを閉じる (Workbook_BeforeClose) ───────
    Debug.Print ""
    Dim closed As Boolean
    closed = wb.CloseWorkbook
    Debug.Print "ワークブックのクローズ: " & IIf(closed, "成功", "キャンセル")

    ' ── 結果確認 ──────────────────────────────────────────
    Debug.Print ""
    Debug.Print "--- 結果確認 ---"
    Debug.Print "監査ログ件数: " & twb.ChangeCount
    Debug.Print "最後のアクティブシート: " & twb.LastActiveSheet

    Set sh1 = Nothing
    Set twb = Nothing
    Set wb  = Nothing
End Sub

' ─────────────────────────────────────────────────────────
' RunDataDemo
' ソート・検索・集計・レポートのデモ
' ─────────────────────────────────────────────────────────
Public Sub RunDataDemo()
    Debug.Print ""
    Debug.Print "=========================================="
    Debug.Print " データ操作デモ"
    Debug.Print "=========================================="

    ' ── シートにデータを準備 ────────────────────────────────
    Dim ws As SimSheet
    Set ws = New SimSheet
    ws.Name = "売上データ"

    ' ヘッダー (Cells 経由で直接セット = Change イベントなし)
    ws.Cells(1, 1).Value = "日付"
    ws.Cells(1, 2).Value = "商品名"
    ws.Cells(1, 3).Value = "数量"
    ws.Cells(1, 4).Value = "単価"
    ws.Cells(1, 5).Value = "金額"

    ' データ行を 2D 配列で準備してループ設定
    Dim salesData(1 To 8, 1 To 5) As Variant
    salesData(1, 1) = "2024/01/15": salesData(1, 2) = "緑茶":    salesData(1, 3) = 10: salesData(1, 4) = 120
    salesData(2, 1) = "2024/01/15": salesData(2, 2) = "クッキー": salesData(2, 3) = 5:  salesData(2, 4) = 250
    salesData(3, 1) = "2024/01/16": salesData(3, 2) = "コーヒー": salesData(3, 3) = 8:  salesData(3, 4) = 180
    salesData(4, 1) = "2024/01/16": salesData(4, 2) = "緑茶":    salesData(4, 3) = 15: salesData(4, 4) = 120
    salesData(5, 1) = "2024/01/17": salesData(5, 2) = "シャンプー":salesData(5, 3) = 3:  salesData(5, 4) = 680
    salesData(6, 1) = "2024/01/17": salesData(6, 2) = "クッキー": salesData(6, 3) = 12: salesData(6, 4) = 250
    salesData(7, 1) = "2024/01/18": salesData(7, 2) = "チョコ":  salesData(7, 3) = 20: salesData(7, 4) = 300
    salesData(8, 1) = "2024/01/18": salesData(8, 2) = "コーヒー": salesData(8, 3) = 6:  salesData(8, 4) = 180

    Dim r As Long, c As Long
    For r = 1 To 8
        For c = 1 To 4
            ws.Cells(r + 1, c).Value = salesData(r, c)
        Next c
        ws.Cells(r + 1, 5).Value = CDbl(salesData(r, 3)) * CDbl(salesData(r, 4))
    Next r

    Dim lastRow As Long
    lastRow = ws.LastRow(2)
    Debug.Print "データ行数: " & (lastRow - 1)

    ' ── ソート (商品名昇順) ──────────────────────────────────
    Debug.Print ""
    Debug.Print "--- 商品名順ソート ---"
    AutoSort ws, 2, 2, lastRow, 5, True

    ws.PrintRange 1, 1, lastRow, 5

    ' ── 商品別集計 ──────────────────────────────────────────
    Debug.Print ""
    Debug.Print "--- 商品別集計 ---"
    Dim products(1 To 5) As String
    products(1) = "緑茶"
    products(2) = "クッキー"
    products(3) = "コーヒー"
    products(4) = "シャンプー"
    products(5) = "チョコ"

    Dim i As Long
    For i = 1 To 5
        Dim cnt    As Long
        Dim sales  As Double
        cnt   = CountIfValue(ws, 2, products(i), 2, lastRow)
        sales = SumIfValue(ws, 2, products(i), 5, 2, lastRow)
        Debug.Print PadRight(products(i), 12) & ": " & cnt & "件  ¥" & FormatNum(sales)
    Next i

    ' ── 合計 ────────────────────────────────────────────────
    Dim grandTotal As Double
    grandTotal = ws.SumRange(2, 5, lastRow, 5)
    Debug.Print ""
    Debug.Print "総売上: ¥" & FormatNum(grandTotal)

    ' ── VLookup 相当の検索 ──────────────────────────────────
    Debug.Print ""
    Debug.Print "--- VLookup: コーヒーの最初の単価 ---"
    Dim foundRow As Long
    foundRow = FindRow(ws, 2, "コーヒー", 2, lastRow)
    If foundRow > 0 Then
        Debug.Print "行 " & foundRow & " : 単価 = ¥" & ws.GetValue(foundRow, 4)
    Else
        Debug.Print "見つかりません"
    End If

    ' ── エラーハンドリング ──────────────────────────────────
    Debug.Print ""
    Debug.Print "--- エラーハンドリング ---"
    On Error Resume Next
    Dim badSheet As SimSheet
    Set badSheet = Nothing
    Dim v As Variant
    v = badSheet.GetValue(1, 1)
    If Err.Number <> 0 Then
        Debug.Print "エラー捕捉: " & Err.Number & " - " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0

    ' ── レポート出力 ────────────────────────────────────────
    Debug.Print ""
    Debug.Print "--- 売上レポート ---"
    Dim report As String
    report = GenerateSalesReport(ws, 1, lastRow)
    Debug.Print report

    Set ws = Nothing
End Sub

' ─────────────────────────────────────────────────────────
' RunValidationDemo
' Sheet1 の Worksheet_Change 入力検証デモ
' ─────────────────────────────────────────────────────────
Public Sub RunValidationDemo()
    Debug.Print ""
    Debug.Print "=========================================="
    Debug.Print " 入力検証デモ (Worksheet_Change)"
    Debug.Print "=========================================="

    Dim wb As SimWorkbook
    Set wb = New SimWorkbook
    wb.Name = "ValidationTest.xlsm"

    Dim ws As SimSheet
    Set ws = wb.AddSheet("売上データ")

    ' ヘッダー
    ws.Cells(1, 1).Value = "日付"
    ws.Cells(1, 2).Value = "商品名"
    ws.Cells(1, 3).Value = "数量"
    ws.Cells(1, 4).Value = "単価"
    ws.Cells(1, 5).Value = "金額"

    ' Sheet1 を接続 (Worksheet_Change が有効になる)
    Dim sh1 As Sheet1
    Set sh1 = New Sheet1
    sh1.AttachSheet ws

    ws.Activate

    Debug.Print ""
    Debug.Print "--- 正常データ入力 ---"
    ws.SetValue 2, 2, "緑茶"
    ws.SetValue 2, 3, 10
    ws.SetValue 2, 4, 120

    Debug.Print ""
    Debug.Print "--- 異常データ入力 ---"
    ws.SetValue 3, 2, ""         ' 空の商品名
    ws.SetValue 4, 3, -5         ' 負の数量
    ws.SetValue 5, 4, "abc"      ' 文字列の単価

    Debug.Print ""
    Debug.Print "--- 前後空白の自動除去 ---"
    ws.SetValue 6, 2, "  コーヒー  "

    Debug.Print ""
    Debug.Print "検証エラー件数: " & sh1.ValidationErrors
    Debug.Print "ステータス: " & sh1.StatusText
    Debug.Print "最終選択行: " & sh1.SelectedRow

    ' ── IsValidDate テスト ──────────────────────────────────
    Debug.Print ""
    Debug.Print "--- 日付検証 ---"
    Dim dates(1 To 4) As String
    dates(1) = "2024/01/15"
    dates(2) = "2024-12-31"
    dates(3) = "20240115"
    dates(4) = "2024/13/01"
    Dim d As Long
    For d = 1 To 4
        Debug.Print dates(d) & " → " & IIf(IsValidDate(dates(d)), "有効", "無効")
    Next d

    Set sh1 = Nothing
    Set wb  = Nothing
End Sub
