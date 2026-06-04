Option Explicit

' ============================================================
' Excel データ処理サンプル
' 機能:
'   1. データの読み込みと集計
'   2. 重複行の削除
'   3. フィルタリング
'   4. 結果の別シートへの出力
' ============================================================


' エントリポイント: リボンのボタンや Alt+F8 から実行
Public Sub ProcessData()
    Dim wsData As Worksheet
    Dim wsResult As Worksheet

    ' シートの取得または作成
    Set wsData = GetOrCreateSheet("Data")
    Set wsResult = GetOrCreateSheet("Result")

    ' データが空なら先にサンプルデータを投入
    If wsData.Cells(2, 1).Value = "" Then
        CreateSampleData wsData
    End If

    ' 処理実行
    wsResult.Cells.Clear
    CopyFilteredData wsData, wsResult
    SummarizeByCategory wsData, wsResult

    ' 結果シートを見やすく整形
    FormatResultSheet wsResult

    MsgBox "処理完了しました。" & vbCrLf & "Result シートを確認してください。", vbInformation
End Sub

' ----------------------------------------------------------
' サンプルデータを Data シートに投入する
' ----------------------------------------------------------
Private Sub CreateSampleData(ws As Worksheet)
    Dim headers As Variant
    Dim rows As Variant
    Dim i As Long

    headers = Array("ID", "商品名", "カテゴリ", "売上", "日付")
    rows = Array( _
        Array(1, "りんご", "果物", 1200, "2026/01/05"), _
        Array(2, "バナナ", "果物", 800, "2026/01/07"), _
        Array(3, "牛乳", "乳製品", 320, "2026/01/07"), _
        Array(4, "チーズ", "乳製品", 950, "2026/01/10"), _
        Array(5, "みかん", "果物", 600, "2026/01/12"), _
        Array(6, "ヨーグルト", "乳製品", 450, "2026/01/12"), _
        Array(7, "ぶどう", "果物", 1500, "2026/01/15"), _
        Array(2, "バナナ", "果物", 800, "2026/01/07"), _
        Array(8, "バター", "乳製品", 380, "2026/01/20") _
    )

    ' ヘッダー書き込み
    Dim col As Long
    For col = 0 To UBound(headers)
        ws.Cells(1, col + 1).Value = headers(col)
    Next col

    ' データ書き込み
    For i = 0 To UBound(rows)
        Dim j As Long
        For j = 0 To UBound(rows(i))
            ws.Cells(i + 2, j + 1).Value = rows(i)(j)
        Next j
    Next i

    ' ヘッダーを太字・背景色設定
    With ws.Rows(1)
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With

    ws.Columns.AutoFit
End Sub

' ----------------------------------------------------------
' 重複 ID を除いたデータを Result シートへコピーする
' ----------------------------------------------------------
Private Sub CopyFilteredData(wsSource As Worksheet, wsDest As Worksheet)
    Dim lastRow As Long
    Dim destRow As Long
    Dim seenIDs As Object      ' Scripting.Dictionary として使用
    Dim i As Long
    Dim currentID As Variant

    lastRow = LastRow(wsSource)
    Set seenIDs = CreateObject("Scripting.Dictionary")

    ' ---- セクションタイトル ----
    wsDest.Cells(1, 1).Value = "■ 重複除去済みデータ"
    wsDest.Cells(1, 1).Font.Bold = True

    ' ヘッダーをコピー
    wsSource.Rows(1).Copy wsDest.Cells(2, 1)
    destRow = 3

    ' データ行を走査
    For i = 2 To lastRow
        currentID = wsSource.Cells(i, 1).Value
        If currentID = "" Then GoTo NextRow

        ' 重複チェック
        If Not seenIDs.exists(currentID) Then
            seenIDs.Add currentID, True
            wsSource.Rows(i).Copy wsDest.Cells(destRow, 1)
            destRow = destRow + 1
        End If

NextRow:
    Next i

    ' 空行を一つ追加してセクション区切り
    wsDest.Cells(destRow + 1, 1).Value = ""
End Sub

' ----------------------------------------------------------
' カテゴリ別集計を Result シートに追記する
' ----------------------------------------------------------
Private Sub SummarizeByCategory(wsSource As Worksheet, wsDest As Worksheet)
    Dim lastRow As Long
    Dim i As Long
    Dim cat As String
    Dim totals As Object       ' カテゴリ -> 売上合計
    Dim counts As Object       ' カテゴリ -> 件数
    Dim seenIDs As Object      ' 重複除去用

    lastRow = LastRow(wsSource)
    Set totals = CreateObject("Scripting.Dictionary")
    Set counts = CreateObject("Scripting.Dictionary")
    Set seenIDs = CreateObject("Scripting.Dictionary")

    ' 重複を除いてカテゴリ別に集計
    For i = 2 To lastRow
        Dim currentID As Variant
        currentID = wsSource.Cells(i, 1).Value
        If currentID = "" Then GoTo SkipRow
        If seenIDs.exists(currentID) Then GoTo SkipRow
        seenIDs.Add currentID, True

        cat = wsSource.Cells(i, 3).Value
        Dim sales As Double
        sales = wsSource.Cells(i, 4).Value

        If totals.exists(cat) Then
            totals(cat) = totals(cat) + sales
            counts(cat) = counts(cat) + 1
        Else
            totals.Add cat, sales
            counts.Add cat, 1
        End If
SkipRow:
    Next i

    ' 集計結果の書き込み開始行を決定
    Dim startRow As Long
    startRow = LastRow(wsDest) + 2

    wsDest.Cells(startRow, 1).Value = "■ カテゴリ別集計"
    wsDest.Cells(startRow, 1).Font.Bold = True
    startRow = startRow + 1

    ' 集計ヘッダー
    wsDest.Cells(startRow, 1).Value = "カテゴリ"
    wsDest.Cells(startRow, 2).Value = "件数"
    wsDest.Cells(startRow, 3).Value = "売上合計"
    wsDest.Cells(startRow, 4).Value = "平均売上"
    With wsDest.Rows(startRow)
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With
    startRow = startRow + 1

    ' カテゴリごとに1行出力
    Dim key As Variant
    For Each key In totals.Keys
        wsDest.Cells(startRow, 1).Value = key
        wsDest.Cells(startRow, 2).Value = counts(key)
        wsDest.Cells(startRow, 3).Value = totals(key)
        wsDest.Cells(startRow, 4).Value = totals(key) / counts(key)
        startRow = startRow + 1
    Next key
End Sub

' ----------------------------------------------------------
' Result シートの列幅・数値書式を整える
' ----------------------------------------------------------
Private Sub FormatResultSheet(ws As Worksheet)
    ws.Columns.AutoFit

    ' 売上列(D列=4列目)を通貨書式に
    ws.Columns(4).NumberFormat = "#,##0"
End Sub

' ----------------------------------------------------------
' シートを取得し、存在しなければ作成して返す
' ----------------------------------------------------------
Private Function GetOrCreateSheet(name As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(name)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.name = name
    End If

    Set GetOrCreateSheet = ws
End Function

' ----------------------------------------------------------
' 指定シートの最終データ行番号を返す
' ----------------------------------------------------------
Private Function LastRow(ws As Worksheet) As Long
    LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
End Function
