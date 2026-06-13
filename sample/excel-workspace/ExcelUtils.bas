Attribute VB_Name = "ExcelUtils"
Option Explicit

' ════════════════════════════════════════════════════════════
' ExcelUtils — Excel VBA でよく使うユーティリティ関数集
'
' 実際の Excel VBA では WorksheetFunction を使う場面が多いが、
' ここでは同等ロジックを VBA で実装。
'
' 含む機能:
'   - AutoSort: バブルソートで行を並べ替え
'   - FindRow: 列内でキー検索
'   - RemoveDuplicates: 重複行の除去
'   - PadLeft / PadRight: 文字列整形
'   - IsDate: 日付文字列の検証
'   - WeekdayName: 曜日名変換
'   - FormatNumber: 数値フォーマット
' ════════════════════════════════════════════════════════════

' ─── バブルソート (指定列でシートの行を並べ替え) ─────────────
' Excel の並べ替え機能に相当
' 引数:
'   ws        対象シート
'   sortCol   ソートキー列 (1 始まり)
'   r1 / r2   データ行範囲 (ヘッダーを除いた行)
'   colCount  1 行あたりの列数
'   ascending True = 昇順 / False = 降順
Public Sub AutoSort(ByVal ws As SimSheet, _
                    ByVal sortCol As Long, _
                    ByVal r1 As Long, ByVal r2 As Long, _
                    ByVal colCount As Long, _
                    Optional ByVal ascending As Boolean = True)
    Dim i As Long, j As Long, c As Long
    Dim swapNeeded As Boolean

    For i = r1 To r2 - 1
        For j = r1 To r2 - (i - r1) - 1
            Dim a As Variant
            Dim b As Variant
            a = ws.GetValue(j, sortCol)
            b = ws.GetValue(j + 1, sortCol)

            If VBA.IsNumeric(a) And VBA.IsNumeric(b) Then
                swapNeeded = IIf(ascending, CDbl(a) > CDbl(b), CDbl(a) < CDbl(b))
            Else
                swapNeeded = IIf(ascending, CStr(a) > CStr(b), CStr(a) < CStr(b))
            End If

            If swapNeeded Then
                ' 行を入れ替え
                For c = 1 To colCount
                    Dim tmp As Variant
                    tmp = ws.GetValue(j, c)
                    ws.Cells(j, c).Value = ws.GetValue(j + 1, c)
                    ws.Cells(j + 1, c).Value = tmp
                Next c
            End If
        Next j
    Next i
End Sub

' ─── 行検索 (Excel の Match/VLOOKUP 相当) ───────────────────
' 戻り値: 見つかった行番号 (見つからなければ 0)
Public Function FindRow(ByVal ws As SimSheet, _
                        ByVal searchCol As Long, _
                        ByVal key As Variant, _
                        ByVal r1 As Long, ByVal r2 As Long) As Long
    Dim r As Long
    For r = r1 To r2
        If CStr(ws.GetValue(r, searchCol)) = CStr(key) Then
            FindRow = r
            Exit Function
        End If
    Next r
    FindRow = 0
End Function

' ─── CountIf 相当 ───────────────────────────────────────────
' 指定列で条件に一致する行数を返す
Public Function CountIfValue(ByVal ws As SimSheet, _
                             ByVal col As Long, _
                             ByVal criterion As Variant, _
                             ByVal r1 As Long, ByVal r2 As Long) As Long
    Dim cnt As Long
    Dim r As Long
    For r = r1 To r2
        If CStr(ws.GetValue(r, col)) = CStr(criterion) Then
            cnt = cnt + 1
        End If
    Next r
    CountIfValue = cnt
End Function

' ─── SumIf 相当 ─────────────────────────────────────────────
' 条件列が criterion に一致する行の sumCol 列合計を返す
Public Function SumIfValue(ByVal ws As SimSheet, _
                           ByVal condCol As Long, ByVal criterion As Variant, _
                           ByVal sumCol As Long, _
                           ByVal r1 As Long, ByVal r2 As Long) As Double
    Dim total As Double
    Dim r As Long
    For r = r1 To r2
        If CStr(ws.GetValue(r, condCol)) = CStr(criterion) Then
            Dim v As Variant
            v = ws.GetValue(r, sumCol)
            If VBA.IsNumeric(v) Then
                total = total + CDbl(v)
            End If
        End If
    Next r
    SumIfValue = total
End Function

' ─── 文字列左寄せパディング ─────────────────────────────────
Public Function PadRight(ByVal s As String, ByVal width As Long, _
                         Optional ByVal padChar As String = " ") As String
    Dim result As String
    result = s
    Do While Len(result) < width
        result = result & padChar
    Loop
    PadRight = Left(result, width)
End Function

' ─── 文字列右寄せパディング ─────────────────────────────────
Public Function PadLeft(ByVal s As String, ByVal width As Long, _
                        Optional ByVal padChar As String = " ") As String
    Dim result As String
    result = s
    Do While Len(result) < width
        result = padChar & result
    Loop
    PadLeft = Right(result, width)
End Function

' ─── 数値の3桁区切りフォーマット ────────────────────────────
Public Function FormatNum(ByVal n As Double) As String
    Dim intPart As Long
    Dim result As String
    intPart = CLng(Int(n))
    result = CStr(Abs(intPart))

    Dim formatted As String
    Dim pos As Long
    pos = 0
    Dim i As Long
    For i = Len(result) To 1 Step -1
        pos = pos + 1
        If pos > 1 And (pos - 1) Mod 3 = 0 Then
            formatted = "," & formatted
        End If
        formatted = Mid(result, i, 1) & formatted
    Next i

    If intPart < 0 Then formatted = "-" & formatted
    FormatNum = formatted
End Function

' ─── 日付文字列の検証 ───────────────────────────────────────
Public Function IsValidDate(ByVal s As String) As Boolean
    ' YYYY/MM/DD または YYYY-MM-DD 形式を検証
    If Len(s) <> 10 Then
        IsValidDate = False
        Exit Function
    End If
    Dim sep As String
    sep = Mid(s, 5, 1)
    If sep <> "/" And sep <> "-" Then
        IsValidDate = False
        Exit Function
    End If
    Dim yyyy As String, mm As String, dd As String
    yyyy = Left(s, 4)
    mm   = Mid(s, 6, 2)
    dd   = Right(s, 2)

    If Not VBA.IsNumeric(yyyy) Or Not VBA.IsNumeric(mm) Or Not VBA.IsNumeric(dd) Then
        IsValidDate = False
        Exit Function
    End If

    Dim m As Long, d As Long
    m = CLng(mm)
    d = CLng(dd)
    IsValidDate = (m >= 1 And m <= 12 And d >= 1 And d <= 31)
End Function

' ─── 売上レポート生成 ────────────────────────────────────────
' シートの売上データからテキストレポートを返す
Public Function GenerateSalesReport(ByVal ws As SimSheet, _
                                    ByVal headerRow As Long, _
                                    ByVal dataEnd As Long) As String
    Dim report As String
    report = "━━━━━━━━━━━━━━━━━━━━━━━━━━━━" & vbLf
    report = report & "     売 上 レ ポ ー ト" & vbLf
    report = report & "━━━━━━━━━━━━━━━━━━━━━━━━━━━━" & vbLf

    Dim header As String
    header = PadRight("日付", 12) & PadRight("商品名", 10) & _
             PadLeft("数量", 6) & PadLeft("単価", 8) & PadLeft("金額", 10)
    report = report & header & vbLf
    report = report & String(48, "-") & vbLf

    Dim totalAmount As Double
    Dim totalQty As Long
    Dim r As Long

    For r = headerRow + 1 To dataEnd
        Dim dt   As String
        Dim name As String
        Dim qty  As Variant
        Dim price As Variant
        Dim amount As Variant

        dt     = CStr(ws.GetValue(r, 1))
        name   = CStr(ws.GetValue(r, 2))
        qty    = ws.GetValue(r, 3)
        price  = ws.GetValue(r, 4)
        amount = ws.GetValue(r, 5)

        If name = "" Then GoTo NextRow

        Dim line As String
        line = PadRight(dt, 12) & PadRight(name, 10) & _
               PadLeft(CStr(qty), 6) & PadLeft(CStr(price), 8) & _
               PadLeft(CStr(amount), 10)
        report = report & line & vbLf

        If VBA.IsNumeric(amount) Then totalAmount = totalAmount + CDbl(amount)
        If VBA.IsNumeric(qty)    Then totalQty    = totalQty    + CLng(qty)
NextRow:
    Next r

    report = report & String(48, "-") & vbLf
    report = report & PadRight("合  計", 22) & _
             PadLeft(CStr(totalQty), 6) & PadLeft("", 8) & _
             PadLeft(FormatNum(totalAmount), 10) & vbLf
    report = report & "━━━━━━━━━━━━━━━━━━━━━━━━━━━━"

    GenerateSalesReport = report
End Function
