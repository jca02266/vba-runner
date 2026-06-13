VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Option Explicit

' ════════════════════════════════════════════════════════════
' ThisWorkbook — Excel の ThisWorkbook モジュールを模倣
'
' 実際の Excel では、このモジュールに Workbook_Open 等のイベント
' ハンドラーを書く。ここでは Private WithEvents で SimWorkbook を
' 監視し、同等のイベントを受信する。
'
' 主な構文デモ:
'   - Private WithEvents によるワークブックイベント監視
'   - Workbook_Open: 初期化処理 (シート準備・ヘッダー記入)
'   - Workbook_BeforeClose: 未保存確認
'   - Workbook_SheetChange: 変更監視・監査ログ
'   - Workbook_SheetActivate: アクティブシート切り替え検知
' ════════════════════════════════════════════════════════════

Private WithEvents m_wb     As SimWorkbook
Private WithEvents m_ws1    As SimSheet     ' Sheet1 の変更も直接監視
Private m_auditLog          As String
Private m_changeCount       As Long
Private m_lastActiveSheet   As String

' ─── ワークブックを紐付け ────────────────────────────────────
Public Sub AttachWorkbook(ByVal wb As SimWorkbook)
    Set m_wb     = wb
    m_auditLog   = ""
    m_changeCount = 0
End Sub

Public Sub AttachSheet1(ByVal ws As SimSheet)
    Set m_ws1 = ws
End Sub

' ─── Workbook_Open イベント ──────────────────────────────────
Private Sub m_wb_Open()
    Debug.Print "=== Workbook_Open: " & m_wb.Name & " ==="

    ' ヘッダー行のセットアップ (Excel でよくある起動時初期化)
    Dim ws As SimSheet
    Set ws = m_wb.Sheets("売上データ")
    ws.Cells(1, 1).Value = "日付"
    ws.Cells(1, 2).Value = "商品名"
    ws.Cells(1, 3).Value = "数量"
    ws.Cells(1, 4).Value = "単価"
    ws.Cells(1, 5).Value = "金額"

    Set ws = m_wb.Sheets("集計")
    ws.Cells(1, 1).Value = "集計日時"
    ws.Cells(1, 2).Value = "総売上"
    ws.Cells(1, 3).Value = "件数"

    Debug.Print "ヘッダー初期化完了"
End Sub

' ─── Workbook_BeforeClose イベント ──────────────────────────
Private Sub m_wb_BeforeClose(ByRef Cancel As Boolean)
    Debug.Print "=== Workbook_BeforeClose ==="
    If Not m_wb.Saved Then
        ' 実際の Excel では MsgBox で確認するが、ここでは自動保存
        Debug.Print "未保存の変更あり — 自動保存します"
        m_wb.Saved = True
    End If
    If m_changeCount > 0 Then
        Debug.Print "監査ログ: " & m_changeCount & " 件の変更を記録しました"
    End If
    Cancel = False     ' False = 閉じる, True = キャンセル
End Sub

' ─── Workbook_SheetChange イベント ──────────────────────────
Private Sub m_wb_SheetChange(ByVal Sh As SimSheet, ByVal Target As SimCell)
    m_changeCount = m_changeCount + 1
    Dim entry As String
    entry = "[" & m_changeCount & "] " & Sh.Name & "!" & Target.Address & _
            " = " & CStr(Target.Value)
    m_auditLog = m_auditLog & entry & vbLf
    Debug.Print "変更検知: " & entry
End Sub

' ─── Workbook_SheetActivate イベント ────────────────────────
Private Sub m_wb_SheetActivate(ByVal Sh As SimSheet)
    m_lastActiveSheet = Sh.Name
    Debug.Print "シート切り替え → " & Sh.Name
End Sub

' ─── Sheet1 独自変更ハンドラー ──────────────────────────────
' ThisWorkbook から直接 Sheet1 の変更も監視できる (実際の Excel でも可能)
Private Sub m_ws1_Change(ByVal Target As SimCell)
    ' ワークブックレベルの SheetChange イベントにバブリング
    ' (実際の Excel ではシートの Change が自動的に Workbook_SheetChange を呼ぶ)
    m_wb.NotifySheetChange m_ws1, Target

    ' 金額列 (E列 = 5) の自動計算: 数量 × 単価
    If Target.Column = 3 Or Target.Column = 4 Then
        Dim r As Long
        r = Target.Row
        If r > 1 Then     ' ヘッダー行はスキップ
            Dim qty As Variant
            Dim price As Variant
            qty   = m_ws1.GetValue(r, 3)
            price = m_ws1.GetValue(r, 4)
            If VBA.IsNumeric(qty) And VBA.IsNumeric(price) Then
                m_ws1.Cells(r, 5).Value = CDbl(qty) * CDbl(price)
                Debug.Print "  自動計算: " & r & "行目 金額 = " & _
                            CDbl(qty) * CDbl(price)
            End If
        End If
    End If
End Sub

' ─── パブリックアクセサー (テスト用) ────────────────────────
Public Property Get AuditLog() As String
    AuditLog = m_auditLog
End Property

Public Property Get ChangeCount() As Long
    ChangeCount = m_changeCount
End Property

Public Property Get LastActiveSheet() As String
    LastActiveSheet = m_lastActiveSheet
End Property
