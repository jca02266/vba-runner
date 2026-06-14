VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Option Explicit

' ════════════════════════════════════════════════════════════
' ThisWorkbook — Excel の ThisWorkbook モジュールを模倣
'
' 実際の Excel では Workbook オブジェクトが自動的にバインドされるが、
' ここでは "Private WithEvents Workbook As SimWorkbook" で同等の
' イベントバインドを実現する。
'
' ハンドラー命名規則: {変数名}_{イベント名}
'   Workbook_Open         ← Workbook_Open  (Excel の命名と完全一致)
'   Workbook_BeforeClose  ← Workbook_BeforeClose
'   Workbook_SheetChange  ← Workbook_SheetChange
'   Workbook_SheetActivate← Workbook_SheetActivate
'
' 主な構文デモ:
'   - Private WithEvents Workbook As SimWorkbook
'   - Workbook_Open: 初期化処理 (シート準備・ヘッダー記入)
'   - Workbook_BeforeClose: 未保存確認
'   - Workbook_SheetChange: 変更監視・監査ログ・金額自動計算
'   - Workbook_SheetActivate: アクティブシート切り替え検知
' ════════════════════════════════════════════════════════════

Private WithEvents Workbook As SimWorkbook

Private m_auditLog          As String
Private m_changeCount       As Long
Private m_lastActiveSheet   As String

' ─── ワークブックを紐付け ────────────────────────────────────
Public Sub AttachWorkbook(ByVal wb As SimWorkbook)
    Set Workbook      = wb
    m_auditLog        = ""
    m_changeCount     = 0
    m_lastActiveSheet = ""
End Sub

' ─── Workbook_Open ──────────────────────────────────────────
' Excel の ThisWorkbook モジュールで Workbook_Open を書く際の実装パターン
Private Sub Workbook_WorkbookOpen()
    Debug.Print "=== Workbook_Open: " & Workbook.Name & " ==="

    ' ヘッダー行のセットアップ (Excel でよくある起動時初期化)
    Dim ws As SimSheet
    Set ws = Workbook.Sheets("売上データ")
    ws.Cells(1, 1).Value = "日付"
    ws.Cells(1, 2).Value = "商品名"
    ws.Cells(1, 3).Value = "数量"
    ws.Cells(1, 4).Value = "単価"
    ws.Cells(1, 5).Value = "金額"

    Set ws = Workbook.Sheets("集計")
    ws.Cells(1, 1).Value = "集計日時"
    ws.Cells(1, 2).Value = "総売上"
    ws.Cells(1, 3).Value = "件数"

    Debug.Print "ヘッダー初期化完了"
End Sub

' ─── Workbook_BeforeClose ───────────────────────────────────
' Excel の ThisWorkbook モジュールで Workbook_BeforeClose を書く際の実装パターン
Private Sub Workbook_BeforeClose(ByRef Cancel As Boolean)
    Debug.Print "=== Workbook_BeforeClose ==="
    If Not Workbook.Saved Then
        Debug.Print "未保存の変更あり — 自動保存します"
        Workbook.Saved = True
    End If
    If m_changeCount > 0 Then
        Debug.Print "監査ログ: " & m_changeCount & " 件の変更を記録しました"
    End If
    Cancel = False
End Sub

' ─── Workbook_SheetChange ───────────────────────────────────
' Excel の ThisWorkbook モジュールで Workbook_SheetChange を書く際の実装パターン
' 引数: Sh = 変更されたシート, Target = 変更されたセル
Private Sub Workbook_SheetChange(ByVal Sh As SimSheet, ByVal Target As SimCell)
    m_changeCount = m_changeCount + 1
    Dim entry As String
    entry = "[" & m_changeCount & "] " & Sh.Name & "!" & Target.Address & _
            " = " & CStr(Target.Value)
    m_auditLog = m_auditLog & entry & vbLf
    Debug.Print "変更検知: " & entry

    ' 売上データシートの金額列 (E=5) を自動計算
    ' (数量 or 単価が変わったとき、数量×単価を金額セルに書く)
    If Sh.Name = "売上データ" Then
        Dim c As Long
        c = Target.Column
        If c = 3 Or c = 4 Then          ' 数量 or 単価列
            Dim r As Long
            r = Target.Row
            If r > 1 Then               ' ヘッダー行はスキップ
                Dim qty As Variant
                Dim price As Variant
                qty   = Sh.GetValue(r, 3)
                price = Sh.GetValue(r, 4)
                If VBA.IsNumeric(qty) And VBA.IsNumeric(price) Then
                    Sh.Cells(r, 5).Value = CDbl(qty) * CDbl(price)
                    Debug.Print "  自動計算: " & r & "行目 金額 = " & _
                                CDbl(qty) * CDbl(price)
                End If
            End If
        End If
    End If
End Sub

' ─── Workbook_SheetActivate ─────────────────────────────────
Private Sub Workbook_SheetActivate(ByVal Sh As SimSheet)
    m_lastActiveSheet = Sh.Name
    Debug.Print "シート切り替え → " & Sh.Name
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
