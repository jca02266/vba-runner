VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Option Explicit

' ════════════════════════════════════════════════════════════
' Sheet1 — Excel の Sheet1 (Worksheet) コードビハインドを模倣
'
' 実際の Excel では Worksheet_Change, Worksheet_SelectionChange,
' Worksheet_Activate 等がシートに直接記述できる。
' ここでは Private WithEvents で SimSheet を監視して同等の動作を実現。
'
' 主な構文デモ:
'   - Worksheet_Change: 入力値の自動検証・変換
'   - Worksheet_Activate: ステータスバー更新
'   - Worksheet_Deactivate: 選択状態のクリア
'   - Private ヘルパーによる入力検証ロジック
' ════════════════════════════════════════════════════════════

Private WithEvents m_ws      As SimSheet
Private m_validationErrors   As Long
Private m_statusText         As String
Private m_selectedRow        As Long
Private m_isInitialized      As Boolean

' ─── ワークシート紐付け ──────────────────────────────────────
Public Sub AttachSheet(ByVal ws As SimSheet)
    Set m_ws           = ws
    m_validationErrors = 0
    m_statusText       = ""
    m_selectedRow      = 0
    m_isInitialized    = True
End Sub

' ─── Worksheet_Activate イベント ────────────────────────────
Private Sub m_ws_Activate()
    Debug.Print "--- Sheet1 アクティブ ---"
    UpdateStatus
End Sub

' ─── Worksheet_Deactivate イベント ──────────────────────────
Private Sub m_ws_Deactivate()
    Debug.Print "--- Sheet1 非アクティブ ---"
    m_selectedRow = 0
End Sub

' ─── Worksheet_Change イベント ──────────────────────────────
' 入力値の自動検証・変換 (Excel VBA でよくある実装パターン)
Private Sub m_ws_Change(ByVal Target As SimCell)
    Dim r As Long
    Dim c As Long
    r = Target.Row
    c = Target.Column

    If r <= 1 Then Exit Sub     ' ヘッダー行はスキップ

    Select Case c
    Case 2  ' 商品名列: 大文字化 / 空白チェック
        If Not ValidateName(Target.Value) Then
            Debug.Print "  [警告] " & r & "行目 商品名が空です"
            m_validationErrors = m_validationErrors + 1
        Else
            ' 前後空白を除去
            If CStr(Target.Value) <> Trim(CStr(Target.Value)) Then
                Target.Value = Trim(CStr(Target.Value))
                Debug.Print "  [自動修正] " & r & "行目 商品名の空白を除去"
            End If
        End If

    Case 3  ' 数量列: 正の整数チェック
        If Not ValidateQuantity(Target.Value) Then
            Debug.Print "  [エラー] " & r & "行目 数量は正の整数を入力してください (入力値: " & Target.Value & ")"
            m_validationErrors = m_validationErrors + 1
            Target.Value = 0
        End If

    Case 4  ' 単価列: 正の数値チェック
        If Not ValidatePrice(Target.Value) Then
            Debug.Print "  [エラー] " & r & "行目 単価は正の数値を入力してください (入力値: " & Target.Value & ")"
            m_validationErrors = m_validationErrors + 1
            Target.Value = 0
        End If
    End Select

    m_selectedRow = r
    UpdateStatus
End Sub

' ─── 検証ヘルパー ────────────────────────────────────────────
Private Function ValidateName(ByVal v As Variant) As Boolean
    If VBA.IsEmpty(v) Or CStr(v) = "" Then
        ValidateName = False
    Else
        ValidateName = True
    End If
End Function

Private Function ValidateQuantity(ByVal v As Variant) As Boolean
    If Not VBA.IsNumeric(v) Then
        ValidateQuantity = False
        Exit Function
    End If
    If CDbl(v) < 0 Or CDbl(v) <> CLng(v) Then
        ValidateQuantity = False
        Exit Function
    End If
    ValidateQuantity = True
End Function

Private Function ValidatePrice(ByVal v As Variant) As Boolean
    If Not VBA.IsNumeric(v) Then
        ValidatePrice = False
        Exit Function
    End If
    If CDbl(v) < 0 Then
        ValidatePrice = False
        Exit Function
    End If
    ValidatePrice = True
End Function

' ─── ステータス更新 (Excel のステータスバー更新に相当) ─────────
Private Sub UpdateStatus()
    Dim lastR As Long
    lastR = m_ws.LastRow(2)    ' 商品名列の最終行
    Dim dataRows As Long
    dataRows = IIf(lastR > 1, lastR - 1, 0)
    m_statusText = "データ: " & dataRows & " 件"
    If m_validationErrors > 0 Then
        m_statusText = m_statusText & "  警告: " & m_validationErrors & " 件"
    End If
    Debug.Print "  [ステータス] " & m_statusText
End Sub

' ─── パブリックアクセサー (テスト用) ────────────────────────
Public Property Get ValidationErrors() As Long
    ValidationErrors = m_validationErrors
End Property

Public Property Get StatusText() As String
    StatusText = m_statusText
End Property

Public Property Get SelectedRow() As Long
    SelectedRow = m_selectedRow
End Property
