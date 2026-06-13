VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SimCell"
Option Explicit

' ════════════════════════════════════════════════════════════
' SimCell — Excel の Cell (Range の単一セル) をシミュレート
'
' 実際の Excel では Range オブジェクトの 1 つのセルに相当。
' ここでは Value / Address / Row / Column / NumberFormat を提供。
' ════════════════════════════════════════════════════════════

Private m_value        As Variant
Private m_row          As Long
Private m_col          As Long
Private m_numberFormat As String
Private m_formula      As String

' ─── 初期化 ─────────────────────────────────────────────────
Public Sub Init(ByVal r As Long, ByVal c As Long)
    m_row          = r
    m_col          = c
    m_value        = ""
    m_numberFormat = "General"
    m_formula      = ""
End Sub

' ─── Value プロパティ ────────────────────────────────────────
Public Property Get Value() As Variant
    Value = m_value
End Property

Public Property Let Value(ByVal v As Variant)
    m_value  = v
    m_formula = ""              ' 値を直接セットすると数式はクリア
End Property

' ─── Formula プロパティ ─────────────────────────────────────
Public Property Get Formula() As String
    Formula = m_formula
End Property

Public Property Let Formula(ByVal f As String)
    m_formula = f
    ' シミュレーションなので数式の評価は行わず文字列を保持
End Property

' ─── 位置プロパティ ─────────────────────────────────────────
Public Property Get Row() As Long
    Row = m_row
End Property

Public Property Get Column() As Long
    Column = m_col
End Property

' ─── Address (A1 形式) ──────────────────────────────────────
Public Property Get Address() As String
    Dim colLetter As String
    Dim n As Long
    n = m_col
    Do While n > 0
        Dim colRem As Long
        colRem = (n - 1) Mod 26
        colLetter = Chr(65 + colRem) & colLetter
        n = (n - 1) \ 26
    Loop
    Address = "$" & colLetter & "$" & m_row
End Property

' ─── NumberFormat プロパティ ────────────────────────────────
Public Property Get NumberFormat() As String
    NumberFormat = m_numberFormat
End Property

Public Property Let NumberFormat(ByVal fmt As String)
    m_numberFormat = fmt
End Property

' ─── ユーティリティ ──────────────────────────────────────────
Public Function IsBlank() As Boolean
    IsBlank = (VBA.IsEmpty(m_value) Or CStr(m_value) = "")
End Function

Public Function IsNumericValue() As Boolean
    IsNumericValue = VBA.IsNumeric(m_value)
End Function

Public Function Text() As String
    ' Excel の Range.Text 相当: 書式を適用した文字列
    If VBA.IsNumeric(m_value) Then
        Select Case m_numberFormat
        Case "0", "General"
            Text = CStr(CLng(m_value))
        Case "0.00"
            Text = Format(CDbl(m_value), "0.00")
        Case "#,##0"
            Text = Format(CLng(m_value), "#,##0")
        Case "#,##0.00"
            Text = Format(CDbl(m_value), "#,##0.00")
        Case Else
            Text = CStr(m_value)
        End Select
    Else
        Text = CStr(m_value)
    End If
End Function
