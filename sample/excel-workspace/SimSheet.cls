VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SimSheet"
Option Explicit

' ════════════════════════════════════════════════════════════
' SimSheet — Excel の Worksheet をシミュレート
'
' セルは Scripting.Dictionary に "R{row}C{col}" キーで格納。
' イベント: Change, Activate, Deactivate
' ════════════════════════════════════════════════════════════

' ─── イベント宣言 ────────────────────────────────────────────
Public Event Change(ByVal Target As SimCell)
Public Event Activate()
Public Event Deactivate()
Public Event Calculate()

' ─── 内部状態 ────────────────────────────────────────────────
Private m_name        As String
Private m_cells       As Object      ' Scripting.Dictionary
Private m_maxRow      As Long
Private m_maxCol      As Long
Private m_isActive    As Boolean
Private m_calcCount   As Long

' ─── 初期化 ─────────────────────────────────────────────────
Private Sub Class_Initialize()
    Set m_cells  = CreateObject("Scripting.Dictionary")
    m_maxRow     = 0
    m_maxCol     = 0
    m_isActive   = False
    m_calcCount  = 0
End Sub

Private Sub Class_Terminate()
    Set m_cells = Nothing
End Sub

' ─── シート名 ───────────────────────────────────────────────
Public Property Get Name() As String
    Name = m_name
End Property

Public Property Let Name(ByVal n As String)
    m_name = n
End Property

' ─── セルアクセス ────────────────────────────────────────────
' Excel の Cells(row, col) 相当
Public Function Cells(ByVal r As Long, ByVal c As Long) As SimCell
    Dim key As String
    key = "R" & r & "C" & c
    If Not m_cells.Exists(key) Then
        Dim cell As SimCell
        Set cell = New SimCell
        cell.Init r, c
        m_cells.Add key, cell
        If r > m_maxRow Then m_maxRow = r
        If c > m_maxCol Then m_maxCol = c
    End If
    Set Cells = m_cells(key)
End Function

' ─── セルへの値書き込み (イベントを発火) ─────────────────────
Public Sub SetValue(ByVal r As Long, ByVal c As Long, ByVal v As Variant)
    Dim cell As SimCell
    Set cell = Me.Cells(r, c)
    cell.Value = v
    RaiseEvent Change(cell)
End Sub

' ─── セル値の読み取り ────────────────────────────────────────
Public Function GetValue(ByVal r As Long, ByVal c As Long) As Variant
    Dim key As String
    key = "R" & r & "C" & c
    If m_cells.Exists(key) Then
        GetValue = m_cells(key).Value
    Else
        GetValue = ""
    End If
End Function

' ─── 範囲の集計 ─────────────────────────────────────────────
' Excel の WorksheetFunction.Sum 相当
Public Function SumRange(ByVal r1 As Long, ByVal c1 As Long, _
                         ByVal r2 As Long, ByVal c2 As Long) As Double
    Dim total As Double
    Dim r As Long, c As Long
    For r = r1 To r2
        For c = c1 To c2
            Dim v As Variant
            v = Me.GetValue(r, c)
            If VBA.IsNumeric(v) Then
                total = total + CDbl(v)
            End If
        Next c
    Next r
    SumRange = total
End Function

' ─── 範囲の件数 (空白以外) ──────────────────────────────────
Public Function CountRange(ByVal r1 As Long, ByVal c1 As Long, _
                           ByVal r2 As Long, ByVal c2 As Long) As Long
    Dim cnt As Long
    Dim r As Long, c As Long
    For r = r1 To r2
        For c = c1 To c2
            Dim v As Variant
            v = Me.GetValue(r, c)
            If Not (VBA.IsEmpty(v) Or CStr(v) = "") Then
                cnt = cnt + 1
            End If
        Next c
    Next r
    CountRange = cnt
End Function

' ─── 検索 (VLookup 相当) ────────────────────────────────────
' 指定列でキー検索し、resultCol 列の値を返す。見つからなければ ""
Public Function VLookup(ByVal key As Variant, _
                        ByVal keyCol As Long, ByVal resultCol As Long, _
                        ByVal r1 As Long, ByVal r2 As Long) As Variant
    Dim r As Long
    For r = r1 To r2
        If CStr(Me.GetValue(r, keyCol)) = CStr(key) Then
            VLookup = Me.GetValue(r, resultCol)
            Exit Function
        End If
    Next r
    VLookup = ""
End Function

' ─── シートの最終行取得 ─────────────────────────────────────
' Excel の Cells(Rows.Count, col).End(xlUp).Row 相当
Public Function LastRow(ByVal col As Long) As Long
    Dim r As Long
    For r = m_maxRow To 1 Step -1
        Dim v As Variant
        v = Me.GetValue(r, col)
        If Not (VBA.IsEmpty(v) Or CStr(v) = "") Then
            LastRow = r
            Exit Function
        End If
    Next r
    LastRow = 0
End Function

' ─── シートをクリア ─────────────────────────────────────────
Public Sub Clear()
    Set m_cells = CreateObject("Scripting.Dictionary")
    m_maxRow = 0
    m_maxCol = 0
End Sub

' ─── アクティブ化 ───────────────────────────────────────────
Public Sub Activate()
    m_isActive = True
    RaiseEvent Activate
End Sub

Public Sub Deactivate()
    m_isActive = False
    RaiseEvent Deactivate
End Sub

Public Property Get IsActive() As Boolean
    IsActive = m_isActive
End Property

' ─── 再計算カウント (テスト用) ──────────────────────────────
Public Sub Recalculate()
    m_calcCount = m_calcCount + 1
    RaiseEvent Calculate
End Sub

Public Property Get CalcCount() As Long
    CalcCount = m_calcCount
End Property

' ─── 範囲をデバッグ出力 ─────────────────────────────────────
Public Sub PrintRange(ByVal r1 As Long, ByVal c1 As Long, _
                      ByVal r2 As Long, ByVal c2 As Long)
    Dim r As Long, c As Long
    For r = r1 To r2
        Dim line As String
        line = ""
        For c = c1 To c2
            If Len(line) > 0 Then line = line & Chr(9)
            line = line & CStr(Me.GetValue(r, c))
        Next c
        Debug.Print line
    Next r
End Sub
