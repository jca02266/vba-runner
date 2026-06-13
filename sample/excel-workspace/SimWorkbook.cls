VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SimWorkbook"
Option Explicit

' ════════════════════════════════════════════════════════════
' SimWorkbook — Excel の Workbook をシミュレート
'
' シートの管理と、Workbook レベルのイベントを提供。
' イベント: Open, BeforeClose, SheetChange, SheetActivate
' ════════════════════════════════════════════════════════════

' ─── イベント宣言 ────────────────────────────────────────────
Public Event Open()
Public Event BeforeClose(ByRef Cancel As Boolean)
Public Event SheetChange(ByVal Sh As SimSheet, ByVal Target As SimCell)
Public Event SheetActivate(ByVal Sh As SimSheet)
Public Event SheetDeactivate(ByVal Sh As SimSheet)

' ─── 内部状態 ────────────────────────────────────────────────
Private m_name      As String
Private m_sheets    As Object       ' Scripting.Dictionary: name -> SimSheet
Private m_sheetOrder As Object      ' Scripting.Dictionary: index -> name
Private m_sheetCount As Long
Private m_saved     As Boolean
Private m_path      As String

' ─── 初期化 ─────────────────────────────────────────────────
Private Sub Class_Initialize()
    Set m_sheets     = CreateObject("Scripting.Dictionary")
    Set m_sheetOrder = CreateObject("Scripting.Dictionary")
    m_sheetCount     = 0
    m_saved          = True
    m_path           = ""
End Sub

Private Sub Class_Terminate()
    Set m_sheets     = Nothing
    Set m_sheetOrder = Nothing
End Sub

' ─── ワークブック名 ──────────────────────────────────────────
Public Property Get Name() As String
    Name = m_name
End Property

Public Property Let Name(ByVal n As String)
    m_name = n
End Property

Public Property Get Path() As String
    Path = m_path
End Property

Public Property Let Path(ByVal p As String)
    m_path = p
End Property

Public Property Get FullName() As String
    If m_path <> "" Then
        FullName = m_path & "\" & m_name
    Else
        FullName = m_name
    End If
End Property

' ─── Saved フラグ ────────────────────────────────────────────
Public Property Get Saved() As Boolean
    Saved = m_saved
End Property

Public Property Let Saved(ByVal s As Boolean)
    m_saved = s
End Property

' ─── シート管理 ─────────────────────────────────────────────
Public Function AddSheet(ByVal sheetName As String) As SimSheet
    Dim ws As SimSheet
    Set ws = New SimSheet
    ws.Name = sheetName
    ws.SetParentWorkbook Me    ' Workbook_SheetChange バブリングのため親を設定
    m_sheetCount = m_sheetCount + 1
    m_sheets.Add LCase(sheetName), ws
    m_sheetOrder.Add m_sheetCount, sheetName
    Set AddSheet = ws
End Function

Public Function Sheets(ByVal nameOrIndex As Variant) As SimSheet
    Dim key As String
    If VBA.IsNumeric(nameOrIndex) Then
        Dim idx As Long
        idx = CLng(nameOrIndex)
        If idx < 1 Or idx > m_sheetCount Then
            Err.Raise 9, "SimWorkbook.Sheets", "Index out of range"
        End If
        key = LCase(CStr(m_sheetOrder(idx)))
    Else
        key = LCase(CStr(nameOrIndex))
    End If
    If Not m_sheets.Exists(key) Then
        Err.Raise 9, "SimWorkbook.Sheets", "Sheet not found: " & nameOrIndex
    End If
    Set Sheets = m_sheets(key)
End Function

Public Property Get SheetCount() As Long
    SheetCount = m_sheetCount
End Property

' ─── ワークブックを開く (イベント発火) ──────────────────────
' ※ Open / Close はVBAキーワードのため、メソッド名は OpenWorkbook / CloseWorkbook とする
Public Sub OpenWorkbook()
    RaiseEvent Open
End Sub

' ─── ワークブックを閉じる (イベント発火) ────────────────────
Public Function CloseWorkbook() As Boolean
    Dim cancelFlag As Boolean
    cancelFlag = False
    RaiseEvent BeforeClose(cancelFlag)
    CloseWorkbook = Not cancelFlag
End Function

' ─── セル変更を通知 (Worksheet -> Workbook へのバブリング) ──
Public Sub NotifySheetChange(ByVal ws As SimSheet, ByVal cell As SimCell)
    m_saved = False
    RaiseEvent SheetChange(ws, cell)
End Sub

' ─── シートアクティブ化通知 ─────────────────────────────────
Public Sub NotifySheetActivate(ByVal ws As SimSheet)
    RaiseEvent SheetActivate(ws)
End Sub

Public Sub NotifySheetDeactivate(ByVal ws As SimSheet)
    RaiseEvent SheetDeactivate(ws)
End Sub
