' Excel API スタブモードの動作確認用

' Application.ScreenUpdating / Calculation / EnableEvents
Function TestAppProperties() As String
    Dim oldSU As Boolean
    oldSU = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Application.Calculation = -4135 ' xlCalculationManual
    Application.EnableEvents = False
    Application.ScreenUpdating = oldSU
    TestAppProperties = "ok"
End Function

' ActiveSheet / Range / Value
Function TestCellReadWrite() As Long
    Dim ws As Object
    Set ws = ActiveSheet
    ws.Cells(1, 1).Value = 42
    TestCellReadWrite = ws.Cells(1, 1).Value
End Function

' ws.Rows.Count / ws.Columns.Count
Function TestRowsColumnsCount() As Long
    Dim ws As Object
    Set ws = ActiveSheet
    Dim r As Long, c As Long
    r = ws.Rows.Count
    c = ws.Columns.Count
    TestRowsColumnsCount = r + c
End Function

' End(xlUp).Row パターン（lastRow 特定）
Function TestEndXlUp() As Long
    Dim ws As Object
    Set ws = ActiveSheet
    ' スタブではデータなし → End(xlUp).Row = 1
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
    TestEndXlUp = lastRow
End Function

' Sheets("name") でシートを取得
Function TestSheets() As String
    Dim ws As Object
    Set ws = Sheets("Data")
    ws.Cells(1, 1).Value = "hello"
    TestSheets = ws.Name
End Function

' A1 の値を読み取るだけ（書き込まない）
Function ReadA1() As Variant
    ReadA1 = ActiveSheet.Cells(1, 1).Value
End Function

' ノーオプが例外を投げないことを確認
Function TestNoOps() As String
    ActiveSheet.Activate
    ActiveSheet.Select
    ActiveSheet.Rows(1).Hidden = True
    ActiveSheet.Columns(1).ColumnWidth = 10
    ActiveSheet.Range("A1:C1").NumberFormat = "@"
    ActiveSheet.Range("A1").Font.Bold = True
    ActiveSheet.Range("A1").Interior.Color = 255
    TestNoOps = "ok"
End Function
