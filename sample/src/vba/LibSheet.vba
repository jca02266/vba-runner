' 設定テーブルから項目名で値を取得する
' tbl: 設定テーブルの ListObject（呼び出し元で取得して渡す）
' itemName: 「項目」列の値
Function GetSetting(tbl As ListObject, itemName As String) As Variant
    Dim colItem As ListColumn
    Dim colValue As ListColumn
    Set colItem = tbl.ListColumns("項目")
    Set colValue = tbl.ListColumns("値")
    Dim i As Long
    For i = 1 To tbl.ListRows.Count
        If colItem.DataBodyRange(i, 1).Value = itemName Then
            GetSetting = colValue.DataBodyRange(i, 1).Value
            Exit Function
        End If
    Next i
    Err.Raise 1004, "GetSetting", "設定項目 '" & itemName & "' が見つかりません"
End Function

Function NewWorksheet(name As String, Optional wb As Workbook = Nothing) As Worksheet
    Dim ws As Worksheet
    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If

    On Error GoTo ErrorHandler
    Set ws = wb.Worksheets(name)
    On Error GoTo 0

    ws.Cells.Delete shift:=xlUp

    Set NewWorksheet = ws
    Exit Function

ErrorHandler:
    Set ws = wb.Worksheets.Add
    ws.Name = name
    Resume Next
End Function
