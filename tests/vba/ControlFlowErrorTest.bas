' 制御フローエラー番号テスト (Error 3, 13, 52)

' Error 3: Return without GoSub
Sub TestReturnWithoutGoSub(assert)
    Dim savedErr As Long
    On Error Resume Next
    Err.Clear
    Return
    savedErr = Err.Number
    On Error GoTo 0

    assert.Assert savedErr, 3, "Return without GoSub -> Error 3"
End Sub

' Error 13: For Each に非コレクション値を渡す
Sub TestForEachNonCollection(assert)
    Dim savedErr As Long
    On Error Resume Next
    Err.Clear
    Dim i
    Dim x As Long
    x = 123
    For Each i In x
        ' body never runs
    Next i
    savedErr = Err.Number
    On Error GoTo 0

    assert.Assert savedErr, 13, "For Each 非コレクション -> Error 13"
End Sub

' Error 52: オープンしていないファイル番号への Write
Sub TestWriteToClosedFile(assert)
    Dim savedErr As Long
    On Error Resume Next
    Err.Clear
    Print #99, "hello"
    savedErr = Err.Number
    On Error GoTo 0

    assert.Assert savedErr, 52, "未オープンファイルへ Print # -> Error 52"
End Sub

' Error 52: オープンしていないファイル番号への Input
Sub TestInputFromClosedFile(assert)
    Dim savedErr As Long
    On Error Resume Next
    Err.Clear
    Dim s As String
    Input #99, s
    savedErr = Err.Number
    On Error GoTo 0

    assert.Assert savedErr, 52, "未オープンファイルから Input # -> Error 52"
End Sub
