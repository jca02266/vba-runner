' プロシージャ/クラス未定義エラーテスト (Error 35, 429)

' Error 35: 未定義の関数を引数付きで呼び出す
Sub TestUndefinedFuncWithArgs(assert)
    Dim savedErr As Long
    On Error Resume Next
    Err.Clear
    Dim result
    result = NoSuchFunction(42)
    savedErr = Err.Number
    On Error GoTo 0

    assert.Assert savedErr, 35, "未定義 Function(引数あり) -> Error 35"
End Sub

' Error 35: 未定義の Sub を引数付きで呼び出す
Sub TestUndefinedSubWithArgs(assert)
    Dim savedErr As Long
    On Error Resume Next
    Err.Clear
    Call NoSuchSub(1, 2, 3)
    savedErr = Err.Number
    On Error GoTo 0

    assert.Assert savedErr, 35, "未定義 Sub(引数あり) -> Error 35"
End Sub

' Error 429: 未定義クラスの New
Sub TestNewUnknownClass(assert)
    Dim savedErr As Long
    On Error Resume Next
    Err.Clear
    Dim obj As Object
    Set obj = New NoSuchClass
    savedErr = Err.Number
    On Error GoTo 0

    assert.Assert savedErr, 429, "New 未定義クラス -> Error 429"
End Sub

' Error 429: CreateObject で未サポート ProgID
Sub TestCreateObjectUnknown(assert)
    Dim savedErr As Long
    On Error Resume Next
    Err.Clear
    Dim obj As Object
    Set obj = CreateObject("NoSuch.ProgId")
    savedErr = Err.Number
    On Error GoTo 0

    assert.Assert savedErr, 429, "CreateObject 未サポート ProgID -> Error 429"
End Sub
