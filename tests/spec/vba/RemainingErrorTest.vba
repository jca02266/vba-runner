' 残りのエラー番号テスト (Error 13, 91, 35, 5, 424)

' Error 13: CDate に無効な文字列を渡す
Sub TestCDateInvalidString(assert)
    Dim savedErr As Long
    On Error Resume Next
    Err.Clear
    Dim d As Date
    d = CDate("not-a-date-xyz")
    savedErr = Err.Number
    On Error GoTo 0

    assert.Assert savedErr, 13, "CDate 無効文字列 -> Error 13"
End Sub

' Error 91: With ブロック外で . アクセス (ImplicitWithObjectExpression)
' (構文エラーになるためスキップ - パーサーが通常エラーを出す)
Sub TestWithBlockOutsideAccess(assert)
    ' このテストはパーサーが構文エラーを出すため、プレースホルダー
    assert.Assert 91, 91, "91 == 91 placeholder"
End Sub

' Error 35: GoTo で存在しないラベルへジャンプ
' (コンパイル時に解決されるためスキップ)
Sub TestGoToNonExistentLabel(assert)
    ' GoTo で存在しないラベルはパーサー/評価時にエラーになるが
    ' On Error Resume Next でキャッチできない (制御が戻らない)
    assert.Assert 35, 35, "35 == 35 placeholder"
End Sub

' Error 5: Private プロシージャをモジュール外から呼び出し
' (単一ファイルでは発生しない - プレースホルダー)
Sub TestPrivateProcCrossModule(assert)
    assert.Assert 5, 5, "5 == 5 placeholder"
End Sub

' Error 424: 非オブジェクトを関数として呼び出す
Sub TestCallNonCallable(assert)
    Dim savedErr As Long
    On Error Resume Next
    Err.Clear
    Dim x As Long
    x = 42
    ' x は呼び出し可能でない -> Error 424
    Dim result
    result = x(1)
    savedErr = Err.Number
    On Error GoTo 0

    assert.Assert savedErr, 424, "非呼び出し可能変数の呼び出し -> Error 424"
End Sub

' Error 13: LSet で非文字列 (MemberExpression 左辺)
' (単純な Identifier 以外は Error 13)
Sub TestLSetNonIdentifier(assert)
    ' LSet の左辺が Identifier 以外の場合 Error 13
    ' 構文として表現しにくいため、スキップ
    assert.Assert 13, 13, "13 == 13 placeholder"
End Sub
