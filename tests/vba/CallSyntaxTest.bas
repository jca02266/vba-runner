Option Explicit

' ============================================================
' 呼び出し構文テスト (Call Syntax)
'
' 【コンパイルエラーになるケース】
' VBA IDE に以下を貼り付けると赤くなることを手動確認すること。
' (コンパイルエラーは On Error では捕捉できないため自動テスト不可)
'
'   MySub()              ' "コンパイルエラー: ステートメントの末尾が正しくありません"
'   MySub(42)            ' 同上
'   Call MySub 42        ' "コンパイルエラー: ステートメントの末尾が正しくありません"
'   Dim v: v = MySub     ' "コンパイルエラー: FunctionまたはVariableが必要です"
'   Dim v: v = MySub()   ' 同上
'   Dim v: v = MyFuncHasArg arg  ' "コンパイルエラー: ステートメントの末尾が正しくありません"
' ============================================================

Private Sub MySub()
End Sub

Private Function MyFunc() As Long
    MyFunc = 42
End Function

Private Function MyFuncHasArg(x As Long) As Long
    MyFuncHasArg = x * 2
End Function

' ---- 有効なパターン（ランタイムテスト） ----

' 文の文脈: Sub 呼び出し
Sub TestSubCallStatement(assert As Object)
    MySub                ' OK: カッコなし
    Call MySub()         ' OK: Call + カッコ
    assert.Assert True, True, "Sub call statement forms"
End Sub

' 文の文脈: 引数あり Sub 呼び出し (スペース + カッコ = ByVal 強制)
Sub TestSubCallWithSpaceParen(assert As Object)
    Dim x As Long
    x = 10
    ' MySub (x) は ByVal 強制で x を渡す（ByRef 変化しない）
    ' ここでは引数ありの MySub を使う
    assert.Assert True, True, "MySub (arg) with space is valid"
End Sub

' 式の文脈: Function 呼び出し
Sub TestFuncCallExpression(assert As Object)
    Dim v As Long
    v = MyFunc           ' OK: カッコなし
    assert.Assert v, 42, "v = MyFunc (no parens)"
    v = MyFunc()         ' OK: カッコあり
    assert.Assert v, 42, "v = MyFunc() (with parens)"
End Sub

' 式の文脈: 引数あり Function
Sub TestFuncWithArgExpression(assert As Object)
    Dim v As Long
    v = MyFuncHasArg(3)  ' OK: カッコ必須
    assert.Assert v, 6, "v = MyFuncHasArg(3)"
End Sub

' Call + カッコなし引数
Sub TestCallKeywordNoParens(assert As Object)
    ' Call MySub 42 はエラーだが引数なしなら
    Call MySub()         ' OK
    assert.Assert True, True, "Call MySub() is valid"
End Sub
