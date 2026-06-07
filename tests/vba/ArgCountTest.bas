Option Explicit

' エラー系テスト: 引数の数が一致しない場合の挙動確認
'
' 注意: 本物のVBAでは引数不一致はコンパイルエラーのため、
' On Error Resume Next でもトラップできず、このファイル自体が実行不能になる。
' このテストはインタープリターの実行時エラー挙動を記録・検証するためのもの。

' ---- テスト用ヘルパープロシージャ ----

Function SumTwo(a, b) As Long
    SumTwo = a + b
End Function

Function NoArgFunc() As Long
    NoArgFunc = 42
End Function

Sub OneArgSub(x)
End Sub

' ---- テストケース ----
' On Error Resume Next が有効な間に assert.IsTrue を呼ぶと、
' IsTrue 内の Err.Raise も飲み込まれてテストが誤ってパスしてしまう。
' そのため、Err.Number を savedErr に退避してから On Error GoTo 0 を呼び、
' エラーハンドリングを無効化した後でアサーションを行う。

' 引数が多すぎる: Function に定義より多い引数を渡す
Sub Test_TooManyArgsToFunction(assert)
    Dim savedErr As Long
    On Error Resume Next
    Err.Clear
    Dim result
    result = SumTwo(1, 2, 3)
    savedErr = Err.Number
    On Error GoTo 0
    assert.IsTrue savedErr <> 0, "引数3個 -> 定義2個: エラーになるべき (Err=" & savedErr & ")"
End Sub

' 引数が少なすぎる: Function に定義より少ない引数を渡す
Sub Test_TooFewArgsToFunction(assert)
    Dim savedErr As Long
    On Error Resume Next
    Err.Clear
    Dim result
    result = SumTwo(1)
    savedErr = Err.Number
    On Error GoTo 0
    assert.IsTrue savedErr <> 0, "引数1個 -> 定義2個: エラーになるべき (Err=" & savedErr & ")"
End Sub

' 引数なし定義の Function に引数を渡す
Sub Test_ArgsToNoArgFunction(assert)
    Dim savedErr As Long
    On Error Resume Next
    Err.Clear
    Dim result
    result = NoArgFunc(1)
    savedErr = Err.Number
    On Error GoTo 0
    assert.IsTrue savedErr <> 0, "引数1個 -> 定義0個: エラーになるべき (Err=" & savedErr & ")"
End Sub

' 引数が多すぎる: Sub に定義より多い引数を渡す
Sub Test_TooManyArgsToSub(assert)
    Dim savedErr As Long
    On Error Resume Next
    Err.Clear
    OneArgSub 1, 2
    savedErr = Err.Number
    On Error GoTo 0
    assert.IsTrue savedErr <> 0, "引数2個 -> 定義1個のSub: エラーになるべき (Err=" & savedErr & ")"
End Sub

' 引数なし呼び出し: 引数1個定義の Sub に引数なしで渡す
Sub Test_ArgsToNoArgSub(assert)
    Dim savedErr As Long
    On Error Resume Next
    Err.Clear
    OneArgSub
    savedErr = Err.Number
    On Error GoTo 0
    assert.IsTrue savedErr <> 0, "引数0個 -> 定義1個のSub: エラーになるべき (Err=" & savedErr & ")"
End Sub
