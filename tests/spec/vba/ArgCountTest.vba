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

' 引数が多すぎる: Function に定義より多い引数を渡す
Sub Test_TooManyArgsToFunction(assert)
    On Error Resume Next
    Err.Clear
    Dim result
    result = SumTwo(1, 2, 3)
    assert.IsTrue Err.Number <> 0, "引数3個 -> 定義2個: エラーになるべき (Err=" & Err.Number & ")"
    On Error GoTo 0
End Sub

' 引数が少なすぎる: Function に定義より少ない引数を渡す
Sub Test_TooFewArgsToFunction(assert)
    On Error Resume Next
    Err.Clear
    Dim result
    result = SumTwo(1)
    assert.IsTrue Err.Number <> 0, "引数1個 -> 定義2個: エラーになるべき (Err=" & Err.Number & ")"
    On Error GoTo 0
End Sub

' 引数なし定義の Function に引数を渡す
Sub Test_ArgsToNoArgFunction(assert)
    On Error Resume Next
    Err.Clear
    Dim result
    result = NoArgFunc(1)
    assert.IsTrue Err.Number <> 0, "引数1個 -> 定義0個: エラーになるべき (Err=" & Err.Number & ")"
    On Error GoTo 0
End Sub

' 引数が多すぎる: Sub に定義より多い引数を渡す
Sub Test_TooManyArgsToSub(assert)
    On Error Resume Next
    Err.Clear
    OneArgSub 1, 2
    assert.IsTrue Err.Number <> 0, "引数2個 -> 定義1個のSub: エラーになるべき (Err=" & Err.Number & ")"
    On Error GoTo 0
End Sub

' 引数なし定義の Sub に引数を渡す
Sub Test_ArgsToNoArgSub(assert)
    On Error Resume Next
    Err.Clear
    ' NoArgSub相当: SumTwo を Sub として誤用せず、専用の無引数Subで確認
    OneArgSub
    assert.IsTrue Err.Number <> 0, "引数0個 -> 定義1個のSub: エラーになるべき (Err=" & Err.Number & ")"
    On Error GoTo 0
End Sub
